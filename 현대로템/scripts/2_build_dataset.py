# -*- coding: utf-8 -*-
import re, json, sys
sys.path.insert(0, "/tmp/claude-0/-home-user-jooho96/2e90a833-c170-58f0-a6f1-86983b335b93/scratchpad")
import parse2

REPORTS = [
    ("2024Q4", "[기재정정]사업보고서 (2024.12)_20250321", "A"),   # 당기=FY2024
    ("2025Q2", "반기보고서 (2025.06)_20250814", "A"),             # 당기=2025 H1
    ("2025Q3", "분기보고서 (2025.09)_20251114", "A"),             # 당기=2025 9M YTD
    ("2025Q4", "사업보고서 (2025.12)_20260319", "A"),             # 당기=FY2025
    ("2026Q1", "분기보고서 (2026.03)_20260515", "B"),             # 당기=2026 Q1 YTD(3M)
]
LABS = [r[0] for r in REPORTS]

def num(s):
    if s is None: return None
    s = re.sub(r"^(KRW|USD|EUR|JPY|GBP|AUD|SGD|TWD|NZD|CAD)\s*", "", s.strip())
    if s in ("-", "", "–", "—"): return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s2 = re.sub(r"[(),\s]", "", s).replace("△", "-").replace("▲", "-")
    try:
        v = float(s2); return -v if neg else v
    except ValueError:
        return None

def first_date(s):
    m = re.search(r"\d{4}-\d{2}-\d{2}", s or "")
    return m.group(0) if m else (s or "").strip()[:10]

def extract_A(grid):
    """origin flag 포함. 반환: rows(dict)."""
    hdr = next((i for i,r in enumerate(grid) if "계약명" in [c[0] for c in r]), None)
    if hdr is None: return []
    rows = []
    for row in grid[hdr+1:]:
        texts=[c[0] for c in row]
        if "당기" in texts and "누적" in texts and "계약금액" in texts: continue
        if len(row) < 8: continue
        g=lambda k: row[k][0] if k<len(row) else ""
        og=lambda k: row[k][1] if k<len(row) else True
        name=g(1)
        if not name or name=="계약명": continue
        rows.append(dict(sinbo=g(0), name=name, party=g(2), start=g(3), end=g(4),
                         amt=g(5), curr=g(6), cum=g(7),
                         shared=(not og(7))))  # 누적 셀 상속 → 묶음 follower
    return rows

def extract_B(grid):
    rows=[]
    hdr = next((i for i,r in enumerate(grid) if "누적" in " ".join(c[0] for c in r) and "계약금액" in " ".join(c[0] for c in r)), 0)
    for row in grid[hdr+1:]:
        if len(row) < 5: continue
        merged=row[1][0]
        if "계약명" not in merged and "계약기간" not in merged:
            # 상속으로 병합셀이 비었을 수 있음
            pass
        mn=re.search(r"계약명\s*:\s*(.*?)(?:계약상대방\s*:|계약기간\s*:|$)", merged)
        mp=re.search(r"계약상대방\s*:\s*(.*?)(?:계약기간\s*:|$)", merged)
        mk=re.search(r"계약기간\s*:\s*(.*)$", merged)
        name=(mn.group(1) if mn else merged).strip()
        party=(mp.group(1) if mp else "").strip()
        period=(mk.group(1) if mk else "")
        pm=re.findall(r"\d{4}-\d{2}-\d{2}", period)
        start=pm[0] if pm else ""
        end=pm[1] if len(pm)>1 else ("-" if "~" in period and len(pm)==1 else "")
        if not name: continue
        # origin flag: 누적셀(col4) 상속 여부
        shared = (len(row)>4 and not row[4][1])
        rows.append(dict(sinbo=row[0][0], name=name, party=party, start=start, end=end,
                         amt=row[2][0] if len(row)>2 else "", curr=row[3][0] if len(row)>3 else "",
                         cum=row[4][0] if len(row)>4 else "", shared=shared))
    return rows

def collapse(rows):
    """rowspan 묶음 병합: shared=True 행을 직전 owner에 흡수."""
    out=[]
    for r in rows:
        if r["shared"] and out:
            out[-1]["subnames"].append(r["name"])
            # owner에 하위 계약명 추가 (금액은 공유이므로 owner 것 유지)
            continue
        r=dict(r); r["subnames"]=[]
        out.append(r)
    return out

# ---- 수집 ----
per_report={}
for lab, folder, schema in REPORTS:
    txt,src=parse2.readtext(folder)
    grid=parse2._feed(parse2.find_table(txt))
    rows = extract_A(grid) if schema=="A" else extract_B(grid)
    rows = collapse(rows)
    per_report[lab]=rows
    print(f"{lab}: {len(rows)} contracts (after collapse)  src={src}")

# ---- 마스터: 키 = 신고일자(첫날짜) + 계약시작일 ----
master={}
for lab in LABS:
    for r in per_report[lab]:
        key=first_date(r["sinbo"])+"|"+(r["start"] or "")
        m=master.setdefault(key, dict(sinbo=first_date(r["sinbo"]), names=[], party=[], start=r["start"], end="",
                                      amt={}, curr={}, cum={}, subnames=set()))
        if r["name"] not in m["names"]: m["names"].append(r["name"])
        if r["party"] and r["party"] not in m["party"]: m["party"].append(r["party"])
        if r["end"] and r["end"]!="-": m["end"]=r["end"]
        for sn in r["subnames"]: m["subnames"].add(sn)
        a=num(r["amt"]);  m["amt"][lab]=(r["amt"], a)
        m["curr"][lab]=num(r["curr"])
        m["cum"][lab]=num(r["cum"])

print(f"\n마스터 계약 수: {len(master)}")

# ---- 분기 매출인식 계산 (당기 YTD 차분) ----
def q_calc(curr):
    g=lambda l: curr.get(l)
    def sub(a,b):
        if a is None: return None
        return a-(b or 0)
    return {
        "2024년": g("2024Q4"),
        "2025_상반기": g("2025Q2"),
        "2025_3Q": sub(g("2025Q3"), g("2025Q2")),
        "2025_4Q": sub(g("2025Q4"), g("2025Q3")),
        "2026_1Q": g("2026Q1"),
    }

for key,m in master.items():
    m["quarters"]=q_calc(m["curr"])
    # 계약금액: Schema A(원화 백만) 우선 2025Q4>2025Q3>2025Q2>2024Q4, 없으면 2026Q1(외화)
    amt=None; amt_raw=""
    for l in ["2025Q4","2025Q3","2025Q2","2024Q4"]:
        if l in m["amt"] and m["amt"][l][1] is not None and not re.match(r"^(USD|EUR|JPY|GBP|AUD|SGD|TWD|NZD|CAD)", m["amt"][l][0]):
            amt=m["amt"][l][1]; amt_raw=m["amt"][l][0]; break
    if amt is None and "2026Q1" in m["amt"]:
        amt_raw=m["amt"]["2026Q1"][0]; amt=m["amt"]["2026Q1"][1]
    m["amt_krw"]=amt; m["amt_raw"]=amt_raw
    # 최신 누적
    latest=None
    for l in reversed(LABS):
        if m["cum"].get(l) is not None: latest=m["cum"][l]; break
    m["cum_latest"]=latest

json.dump(master, open("dataset.json","w",encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
print("saved dataset.json")

# ---- 검증 출력 ----
rowsout=[]
for key,m in master.items():
    name=sorted(m["names"], key=len)[-1]
    q=m["quarters"]
    rowsout.append((m.get("cum_latest") or 0, name, m["party"][0] if m["party"] else "",
                    m["amt_krw"], m["cum_latest"], q))
rowsout.sort(key=lambda x:-x[0])
print(f"\n{'계약명':40}{'계약금액':>10}{'최신누적':>10} | {'FY24':>8}{'25상반':>8}{'25_3Q':>8}{'25_4Q':>8}{'26_1Q':>8}")
for _,name,party,amt,cum,q in rowsout:
    f=lambda v: f"{int(v):>8,}" if isinstance(v,(int,float)) else f"{'-':>8}"
    print(f"{name[:39]:40}{f(amt)}{f(cum)} | {f(q['2024년'])}{f(q['2025_상반기'])}{f(q['2025_3Q'])}{f(q['2025_4Q'])}{f(q['2026_1Q'])}")
