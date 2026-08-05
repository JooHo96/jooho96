# -*- coding: utf-8 -*-
import re, json, parse2

PERIODS=[("반기보고서 (2021.06)_20210817","2021Q2","반기"),("분기보고서 (2021.09)_20211115","2021Q3","분기"),
 ("사업보고서 (2021.12)_20220315","2021Q4","사업"),("분기보고서 (2022.03)_20220516","2022Q1","분기"),
 ("반기보고서 (2022.06)_20220816","2022Q2","반기"),("분기보고서 (2022.09)_20221114","2022Q3","분기"),
 ("사업보고서 (2022.12)_20230310","2022Q4","사업"),("분기보고서 (2023.03)_20230515","2023Q1","분기"),
 ("반기보고서 (2023.06)_20230814","2023Q2","반기"),("분기보고서 (2023.09)_20231114","2023Q3","분기"),
 ("사업보고서 (2023.12)_20240320","2023Q4","사업"),("분기보고서 (2024.03)_20240516","2024Q1","분기"),
 ("반기보고서 (2024.06)_20240814","2024Q2","반기"),("분기보고서 (2024.09)_20241114","2024Q3","분기"),
 ("[기재정정]사업보고서 (2024.12)_20250321","2024Q4","사업"),("분기보고서 (2025.03)_20250515","2025Q1","분기"),
 ("반기보고서 (2025.06)_20250814","2025Q2","반기"),("분기보고서 (2025.09)_20251114","2025Q3","분기"),
 ("사업보고서 (2025.12)_20260319","2025Q4","사업"),("분기보고서 (2026.03)_20260515","2026Q1","분기")]

def rt(f): return parse2.readtext(f)[0]
def strip(s): return re.sub(r"\s+"," ",re.sub(r"<[^>]+>"," ",s)).strip()
def N(s):
    s=(s or "").strip()
    if s in ("-","","–"): return None
    neg=s.startswith("(")
    s2=re.sub(r"[(),\s]","",s)
    try: v=float(s2); return -v if neg else v
    except: return None

def get_section(t):
    for a in [r"10%\s*이상을\s*차지하는\s*외부\s*고객", r"주요\s*고객에\s*대한\s*공시"]:
        m=re.search(a,t)
        if m: return m.start(),a
    return None,None

def parse_tall(seg):
    """구형식: 각 고객행에 부문. TableParser로."""
    tbs=re.findall(r"<TABLE\b.*?</TABLE>", seg, re.S|re.I)
    for tb in tbs:
        if "고객1" not in tb or "부문" not in tb: continue
        grid=parse2._feed(tb)
        custs=[]
        for row in grid:
            cells=[c[0] for c in row]
            if cells and re.match(r"고객\s*\d", cells[0]):
                seg_lbl=next((x for x in cells if "방산" in x or "철도" in x), "")
                nums=[N(x) for x in cells[1:] if N(x) is not None]
                custs.append({"label":re.sub(r"\s","",cells[0]),"seg":"방산" if "방산" in seg_lbl else ("철도" if "철도" in seg_lbl else ""),
                              "cur_3m":nums[0] if len(nums)>0 else None,
                              "cur_cum":nums[1] if len(nums)>1 else None})
        if custs: return custs
    return []

def parse_wide(seg, ptype):
    txt=strip(seg)
    # 현재기간 헤더 위치
    ncust=0
    for i in range(1,6):
        if f"고객{i}" in txt or f"고객 {i}" in txt: ncust=i
    has_pair = ("3개월" in txt and "누적" in txt) and ptype!="사업"
    # 첫 '수익(매출액)' 뒤 숫자들
    m=re.search(r"수익\(?매출액\)?", txt)
    custs=[]
    if not m: return custs, ncust, has_pair
    after=txt[m.end(): m.end()+300]
    nums=re.findall(r"[\d,]{4,}", after)
    nums=[N(x) for x in nums]
    per = 2 if has_pair else 1
    for i in range(ncust):
        base=i*per
        if base>=len(nums): break
        if has_pair:
            custs.append({"label":f"고객{i+1}","seg":"","cur_3m":nums[base],"cur_cum":nums[base+1] if base+1<len(nums) else None})
        else:
            custs.append({"label":f"고객{i+1}","seg":"","cur_3m":None,"cur_cum":nums[base]})
    return custs, ncust, has_pair

out={}
for folder,lab,ptype in PERIODS:
    t=rt(folder); start,anchor=get_section(t)
    seg=t[start:start+6000]
    tall="부문" in strip(seg)[:400] and "고객1" in seg and re.search(r"<TABLE[^>]*>.*?부문.*?</TABLE>",seg,re.S)
    # 구형식 판단: 세로표에 부문 열
    custs=parse_tall(seg)
    fmt="tall"
    if not custs:
        custs,nc,hp=parse_wide(seg,ptype); fmt="wide"
    # 각주
    foot=" ".join(re.findall(r"\(\*\d?\)[^<(]{5,80}", strip(seg)))
    out[lab]={"ptype":ptype,"fmt":fmt,"custs":custs,"foot":foot[:200]}
    print(f"\n### {lab} ({ptype}/{fmt})")
    for c in custs:
        s3=f"{int(c['cur_3m']):,}" if c['cur_3m'] else "-"
        sc=f"{int(c['cur_cum']):,}" if c['cur_cum'] else "-"
        print(f"   {c['label']:8} seg={c['seg'] or '?':4}  3개월={s3:>15}  누적/당기={sc:>15}")
json.dump(out, open("customers2.json","w",encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
print("\nsaved customers2.json")
