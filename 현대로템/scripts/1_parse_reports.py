# -*- coding: utf-8 -*-
import os, re, json
from html.parser import HTMLParser

SCRATCH = "/tmp/claude-0/-home-user-jooho96/2e90a833-c170-58f0-a6f1-86983b335b93/scratchpad/reports"

PERIODS = [
    ("반기보고서 (2021.06)_20210817", "2021Q2", 20210630),
    ("분기보고서 (2021.09)_20211115", "2021Q3", 20210930),
    ("사업보고서 (2021.12)_20220315", "2021Q4", 20211231),
    ("분기보고서 (2022.03)_20220516", "2022Q1", 20220331),
    ("반기보고서 (2022.06)_20220816", "2022Q2", 20220630),
    ("분기보고서 (2022.09)_20221114", "2022Q3", 20220930),
    ("사업보고서 (2022.12)_20230310", "2022Q4", 20221231),
    ("분기보고서 (2023.03)_20230515", "2023Q1", 20230331),
    ("반기보고서 (2023.06)_20230814", "2023Q2", 20230630),
    ("분기보고서 (2023.09)_20231114", "2023Q3", 20230930),
    ("사업보고서 (2023.12)_20240320", "2023Q4", 20231231),
    ("분기보고서 (2024.03)_20240516", "2024Q1", 20240331),
    ("반기보고서 (2024.06)_20240814", "2024Q2", 20240630),
    ("분기보고서 (2024.09)_20241114", "2024Q3", 20240930),
    ("[기재정정]사업보고서 (2024.12)_20250321", "2024Q4", 20241231),
    ("분기보고서 (2025.03)_20250515", "2025Q1", 20250331),
    ("반기보고서 (2025.06)_20250814", "2025Q2", 20250630),
    ("분기보고서 (2025.09)_20251114", "2025Q3", 20250930),
    ("사업보고서 (2025.12)_20260319", "2025Q4", 20251231),
    ("분기보고서 (2026.03)_20260515", "2026Q1", 20260331),
]


def readtext(folder):
    fp = os.path.join(SCRATCH, folder)
    mains = [f for f in os.listdir(fp) if re.match(r"^\d+\.xml$", f)]
    if not mains:
        mains = [f for f in os.listdir(fp) if f.endswith(".xml")]
    p = os.path.join(fp, mains[0])
    b = open(p, "rb").read()
    for enc in ("utf-8", "cp949"):
        try:
            return b.decode(enc), mains[0]
        except Exception:
            pass
    return b.decode("utf-8", "replace"), mains[0]


class TableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.in_td = False; self.cur_text = []; self.cur_span = (1, 1)
        self.rows = []; self.cur_row = None

    def handle_starttag(self, tag, attrs):
        t = tag.upper(); a = {k.upper(): v for k, v in attrs}
        if t == "TR": self.cur_row = []
        elif t in ("TD", "TH"):
            self.in_td = True; self.cur_text = []
            self.cur_span = (int(a.get("ROWSPAN", "1") or 1), int(a.get("COLSPAN", "1") or 1))
        elif t == "BR" and self.in_td: self.cur_text.append(" ")

    def handle_endtag(self, tag):
        t = tag.upper()
        if t in ("TD", "TH") and self.in_td:
            txt = re.sub(r"\s+", " ", "".join(self.cur_text)).strip()
            self.cur_row.append((txt, self.cur_span[0], self.cur_span[1])); self.in_td = False
        elif t == "TR" and self.cur_row is not None:
            self.rows.append(self.cur_row); self.cur_row = None

    def handle_data(self, data):
        if self.in_td: self.cur_text.append(data)


def build_grid(rows):
    grid = []; carry = {}
    for src in rows:
        out = []; col = 0; srci = 0
        while srci < len(src) or col in carry:
            if col in carry:
                rem, txt = carry[col]; out.append((txt, False)); rem -= 1
                if rem <= 0: del carry[col]
                else: carry[col] = [rem, txt]
                col += 1; continue
            if srci < len(src):
                txt, rs, cs = src[srci]; srci += 1
                for _ in range(cs):
                    out.append((txt, True))
                    if rs > 1: carry[col] = [rs - 1, txt]
                    col += 1
            else: break
        grid.append(out)
    return grid


def find_table(txt):
    # '단일판매'가 나오고 그 뒤 1500자 내 '계약명'이 있으면 실제 표. '기재하지 않' 이면 스킵.
    for m in re.finditer("단일판매", txt):
        window = txt[m.start(): m.start() + 2000]
        if "기재하지 않" in window[:200]:
            continue
        if "계약명" not in window:
            continue
        # 이 지점부터 각주/다음섹션 전까지
        end_m = re.search(r"계약가는.{0,15}단일판매|2\.\s*우발부채", txt[m.start():])
        end = m.start() + (end_m.start() if end_m else 40000)
        seg = txt[m.start():end]
        tables = re.findall(r"<TABLE\b.*?</TABLE>", seg, re.S | re.I)
        for tb in tables:
            if "계약명" in tb and "누적" in tb:
                return tb
    return None


def parse_rows(grid):
    hdr = None
    for i, row in enumerate(grid):
        if "계약명" in [c[0] for c in row]:
            hdr = i; break
    if hdr is None: return []
    out = []
    for row in grid[hdr + 1:]:
        texts = [c[0] for c in row]
        if "당기" in texts and "누적" in texts and "계약금액" in texts:
            continue
        if len(row) < 8:
            continue
        g = lambda k: row[k][0] if k < len(row) else ""
        orig = lambda k: row[k][1] if k < len(row) else True
        name = g(1)
        if not name or name == "계약명":
            continue
        out.append({
            "신고일자": g(0), "계약명": name, "계약상대방": g(2),
            "계약시작일": g(3), "계약종료일": g(4), "계약금액": g(5),
            "매출_당기": g(6), "매출_누적": g(7), "대금_당기": g(8), "대금_누적": g(9),
            "금액공유": (not orig(7)),
        })
    return out


def main():
    result = {}
    for folder, lab, key in PERIODS:
        txt, src = readtext(folder)
        tb = find_table(txt)
        rows = parse_rows(_feed(tb)) if tb else []
        result[lab] = {"key": key, "src": src, "n": len(rows), "rows": rows}
        print(f"{lab}: {len(rows):2d} rows  src={src}")
    out = "/tmp/claude-0/-home-user-jooho96/2e90a833-c170-58f0-a6f1-86983b335b93/scratchpad/parsed2.json"
    json.dump(result, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("saved", out)


def _feed(tb):
    p = TableParser(); p.feed(tb); return build_grid(p.rows)


if __name__ == "__main__":
    main()
