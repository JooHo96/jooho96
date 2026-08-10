"""
HD현대중공업 단일판매·공급계약 공시 수집 → CSV + 엑셀(뉴스수주 양식)

사용법:
    pip install requests openpyxl

    # CSV만 출력
    python hd_heavy_orders.py --api-key YOUR_DART_KEY

    # 엑셀 양식에 직접 입력
    python hd_heavy_orders.py --api-key YOUR_DART_KEY --excel "수주현황.xlsx"

    # 기간 지정 (기본 6년)
    python hd_heavy_orders.py --api-key YOUR_DART_KEY --years 3 --excel "수주현황.xlsx"
"""

import argparse
import calendar
import csv
import io
import re
import time
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

import requests
from openpyxl import load_workbook

CORP_CODE  = "01390344"   # HD현대중공업
CORP_NAME  = "HD현대중공업"
DART_BASE  = "https://opendart.fss.or.kr/api"
SHEET_NAME = "뉴스수주"

# 컬럼 인덱스 (1-based) — 업로드 양식 기준
COL = {
    "A":  1,   # =AC{row}
    "B":  2,   # 회사
    "C":  3,   # 날짜
    "D":  4,   # 출처
    "E":  5,   # 기타
    "F":  6,   # Type(선종)
    "G":  7,   # Name(선주)
    "H":  8,   # Size
    "I":  9,   # Unit(척수)
    "N":  14,  # Built(인도년)
    "O":  15,  # Month(인도월)
    "P":  16,  # Builder
    "Q":  17,  # Contract Date
    "Z":  26,  # Built date
    "AA": 27,  # 총계약금액(백만달러)
    "AB": 28,  # 총계약금액(십억원)
    "AC": 29,  # 척당금액(백만달러)
    "AD": 30,  # 기준환율
    "AE": 31,  # 확정여부
    "AF": 32,  # 날짜 수식
    "AG": 33,  # 상선/특수선
    "AH": 34,  # Clarksons 반영여부
}


# ── DART API ──────────────────────────────────────────────────────────────

def get_list(api_key, bgn_de, end_de):
    results, page = [], 1
    while True:
        try:
            r = requests.get(f"{DART_BASE}/list.json", params={
                "crtfc_key": api_key, "corp_code": CORP_CODE,
                "pblntf_detail_ty": "B002",
                "bgn_de": bgn_de, "end_de": end_de,
                "page_count": 100, "page_no": page,
            }, timeout=15).json()
        except Exception as e:
            print(f"    [오류] {e}")
            break
        if r.get("status") not in ("000",):
            break
        results.extend(r.get("list", []))
        if page * 100 >= int(r.get("total_count", 0)):
            break
        page += 1
        time.sleep(0.3)
    return results


def get_xml(api_key, rcept_no):
    try:
        r = requests.get(f"{DART_BASE}/document.xml",
                         params={"crtfc_key": api_key, "rcept_no": rcept_no}, timeout=20)
        if r.content[:2] == b"PK":
            with zipfile.ZipFile(io.BytesIO(r.content)) as z:
                for n in z.namelist():
                    if n.endswith(".xml"):
                        return z.read(n).decode("utf-8", errors="ignore")
        return r.text
    except Exception as e:
        print(f"    [오류] XML 다운로드: {e}")
        return None


# ── 파싱 ──────────────────────────────────────────────────────────────────

def find(xml, *keys):
    for k in keys:
        m = re.search(rf"{re.escape(k)}[^<]*</[^>]+>\s*<[^>]+>([^<]+)</", xml)
        if m:
            v = re.sub(r"\s+", " ", m.group(1)).strip()
            if v and v not in ("-", "해당없음", "N/A"):
                return v
    return ""


def parse_krw_to_billion(text):
    """원화 금액 → 십억원"""
    text = re.sub(r"[,\s]", "", text)
    m = re.search(r"[\d.]+", text)
    if not m:
        return None
    v = float(m.group())
    if "조" in text:
        return round(v * 1000, 2)
    elif "억" in text:
        return round(v / 10, 2)
    elif "백만원" in text:
        return round(v / 1000, 2)
    return None


def parse_usd_to_million(text):
    """달러 금액 → 백만달러"""
    text = re.sub(r"[,\s]", "", text)
    m = re.search(r"[\d.]+", text)
    if not m:
        return None
    v = float(m.group())
    if "억달러" in text or "억USD" in text.upper():
        return round(v * 100, 2)
    elif "백만달러" in text or "백만USD" in text.upper():
        return round(v, 2)
    elif "천만달러" in text:
        return round(v * 10, 2)
    return None


def parse(xml, rcept_dt, report_nm):
    row = {
        "date": None, "contract_date": None,
        "buyer": "", "vessel_type": "", "vessel_category": "상선",
        "size": "", "quantity": None,
        "delivery_year": None, "delivery_month": None, "delivery_date": None,
        "amount_usd_mil": None, "amount_krw_bil": None,
        "unit_price_usd": None, "exchange_rate": None,
        "confirmed": "O", "clarksons": "",
        "etc": report_nm,
    }

    if len(rcept_dt) == 8:
        try:
            dt = datetime(int(rcept_dt[:4]), int(rcept_dt[4:6]), int(rcept_dt[6:]))
            row["date"] = dt
            row["contract_date"] = dt
        except Exception:
            pass

    row["buyer"] = find(xml, "계약상대방", "거래상대방", "발주처", "매수인", "수요자")

    amt = find(xml, "계약금액", "총계약금액", "공급금액", "거래금액")
    if amt:
        if "달러" in amt or "USD" in amt.upper():
            row["amount_usd_mil"] = parse_usd_to_million(amt)
        else:
            row["amount_krw_bil"] = parse_krw_to_billion(amt)

    purpose = find(xml, "계약내용", "공급내용", "계약목적물", "품목", "선박종류")
    stx = purpose + xml[:5000]
    for kw, label in [
        ("컨테이너", "컨테이너선"), ("LNG", "LNG선"), ("LPG", "LPG선"),
        ("암모니아", "LPG선"), ("VLCC", "VLCC"), ("원유운반", "원유운반선"),
        ("PC선", "P/C선"), ("제품운반", "P/C선"), ("MR탱커", "P/C선"),
        ("벌크", "벌크선"), ("살물선", "벌크선"),
        ("FPSO", "해양"), ("풍력", "해양"), ("해양플랜트", "해양"),
        ("특수선", "특수선"), ("군함", "특수선"), ("잠수함", "특수선"),
    ]:
        if kw.upper() in stx.upper():
            row["vessel_type"] = label
            break
    if row["vessel_type"] in ("해양", "특수선"):
        row["vessel_category"] = row["vessel_type"]

    qty = find(xml, "수량", "척수", "호선수", "선박수")
    if qty:
        m = re.search(r"\d+", qty)
        if m:
            row["quantity"] = int(m.group())

    dlv = find(xml, "납기", "인도예정", "납품예정", "인도일", "공급일정")
    if dlv:
        m = re.search(r"(\d{4})[.\-년]\s*(\d{1,2})", dlv)
        if m:
            y, mo = int(m.group(1)), int(m.group(2))
            row["delivery_year"] = y
            row["delivery_month"] = mo
            try:
                last_day = calendar.monthrange(y, mo)[1]
                row["delivery_date"] = datetime(y, mo, last_day)
            except Exception:
                pass

    return row


# ── 엑셀 쓰기 ─────────────────────────────────────────────────────────────

def find_next_empty_row(ws):
    for r in range(3, ws.max_row + 2):
        if ws.cell(row=r, column=COL["B"]).value is None:
            return r
    return ws.max_row + 1


def already_exists(ws, contract_date):
    if contract_date is None:
        return False
    for r in range(3, ws.max_row + 1):
        b = ws.cell(row=r, column=COL["B"]).value
        q = ws.cell(row=r, column=COL["Q"]).value
        if b == CORP_NAME and isinstance(q, datetime) and q.date() == contract_date.date():
            return True
    return False


def write_excel_row(ws, row_num, data):
    r = row_num
    ws.cell(row=r, column=COL["A"]).value  = f"=AC{r}"
    ws.cell(row=r, column=COL["B"]).value  = CORP_NAME
    ws.cell(row=r, column=COL["C"]).value  = data["date"]
    ws.cell(row=r, column=COL["D"]).value  = "DART"
    ws.cell(row=r, column=COL["E"]).value  = data["etc"] or None
    ws.cell(row=r, column=COL["F"]).value  = data["vessel_type"] or None
    ws.cell(row=r, column=COL["G"]).value  = data["buyer"] or None
    ws.cell(row=r, column=COL["H"]).value  = data["size"] or None
    ws.cell(row=r, column=COL["I"]).value  = data["quantity"]
    ws.cell(row=r, column=COL["N"]).value  = data["delivery_year"]
    ws.cell(row=r, column=COL["O"]).value  = data["delivery_month"]
    ws.cell(row=r, column=COL["P"]).value  = CORP_NAME
    ws.cell(row=r, column=COL["Q"]).value  = data["contract_date"]
    ws.cell(row=r, column=COL["Z"]).value  = data["delivery_date"]
    ws.cell(row=r, column=COL["AA"]).value = data["amount_usd_mil"]

    if data["amount_usd_mil"] and data["exchange_rate"]:
        ws.cell(row=r, column=COL["AB"]).value = f"=AA{r}/AD{r}*1000"
    elif data["amount_krw_bil"]:
        ws.cell(row=r, column=COL["AB"]).value = data["amount_krw_bil"]

    if data["quantity"] and data["amount_usd_mil"]:
        ws.cell(row=r, column=COL["AC"]).value = f"=AA{r}/I{r}"
    elif data["unit_price_usd"]:
        ws.cell(row=r, column=COL["AC"]).value = data["unit_price_usd"]

    ws.cell(row=r, column=COL["AD"]).value = data["exchange_rate"]
    ws.cell(row=r, column=COL["AE"]).value = data["confirmed"]
    ws.cell(row=r, column=COL["AF"]).value = f'=YEAR(Q{r})&"."&MONTH(Q{r})'
    ws.cell(row=r, column=COL["AG"]).value = data["vessel_category"]
    ws.cell(row=r, column=COL["AH"]).value = data["clarksons"] or None


# ── CSV 저장 ──────────────────────────────────────────────────────────────

def save_csv(records, out_path):
    fields = ["회사", "날짜", "공시제목", "선주", "선종", "척수",
              "인도년", "인도월", "금액(원화,십억원)", "금액(달러,백만)", "확정", "상선특수선"]
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for d in records:
            dt = d["contract_date"].strftime("%Y-%m-%d") if d["contract_date"] else ""
            w.writerow({
                "회사":             CORP_NAME,
                "날짜":             dt,
                "공시제목":         d["etc"],
                "선주":             d["buyer"],
                "선종":             d["vessel_type"],
                "척수":             d["quantity"] or "",
                "인도년":           d["delivery_year"] or "",
                "인도월":           d["delivery_month"] or "",
                "금액(원화,십억원)": d["amount_krw_bil"] or "",
                "금액(달러,백만)":   d["amount_usd_mil"] or "",
                "확정":             d["confirmed"],
                "상선특수선":       d["vessel_category"],
            })
    print(f"CSV 저장: {out_path}  ({len(records)}건)")


# ── 메인 ──────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="HD현대중공업 수주계약 수집")
    ap.add_argument("--api-key", required=True, help="DART OpenAPI 인증키")
    ap.add_argument("--years",   type=int, default=6, help="과거 몇 년치 (기본 6)")
    ap.add_argument("--excel",   default="", help="엑셀 파일 경로 (뉴스수주 시트에 입력)")
    ap.add_argument("--csv-out", default="HD현대중공업_수주.csv", help="CSV 저장 경로")
    args = ap.parse_args()

    today = datetime.today()
    start = today.replace(year=today.year - args.years, month=1, day=1)

    # 1년 단위 분할
    ranges, cur = [], start
    while cur < today:
        nxt = min(cur + timedelta(days=364), today)
        ranges.append((cur.strftime("%Y%m%d"), nxt.strftime("%Y%m%d")))
        cur = nxt + timedelta(days=1)

    print(f"HD현대중공업 수주 수집 ({start.strftime('%Y-%m-%d')} ~ {today.strftime('%Y-%m-%d')})")
    print(f"조회 구간: {len(ranges)}개\n")

    # 엑셀 로드
    wb, ws = None, None
    if args.excel:
        ep = Path(args.excel)
        if not ep.exists():
            print(f"[오류] 파일 없음: {ep}")
            return
        wb = load_workbook(str(ep))
        if SHEET_NAME not in wb.sheetnames:
            print(f"[오류] '{SHEET_NAME}' 시트 없음. 시트 목록: {wb.sheetnames}")
            return
        ws = wb[SHEET_NAME]
        print(f"엑셀: {ep}")

    records = []
    added = skipped = 0

    for bgn, end in ranges:
        print(f"  {bgn} ~ {end} 조회...")
        items = get_list(args.api_key, bgn, end)
        print(f"    공시 {len(items)}건")

        for item in items:
            rcept_no  = item.get("rcept_no", "")
            rcept_dt  = item.get("rcept_dt", "")
            report_nm = item.get("report_nm", "")
            print(f"      {report_nm} ({rcept_dt})")

            xml = get_xml(args.api_key, rcept_no)
            if not xml:
                skipped += 1
                continue

            data = parse(xml, rcept_dt, report_nm)

            if ws is not None:
                if already_exists(ws, data["contract_date"]):
                    print(f"      → 중복, 건너뜀")
                    skipped += 1
                    continue
                nr = find_next_empty_row(ws)
                write_excel_row(ws, nr, data)
                print(f"      → {nr}행 입력")
                added += 1
            else:
                added += 1

            records.append(data)
            time.sleep(0.5)

        time.sleep(0.3)

    # 저장
    if wb is not None:
        wb.save(args.excel)
        print(f"\n엑셀 저장 완료: {args.excel}  ({added}건 추가, {skipped}건 건너뜀)")

    save_csv(records, args.csv_out)


if __name__ == "__main__":
    main()
