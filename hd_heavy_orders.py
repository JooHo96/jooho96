"""
HD현대중공업 단일판매·공급계약 공시 수집 → CSV + 엑셀(뉴스수주 양식)

사용법:
    pip install requests openpyxl

    # CSV만 출력 (6년치)
    python hd_heavy_orders.py --api-key YOUR_DART_KEY

    # 엑셀 양식에 직접 입력
    python hd_heavy_orders.py --api-key YOUR_DART_KEY --excel "수주현황.xlsx"
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

COL = {
    "A":  1,  "B":  2,  "C":  3,  "D":  4,  "E":  5,
    "F":  6,  "G":  7,  "H":  8,  "I":  9,
    "N":  14, "O":  15, "P":  16, "Q":  17,
    "Z":  26, "AA": 27, "AB": 28, "AC": 29,
    "AD": 30, "AE": 31, "AF": 32, "AG": 33, "AH": 34,
}


# ── DART API ──────────────────────────────────────────────────────────────

def get_list(api_key, bgn_de, end_de):
    results, page = [], 1
    while True:
        try:
            r = requests.get(f"{DART_BASE}/list.json", params={
                "crtfc_key": api_key, "corp_code": CORP_CODE,
                "bgn_de": bgn_de, "end_de": end_de,
                "page_count": 100, "page_no": page,
            }, timeout=15).json()
        except Exception as e:
            print(f"    [오류] {e}")
            break
        if r.get("status") not in ("000",):
            break
        for item in r.get("list", []):
            nm = item.get("report_nm", "")
            if "단일판매" in nm or "공급계약" in nm:
                results.append(item)
        if page * 100 >= int(r.get("total_count", 0)):
            break
        page += 1
        time.sleep(0.3)
    return results


def get_html(api_key, rcept_no):
    try:
        r = requests.get(f"{DART_BASE}/document.xml",
                         params={"crtfc_key": api_key, "rcept_no": rcept_no}, timeout=20)
        if r.content[:2] == b"PK":
            with zipfile.ZipFile(io.BytesIO(r.content)) as z:
                for n in z.namelist():
                    if n.endswith(".xml") or n.endswith(".html"):
                        return z.read(n).decode("euc-kr", errors="ignore")
        return r.content.decode("euc-kr", errors="ignore")
    except Exception as e:
        print(f"    [오류] 문서 다운로드: {e}")
        return None


# ── HTML 파싱 ─────────────────────────────────────────────────────────────

def extract_td_values(html):
    """
    DART 공시 HTML 테이블에서 레이블→값 매핑 딕셔너리 반환
    <td>레이블</td><td>값</td> 구조
    """
    # td 안의 텍스트 추출
    tds = re.findall(r'<td[^>]*>(.*?)</td>', html, re.DOTALL | re.IGNORECASE)
    result = {}
    clean_tds = []
    for td in tds:
        # HTML 태그 제거, 공백 정리
        text = re.sub(r'<[^>]+>', '', td)
        text = re.sub(r'\s+', ' ', text).strip()
        text = text.replace('\xa0', '').replace('&nbsp;', '').strip()
        clean_tds.append(text)

    # 연속된 td에서 레이블-값 쌍 추출
    i = 0
    while i < len(clean_tds) - 1:
        label = clean_tds[i]
        value = clean_tds[i + 1]
        if label and value and len(label) < 50:
            result[label] = value
        i += 1

    return result, clean_tds


def find_val(mapping, *keys):
    """매핑 딕셔너리에서 키워드 검색"""
    for key in keys:
        for k, v in mapping.items():
            if key in k and v and v not in ('-', '해당없음', 'N/A', '없음'):
                return v
    return ""


def parse_krw(text):
    """원화 금액 문자열 → 십억원"""
    text = re.sub(r'[,\s원]', '', text)
    m = re.search(r'[\d.]+', text)
    if not m:
        return None
    v = float(m.group())
    # 단위 판단: 자릿수로 추정 (원 단위면 10자리 이상)
    raw = re.sub(r'[^\d]', '', text)
    if len(raw) >= 10:          # 원 단위
        return round(v / 1_000_000_000, 3)
    elif len(raw) >= 7:         # 백만원 단위
        return round(v / 1_000, 3)
    elif '억' in text:
        return round(v / 10, 3)
    elif '백만' in text:
        return round(v / 1_000, 3)
    elif '조' in text:
        return round(v * 1_000, 3)
    return None


def parse_usd(text):
    """달러 금액 문자열 → 백만달러"""
    text = re.sub(r'[,\s]', '', text)
    m = re.search(r'[\d.]+', text)
    if not m:
        return None
    v = float(m.group())
    if '억달러' in text or '억USD' in text.upper():
        return round(v * 100, 2)
    elif '백만달러' in text or '백만USD' in text.upper():
        return round(v, 2)
    elif '천만달러' in text:
        return round(v * 10, 2)
    return None


VESSEL_MAP = [
    ("컨테이너",   "컨테이너선"),
    ("LNG",        "LNG선"),
    ("LPG",        "LPG선"),
    ("VLGC",       "LPG선"),
    ("MGC",        "LPG선"),
    ("암모니아",   "LPG선"),
    ("VLCC",       "원유운반선"),
    ("Suezmax",    "원유운반선"),
    ("원유운반",   "원유운반선"),
    ("탱커",       "탱커"),
    ("PC선",       "P/C선"),
    ("제품운반",   "P/C선"),
    ("MR탱커",     "P/C선"),
    ("MR P/C",     "P/C선"),
    ("LR2",        "P/C선"),
    ("PCTC",       "기타"),
    ("자동차운반", "기타"),
    ("벌크",       "벌크선"),
    ("살물선",     "벌크선"),
    ("FPSO",       "해양"),
    ("풍력",       "해양"),
    ("해양플랜트", "해양"),
    ("수상함",     "특수선"),
    ("군함",       "특수선"),
    ("잠수함",     "특수선"),
    ("쇄빙",       "특수선"),
    ("특수선",     "특수선"),
]


def detect_vessel_type(text):
    u = text.upper()
    for kw, label in VESSEL_MAP:
        if kw.upper() in u:
            return label
    return ""


def parse(html, rcept_dt, report_nm):
    mapping, tds = extract_td_values(html)
    full_text = " ".join(tds)

    row = {
        "date": None, "contract_date": None,
        "buyer": "", "vessel_type": "", "vessel_category": "상선",
        "etc": "", "size": "", "quantity": None,
        "delivery_year": None, "delivery_month": None, "delivery_date": None,
        "amount_usd_mil": None, "amount_krw_bil": None,
        "unit_price_usd": None, "exchange_rate": None,
        "confirmed": "O", "clarksons": "",
    }

    # 날짜
    if len(rcept_dt) == 8:
        try:
            dt = datetime(int(rcept_dt[:4]), int(rcept_dt[4:6]), int(rcept_dt[6:]))
            row["date"] = dt
            row["contract_date"] = dt
        except Exception:
            pass

    # 선주 (계약상대방)
    row["buyer"] = find_val(mapping, "계약상대방", "거래상대방", "발주처", "매수인", "수요자")

    # 계약금액
    amt = find_val(mapping, "계약금액", "총계약금액", "공급금액")
    if amt:
        if "달러" in amt or "USD" in amt.upper():
            row["amount_usd_mil"] = parse_usd(amt)
        else:
            row["amount_krw_bil"] = parse_krw(amt)

    # 계약내용 (선종 파악)
    purpose = find_val(mapping, "계약내용", "공급내용", "체결계약명", "계약목적물", "품목")
    row["etc"] = purpose  # E열(기타)에 계약명 저장

    # 선종: 계약내용 + 전체 텍스트에서 탐지
    vtype = detect_vessel_type(purpose) or detect_vessel_type(full_text[:3000])
    row["vessel_type"] = vtype
    if vtype in ("해양", "특수선"):
        row["vessel_category"] = vtype

    # 척수
    qty_str = find_val(mapping, "수량", "척수", "호선수", "선박수", "계약수량")
    if not qty_str:
        # 계약명에서 "N척" 패턴 추출
        m = re.search(r'(\d+)\s*척', purpose + " " + full_text[:500])
        if m:
            qty_str = m.group(1)
    if qty_str:
        m = re.search(r'\d+', qty_str)
        if m:
            row["quantity"] = int(m.group())

    # 인도 예정일
    dlv = find_val(mapping, "납기", "인도예정", "납품예정", "인도일", "공급기간", "납품기간")
    if dlv:
        m = re.search(r'(\d{4})[.\-년]\s*(\d{1,2})', dlv)
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


# ── 엑셀 ──────────────────────────────────────────────────────────────────

def find_next_empty_row(ws):
    for r in range(3, ws.max_row + 2):
        if ws.cell(row=r, column=COL["B"]).value is None:
            return r
    return ws.max_row + 1


def already_exists(ws, contract_date):
    if not contract_date:
        return False
    for r in range(3, ws.max_row + 1):
        b = ws.cell(row=r, column=COL["B"]).value
        q = ws.cell(row=r, column=COL["Q"]).value
        if b == CORP_NAME and isinstance(q, datetime) and q.date() == contract_date.date():
            return True
    return False


def write_excel_row(ws, rn, d):
    ws.cell(row=rn, column=COL["A"]).value  = f"=AC{rn}"
    ws.cell(row=rn, column=COL["B"]).value  = CORP_NAME
    ws.cell(row=rn, column=COL["C"]).value  = d["date"]
    ws.cell(row=rn, column=COL["D"]).value  = "DART"
    ws.cell(row=rn, column=COL["E"]).value  = d["etc"] or None
    ws.cell(row=rn, column=COL["F"]).value  = d["vessel_type"] or None
    ws.cell(row=rn, column=COL["G"]).value  = d["buyer"] or None
    ws.cell(row=rn, column=COL["H"]).value  = d["size"] or None
    ws.cell(row=rn, column=COL["I"]).value  = d["quantity"]
    ws.cell(row=rn, column=COL["N"]).value  = d["delivery_year"]
    ws.cell(row=rn, column=COL["O"]).value  = d["delivery_month"]
    ws.cell(row=rn, column=COL["P"]).value  = CORP_NAME
    ws.cell(row=rn, column=COL["Q"]).value  = d["contract_date"]
    ws.cell(row=rn, column=COL["Z"]).value  = d["delivery_date"]
    ws.cell(row=rn, column=COL["AA"]).value = d["amount_usd_mil"]
    if d["amount_usd_mil"] and d["exchange_rate"]:
        ws.cell(row=rn, column=COL["AB"]).value = f"=AA{rn}/AD{rn}*1000"
    elif d["amount_krw_bil"]:
        ws.cell(row=rn, column=COL["AB"]).value = d["amount_krw_bil"]
    if d["quantity"] and d["amount_usd_mil"]:
        ws.cell(row=rn, column=COL["AC"]).value = f"=AA{rn}/I{rn}"
    elif d["unit_price_usd"]:
        ws.cell(row=rn, column=COL["AC"]).value = d["unit_price_usd"]
    ws.cell(row=rn, column=COL["AD"]).value = d["exchange_rate"]
    ws.cell(row=rn, column=COL["AE"]).value = d["confirmed"]
    ws.cell(row=rn, column=COL["AF"]).value = f'=YEAR(Q{rn})&"."&MONTH(Q{rn})'
    ws.cell(row=rn, column=COL["AG"]).value = d["vessel_category"]
    ws.cell(row=rn, column=COL["AH"]).value = d["clarksons"] or None


# ── CSV ────────────────────────────────────────────────────────────────────

def save_csv(records, path):
    fields = ["회사", "날짜", "공시제목", "선주", "선종", "척수",
              "인도년", "인도월", "금액(원화,십억원)", "금액(달러,백만)", "확정", "상선특수선"]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for d in records:
            dt = d["contract_date"].strftime("%Y-%m-%d") if d["contract_date"] else ""
            w.writerow({
                "회사":             CORP_NAME,
                "날짜":             dt,
                "공시제목":         d.get("report_nm", ""),
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
    print(f"CSV 저장: {path}  ({len(records)}건)")


# ── 메인 ──────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="HD현대중공업 수주계약 수집")
    ap.add_argument("--api-key", required=True)
    ap.add_argument("--years",   type=int, default=6)
    ap.add_argument("--excel",   default="")
    ap.add_argument("--csv-out", default="HD현대중공업_수주.csv")
    args = ap.parse_args()

    today = datetime.today()
    start = today.replace(year=today.year - args.years, month=1, day=1)

    ranges, cur = [], start
    while cur < today:
        nxt = min(cur + timedelta(days=364), today)
        ranges.append((cur.strftime("%Y%m%d"), nxt.strftime("%Y%m%d")))
        cur = nxt + timedelta(days=1)

    print(f"HD현대중공업 수주 수집 ({start.strftime('%Y-%m-%d')} ~ {today.strftime('%Y-%m-%d')})")
    print(f"조회 구간: {len(ranges)}개\n")

    wb, ws = None, None
    if args.excel:
        ep = Path(args.excel)
        if not ep.exists():
            print(f"[오류] 파일 없음: {ep}"); return
        wb = load_workbook(str(ep))
        if SHEET_NAME not in wb.sheetnames:
            print(f"[오류] '{SHEET_NAME}' 시트 없음"); return
        ws = wb[SHEET_NAME]
        print(f"엑셀: {ep}\n")

    records = []
    added = skipped = 0

    for bgn, end in ranges:
        print(f"  {bgn[:4]}년 조회...")
        items = get_list(args.api_key, bgn, end)
        print(f"    공시 {len(items)}건")

        for item in items:
            rcept_no  = item.get("rcept_no", "")
            rcept_dt  = item.get("rcept_dt", "")
            report_nm = item.get("report_nm", "").strip()

            html = get_html(args.api_key, rcept_no)
            if not html:
                skipped += 1
                continue

            data = parse(html, rcept_dt, report_nm)
            data["report_nm"] = report_nm

            if ws is not None:
                if already_exists(ws, data["contract_date"]):
                    skipped += 1
                    continue
                nr = find_next_empty_row(ws)
                write_excel_row(ws, nr, data)
                added += 1
            else:
                added += 1

            print(f"      {data['contract_date'].strftime('%Y-%m-%d') if data['contract_date'] else '?'} | {data['vessel_type'] or '선종미상'} | {data['buyer'] or '선주미상'} | {data['quantity'] or '?'}척 | ${data['amount_usd_mil'] or data['amount_krw_bil'] or '?'}")
            records.append(data)
            time.sleep(0.5)

        time.sleep(0.3)

    if wb:
        wb.save(args.excel)
        print(f"\n엑셀 저장: {args.excel}  ({added}건 추가, {skipped}건 건너뜀)")

    save_csv(records, args.csv_out)


if __name__ == "__main__":
    main()
