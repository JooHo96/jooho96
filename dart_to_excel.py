"""
DART 단일판매·공급계약 공시 → 엑셀 자동 입력 스크립트

일반 모드 (최근 N일):
    python dart_to_excel.py --api-key YOUR_KEY --days 7 --excel 파일.xlsx

과거 6년치 수집 모드:
    python dart_to_excel.py --api-key YOUR_KEY --history --excel 파일.xlsx

CSV 출력만 (엑셀 없이):
    python dart_to_excel.py --api-key YOUR_KEY --history --csv-only
"""

import argparse
import csv
import re
import time
import zipfile
import io
import requests
import openpyxl
from datetime import datetime, timedelta
from openpyxl import load_workbook
from pathlib import Path

# ── 대상 조선사 (회사명: DART corp_code) ──────────────────────────────────
TARGET_COMPANIES = {
    "HD현대중공업":   "01390344",
    "HD현대삼호":     "00332468",
    "HD현대미포":     "00164609",
    "HD한국조선해양": "00164830",
    "삼성중공업":     "00126478",
    "한화오션":       "00111704",
    "HJ중공업":       "00633835",
    "대한조선":       "01561465",
    "한화엔진":       "00361008",
}

# 과거 데이터 수집 대상 (요청 기업만)
HISTORY_COMPANIES = {
    "삼성중공업":     "00126478",
    "한화오션":       "00111704",
    "HD현대중공업":   "01390344",
    "HD현대미포":     "00164609",
    "HD한국조선해양": "00164830",
}

SHEET_NAME = "뉴스수주"
DART_BASE  = "https://opendart.fss.or.kr/api"
HISTORY_YEARS = 6   # 과거 몇 년치

# 컬럼 인덱스 (1-based, openpyxl) — 업로드 양식 기준
COL = {
    "A":  1,   # =AC{row} 수식
    "B":  2,   # 회사
    "C":  3,   # 날짜
    "D":  4,   # 출처
    "E":  5,   # 기타
    "F":  6,   # Type(선종)
    "G":  7,   # Name(선주)
    "H":  8,   # Size
    "I":  9,   # Unit(척수)
    "J":  10,  # Dwt
    "K":  11,  # GT
    "L":  12,  # CGT
    "M":  13,  # Flag
    "N":  14,  # Built(인도년)
    "O":  15,  # Month(인도월)
    "P":  16,  # Builder
    "Q":  17,  # Contract Date
    "R":  18,  # Company
    "S":  19,  # Group Company
    "T":  20,  # Operator
    "U":  21,  # Status
    "V":  22,  # Alternative Fuel Types
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


# ── DART API 호출 ─────────────────────────────────────────────────────────

def get_disclosure_list(api_key: str, corp_code: str, bgn_de: str, end_de: str) -> list:
    """단일판매·공급계약(B002) 공시 목록 조회 (페이지네이션 포함)"""
    url = f"{DART_BASE}/list.json"
    results = []
    page_no = 1

    while True:
        params = {
            "crtfc_key":        api_key,
            "corp_code":        corp_code,
            "pblntf_detail_ty": "B002",
            "bgn_de":           bgn_de,
            "end_de":           end_de,
            "page_count":       100,
            "page_no":          page_no,
        }
        try:
            r = requests.get(url, params=params, timeout=15)
            data = r.json()
        except Exception as e:
            print(f"  [오류] 목록 조회 실패: {e}")
            break

        if data.get("status") == "013":  # 조회 결과 없음
            break
        if data.get("status") != "000":
            print(f"  [경고] DART 응답 상태: {data.get('status')} {data.get('message','')}")
            break

        page_items = data.get("list", [])
        results.extend(page_items)

        total = int(data.get("total_count", 0))
        if page_no * 100 >= total:
            break
        page_no += 1
        time.sleep(0.3)

    return results


def get_disclosure_xml(api_key: str, rcept_no: str) -> str | None:
    """공시 원문 XML 다운로드 (zip → xml 텍스트)"""
    url = f"{DART_BASE}/document.xml"
    params = {"crtfc_key": api_key, "rcept_no": rcept_no}
    try:
        r = requests.get(url, params=params, timeout=20)
        if r.headers.get("content-type", "").startswith("application/zip") or r.content[:2] == b"PK":
            with zipfile.ZipFile(io.BytesIO(r.content)) as z:
                for name in z.namelist():
                    if name.endswith(".xml"):
                        return z.read(name).decode("utf-8", errors="ignore")
        return r.text
    except Exception as e:
        print(f"  [오류] 원문 다운로드 실패: {e}")
        return None


# ── 공시 내용 파싱 ────────────────────────────────────────────────────────

def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def find_value(xml_text: str, *keywords) -> str:
    """키워드 다음 셀 값을 테이블에서 찾아 반환"""
    for kw in keywords:
        # 테이블 태그 기반 탐색
        pattern = rf"{re.escape(kw)}[^<]*</[^>]+>\s*<[^>]+>([^<]+)</"
        m = re.search(pattern, xml_text)
        if m:
            val = clean(m.group(1))
            if val and val not in ("-", "해당없음", "N/A"):
                return val
        # 텍스트 기반 탐색 (줄바꿈 등 변형)
        pattern2 = rf"{re.escape(kw)}[^\n]{{0,20}}\n\s*([^\n<]{{1,200}})"
        m2 = re.search(pattern2, xml_text)
        if m2:
            val = clean(m2.group(1))
            if val and val not in ("-", "해당없음"):
                return val
    return ""


def parse_amount_krw(text: str) -> float | None:
    """금액 텍스트 → 십억원 변환"""
    text = text.replace(",", "").replace(" ", "")
    if not text:
        return None
    m = re.search(r"[\d.]+", text)
    if not m:
        return None
    v = float(m.group())
    if "조" in text:
        return v * 1_000
    elif "억" in text:
        return round(v / 10, 3)       # 억원 → 십억원
    elif "백만원" in text:
        return round(v / 1_000, 3)    # 백만원 → 십억원
    elif "천만" in text:
        return round(v / 100, 3)
    elif "원" in text:
        return round(v / 1_000_000_000, 3)
    return None


def parse_amount_usd(text: str) -> float | None:
    """USD 금액 텍스트 → 백만달러"""
    text = text.replace(",", "").replace(" ", "")
    m = re.search(r"[\d.]+", text)
    if not m:
        return None
    v = float(m.group())
    if "억달러" in text or "억USD" in text.upper():
        return round(v * 100, 2)      # 억달러 → 백만달러
    elif "백만달러" in text or "백만USD" in text.upper():
        return round(v, 2)
    elif "천만달러" in text:
        return round(v * 10, 2)
    return None


def parse_disclosure(xml_text: str, company: str, rcept_dt: str, report_nm: str = "") -> dict | None:
    """공시 XML 파싱 → 엑셀 행 데이터 dict 반환"""
    row = {
        "company":        company,
        "date":           None,
        "source":         "DART",
        "etc":            report_nm,
        "vessel_type":    "",
        "buyer":          "",
        "size":           "",
        "quantity":       None,
        "delivery_year":  None,
        "delivery_month": None,
        "contract_date":  None,
        "amount_usd_mil": None,
        "amount_krw_bil": None,
        "unit_price_usd": None,
        "exchange_rate":  None,
        "confirmed":      "O",
        "vessel_category": "상선",
        "clarksons":      "",
        "delivery_date":  None,
    }

    # 계약일 (공시 접수일 사용)
    if rcept_dt and len(rcept_dt) == 8:
        try:
            dt = datetime(int(rcept_dt[:4]), int(rcept_dt[4:6]), int(rcept_dt[6:8]))
            row["date"] = dt
            row["contract_date"] = dt
        except Exception:
            pass

    # 계약 상대방(선주)
    buyer = find_value(xml_text, "계약상대방", "거래상대방", "매수인", "발주처", "수요자")
    row["buyer"] = buyer

    # 계약금액 (원화 우선, 없으면 외화)
    amount_text = find_value(xml_text, "계약금액", "총계약금액", "공급금액", "거래금액")
    if amount_text:
        if "달러" in amount_text or "USD" in amount_text.upper():
            row["amount_usd_mil"] = parse_amount_usd(amount_text)
        else:
            row["amount_krw_bil"] = parse_amount_krw(amount_text)

    # 선종 분류
    purpose = find_value(xml_text, "계약내용", "공급내용", "계약목적물", "품목", "제품명", "선박종류")
    search_text = purpose + " " + xml_text[:5000]

    vessel_map = [
        ("컨테이너",  "컨테이너선"),
        ("LNG",       "LNG선"),
        ("LPG",       "LPG선"),
        ("암모니아",  "LPG선"),
        ("VLCC",      "VLCC"),
        ("원유운반",  "원유운반선"),
        ("탱커",      "탱커"),
        ("PC선",      "P/C선"),
        ("제품운반",  "P/C선"),
        ("MR탱커",    "P/C선"),
        ("벌크",      "벌크선"),
        ("살물선",    "벌크선"),
        ("풍력",      "해양"),
        ("해양플랜트","해양"),
        ("FPSO",      "해양"),
        ("특수선",    "특수선"),
        ("군함",      "특수선"),
        ("잠수함",    "특수선"),
    ]
    vtype = ""
    for kw, label in vessel_map:
        if kw.upper() in search_text.upper():
            vtype = label
            break
    row["vessel_type"] = vtype
    if vtype in ("해양", "특수선"):
        row["vessel_category"] = vtype

    # 척수
    qty_text = find_value(xml_text, "수량", "척수", "호선수", "선박수")
    if qty_text:
        m = re.search(r"\d+", qty_text)
        if m:
            row["quantity"] = int(m.group())

    # 인도 예정일
    delivery_text = find_value(xml_text, "납기", "인도예정", "납품예정", "공급일정", "인도일")
    if delivery_text:
        m = re.search(r"(\d{4})[.\-년]\s*(\d{1,2})", delivery_text)
        if m:
            y, mo = int(m.group(1)), int(m.group(2))
            row["delivery_year"]  = y
            row["delivery_month"] = mo
            try:
                import calendar
                last_day = calendar.monthrange(y, mo)[1]
                row["delivery_date"] = datetime(y, mo, last_day)
            except Exception:
                pass

    return row


# ── 조회 기간 분할 (DART API: 최대 1년) ──────────────────────────────────

def split_date_ranges(start: datetime, end: datetime, chunk_days: int = 365) -> list:
    """기간을 chunk_days 단위로 분할"""
    ranges = []
    cur = start
    while cur < end:
        nxt = min(cur + timedelta(days=chunk_days - 1), end)
        ranges.append((cur.strftime("%Y%m%d"), nxt.strftime("%Y%m%d")))
        cur = nxt + timedelta(days=1)
    return ranges


# ── 엑셀 유틸 ─────────────────────────────────────────────────────────────

def find_next_empty_row(ws) -> int:
    for r in range(3, ws.max_row + 2):
        if ws.cell(row=r, column=COL["B"]).value is None:
            return r
    return ws.max_row + 1


def already_exists(ws, company: str, contract_date: datetime | None) -> bool:
    if contract_date is None:
        return False
    for r in range(3, ws.max_row + 1):
        b = ws.cell(row=r, column=COL["B"]).value
        q = ws.cell(row=r, column=COL["Q"]).value
        if b == company and isinstance(q, datetime) and q.date() == contract_date.date():
            return True
    return False


def write_row(ws, row_num: int, data: dict):
    r = row_num
    ws.cell(row=r, column=COL["A"]).value  = f"=AC{r}"
    ws.cell(row=r, column=COL["B"]).value  = data["company"]
    ws.cell(row=r, column=COL["C"]).value  = data["date"]
    ws.cell(row=r, column=COL["D"]).value  = data["source"]
    ws.cell(row=r, column=COL["E"]).value  = data["etc"] or None
    ws.cell(row=r, column=COL["F"]).value  = data["vessel_type"] or None
    ws.cell(row=r, column=COL["G"]).value  = data["buyer"] or None
    ws.cell(row=r, column=COL["H"]).value  = data["size"] or None
    ws.cell(row=r, column=COL["I"]).value  = data["quantity"]      # Unit(척수)
    ws.cell(row=r, column=COL["N"]).value  = data["delivery_year"]  # Built(인도년)
    ws.cell(row=r, column=COL["O"]).value  = data["delivery_month"] # Month(인도월)
    ws.cell(row=r, column=COL["P"]).value  = data["company"]        # Builder = 회사명
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


# ── CSV 출력 ──────────────────────────────────────────────────────────────

def write_csv(records: list, out_path: str):
    fieldnames = [
        "회사", "날짜", "출처", "선종", "선주", "척수",
        "인도년", "인도월", "계약금액(백만달러)", "계약금액(십억원)",
        "확정여부", "상선특수선", "비고"
    ]
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for d in records:
            dt = d["contract_date"].strftime("%Y-%m-%d") if d["contract_date"] else ""
            w.writerow({
                "회사":            d["company"],
                "날짜":            dt,
                "출처":            d["source"],
                "선종":            d["vessel_type"],
                "선주":            d["buyer"],
                "척수":            d["quantity"] or "",
                "인도년":          d["delivery_year"] or "",
                "인도월":          d["delivery_month"] or "",
                "계약금액(백만달러)": d["amount_usd_mil"] or "",
                "계약금액(십억원)":  d["amount_krw_bil"] or "",
                "확정여부":        d["confirmed"],
                "상선특수선":      d["vessel_category"],
                "비고":            d["etc"],
            })
    print(f"CSV 저장: {out_path}  ({len(records)}건)")


# ── 공시 수집 공통 로직 ───────────────────────────────────────────────────

def collect_disclosures(api_key: str, companies: dict, date_ranges: list,
                        ws=None, verbose: bool = True) -> list:
    """
    companies: {이름: corp_code}
    date_ranges: [(bgn_de, end_de), ...]
    ws: 엑셀 시트 (None이면 중복 체크 없이 수집만)
    반환: 파싱된 data dict 리스트
    """
    all_records = []
    added = skipped = 0

    for company, corp_code in companies.items():
        print(f"\n{'='*50}")
        print(f"[{company}] 조회 시작 (corp_code: {corp_code})")

        company_records = []

        for bgn_de, end_de in date_ranges:
            if verbose:
                print(f"  기간: {bgn_de} ~ {end_de}")
            items = get_disclosure_list(api_key, corp_code, bgn_de, end_de)

            for item in items:
                rcept_no  = item.get("rcept_no", "")
                rcept_dt  = item.get("rcept_dt", "")
                report_nm = item.get("report_nm", "")

                if verbose:
                    print(f"    공시: {report_nm} ({rcept_dt})")

                xml_text = get_disclosure_xml(api_key, rcept_no)
                if not xml_text:
                    skipped += 1
                    continue

                data = parse_disclosure(xml_text, company, rcept_dt, report_nm)
                if data is None:
                    skipped += 1
                    continue

                # 엑셀 있으면 중복 체크 및 쓰기
                if ws is not None:
                    if already_exists(ws, company, data["contract_date"]):
                        if verbose:
                            print(f"    → 이미 존재, 건너뜀")
                        skipped += 1
                        continue
                    next_row = find_next_empty_row(ws)
                    write_row(ws, next_row, data)
                    added += 1
                    if verbose:
                        print(f"    → {next_row}행에 추가")
                else:
                    added += 1

                company_records.append(data)
                time.sleep(0.5)

            time.sleep(0.3)  # 기간 간 대기

        print(f"  [{company}] 수집: {len(company_records)}건")
        all_records.extend(company_records)

    print(f"\n총계: {added}건 처리, {skipped}건 건너뜀")
    return all_records


# ── 메인 ──────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="DART 단일판매·공급계약 → 엑셀/CSV")
    parser.add_argument("--api-key",   required=True, help="DART OpenAPI 인증키")
    parser.add_argument("--days",      type=int, default=7, help="최근 N일 조회 (기본 7)")
    parser.add_argument("--excel",     default="", help="엑셀 파일 경로")
    parser.add_argument("--history",   action="store_true", help=f"과거 {HISTORY_YEARS}년치 수집")
    parser.add_argument("--csv-only",  action="store_true", help="CSV만 출력 (엑셀 불필요)")
    parser.add_argument("--csv-out",   default="dart_history.csv", help="CSV 저장 경로")
    parser.add_argument("--companies", default="all",
                        help="수집 대상: all | history (삼성,한화오션,HD현대중공업,HD현대미포,HD한국조선해양)")
    args = parser.parse_args()

    # 대상 기업 선택
    if args.companies == "history" or args.history:
        companies = HISTORY_COMPANIES
    else:
        companies = TARGET_COMPANIES

    # 날짜 범위 설정
    today = datetime.today()
    if args.history:
        start = today.replace(year=today.year - HISTORY_YEARS, month=1, day=1)
        end   = today
        print(f"▶ 과거 {HISTORY_YEARS}년치 수집 모드: {start.strftime('%Y-%m-%d')} ~ {end.strftime('%Y-%m-%d')}")
    else:
        end   = today
        start = end - timedelta(days=args.days)
        print(f"▶ 최근 {args.days}일 수집 모드: {start.strftime('%Y-%m-%d')} ~ {end.strftime('%Y-%m-%d')}")

    # DART API: 최대 1년 단위로 분할
    date_ranges = split_date_ranges(start, end)
    print(f"▶ 조회 구간: {len(date_ranges)}개")

    # 엑셀 로드 (csv-only가 아닐 때)
    ws = None
    wb = None
    if not args.csv_only:
        if not args.excel:
            print("[오류] --excel 경로를 지정하거나 --csv-only 옵션을 사용하세요.")
            return
        excel_path = Path(args.excel)
        if not excel_path.exists():
            print(f"[오류] 엑셀 파일 없음: {excel_path}")
            return
        wb = load_workbook(str(excel_path))
        if SHEET_NAME not in wb.sheetnames:
            print(f"[오류] 시트 '{SHEET_NAME}' 없음. 시트 목록: {wb.sheetnames}")
            return
        ws = wb[SHEET_NAME]
        print(f"▶ 엑셀: {excel_path}  시트: {SHEET_NAME}")

    # 수집 실행
    records = collect_disclosures(
        api_key=args.api_key,
        companies=companies,
        date_ranges=date_ranges,
        ws=ws,
        verbose=True,
    )

    # 엑셀 저장
    if wb is not None and ws is not None:
        save_path = str(excel_path)
        wb.save(save_path)
        print(f"\n▶ 엑셀 저장 완료: {save_path}")

    # CSV 저장 (--csv-only 또는 --history 시 항상 출력)
    if args.csv_only or args.history:
        write_csv(records, args.csv_out)


if __name__ == "__main__":
    main()
