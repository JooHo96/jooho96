"""
DART 단일판매·공급계약 공시 → 엑셀 자동 입력 스크립트
사용법: python dart_to_excel.py --api-key YOUR_KEY [--days 7]
"""

import argparse
import re
import time
import zipfile
import io
import json
import requests
import openpyxl
from datetime import datetime, timedelta
from openpyxl import load_workbook
from pathlib import Path

# ── 대상 조선사 (회사명: DART corp_code) ──────────────────────────────────
# corp_code는 아래 get_corp_code()로 자동 조회하거나 직접 지정
TARGET_COMPANIES = {
    "HD현대중공업":  "00164788",
    "HD현대삼호":    "00105946",
    "HD현대미포":    "00102773",
    "삼성중공업":    "00126380",
    "한화오션":      "00105257",
    "HJ중공업":      "00106592",
    "대한조선":      "01308627",
    "한화엔진":      "00106563",
}

EXCEL_PATH = Path("/root/.claude/uploads/e3908c3c-2101-4e5a-9818-257097a4355c/b56bbae4-______.xlsx")
SHEET_NAME = "뉴스수주"

# 컬럼 인덱스 (1-based, openpyxl)
COL = {
    "A":  1,   # =AC{row} 수식
    "B":  2,   # 회사
    "C":  3,   # 날짜
    "D":  4,   # 출처
    "E":  5,   # 기타(선형 등)
    "F":  6,   # Type(선종)
    "G":  7,   # Name(선주)
    "H":  8,   # Size(TEU 등)
    "I":  9,   # 척수(Dwt 칸 활용)
    "J":  10,  # GT
    "K":  11,  # CGT
    "L":  12,  # Flag
    "M":  13,  # Built
    "N":  14,  # 인도년
    "O":  15,  # 인도월
    "P":  16,  # Builder (=B{row})
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

DART_BASE = "https://opendart.fss.or.kr/api"


# ── DART API 호출 ─────────────────────────────────────────────────────────

def get_disclosure_list(api_key: str, corp_code: str, bgn_de: str, end_de: str) -> list:
    """단일판매·공급계약(B002) 공시 목록 조회"""
    url = f"{DART_BASE}/list.json"
    params = {
        "crtfc_key": api_key,
        "corp_code": corp_code,
        "pblntf_detail_ty": "B002",
        "bgn_de": bgn_de,
        "end_de": end_de,
        "page_count": 100,
    }
    try:
        r = requests.get(url, params=params, timeout=10)
        data = r.json()
        if data.get("status") == "000":
            return data.get("list", [])
        return []
    except Exception as e:
        print(f"  [오류] 목록 조회 실패: {e}")
        return []


def get_disclosure_xml(api_key: str, rcept_no: str) -> str | None:
    """공시 원문 XML 다운로드 (zip → xml 텍스트 반환)"""
    url = f"{DART_BASE}/document.xml"
    params = {"crtfc_key": api_key, "rcept_no": rcept_no}
    try:
        r = requests.get(url, params=params, timeout=15)
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
    """키워드 다음 셀 값을 찾아 반환"""
    for kw in keywords:
        pattern = rf"{re.escape(kw)}[^<]*</[^>]+>\s*<[^>]+>([^<]+)</"
        m = re.search(pattern, xml_text)
        if m:
            return clean(m.group(1))
    return ""


def parse_amount(text: str) -> float | None:
    """'123,456백만원' 또는 '123.4백만달러' 형태에서 숫자 추출"""
    text = text.replace(",", "")
    m = re.search(r"[\d.]+", text)
    return float(m.group()) if m else None


def parse_disclosure(xml_text: str, company: str, rcept_dt: str) -> dict | None:
    """공시 XML 파싱 → 엑셀 행 데이터 dict 반환"""
    row = {
        "company": company,
        "date": None,
        "source": "DART",
        "etc": "",
        "vessel_type": "",
        "buyer": "",
        "size": "",
        "quantity": None,
        "delivery_year": None,
        "delivery_month": None,
        "contract_date": None,
        "amount_usd_mil": None,  # 총계약금액 백만달러
        "amount_krw_bil": None,  # 총계약금액 십억원
        "unit_price_usd": None,  # 척당 백만달러
        "exchange_rate": None,
        "confirmed": "O",
        "vessel_category": "상선",
        "clarksons": "",
        "delivery_date": None,
    }

    # 계약일
    rcept_dt_parsed = None
    if rcept_dt and len(rcept_dt) == 8:
        try:
            rcept_dt_parsed = datetime(int(rcept_dt[:4]), int(rcept_dt[4:6]), int(rcept_dt[6:8]))
            row["date"] = rcept_dt_parsed
            row["contract_date"] = rcept_dt_parsed
        except Exception:
            pass

    # 계약 상대방(선주)
    buyer = find_value(xml_text, "계약상대방", "거래상대방", "매수인", "발주처")
    row["buyer"] = buyer

    # 계약금액
    amount_text = find_value(xml_text, "계약금액", "총계약금액")
    if amount_text:
        # 원화 금액 (억원 또는 백만원)
        if "억" in amount_text:
            v = parse_amount(amount_text)
            if v:
                row["amount_krw_bil"] = round(v / 10, 2)  # 억원 → 십억원
        elif "백만원" in amount_text or "원" in amount_text.lower():
            v = parse_amount(amount_text)
            if v:
                row["amount_krw_bil"] = round(v / 1000, 2)  # 백만원 → 십억원
        elif "달러" in amount_text or "USD" in amount_text.upper():
            v = parse_amount(amount_text)
            if v:
                # 백만달러 단위 판단
                if "억" in amount_text:
                    row["amount_usd_mil"] = round(v * 100, 2)
                else:
                    row["amount_usd_mil"] = v

    # 계약 목적물 (선종 추측)
    purpose = find_value(xml_text, "계약내용", "공급내용", "계약목적물", "품목")
    vtype = ""
    for kw, label in [
        ("컨테이너", "컨테이너선"), ("LNG", "LNG선"), ("LPG", "LPG선"),
        ("VLCC", "원유운반선"), ("원유", "원유운반선"), ("탱커", "원유운반선"),
        ("PC선", "P/C선"), ("제품운반", "P/C선"), ("벌크", "벌크선"),
        ("특수선", "특수선"), ("해양", "해양"), ("풍력", "해양"),
    ]:
        if kw in purpose or kw in xml_text[:3000]:
            vtype = label
            break
    row["vessel_type"] = vtype

    # 특수선 분류
    if vtype in ("특수선", "해양"):
        row["vessel_category"] = vtype

    # 척수
    qty_text = find_value(xml_text, "수량", "척수", "호선수")
    if qty_text:
        m = re.search(r"\d+", qty_text)
        if m:
            row["quantity"] = int(m.group())

    # 인도 예정일
    delivery_text = find_value(xml_text, "납기", "인도예정", "납품예정", "공급일정")
    if delivery_text:
        m = re.search(r"(\d{4})[.\-년](\d{1,2})", delivery_text)
        if m:
            row["delivery_year"] = int(m.group(1))
            row["delivery_month"] = int(m.group(2))
            try:
                import calendar
                last_day = calendar.monthrange(row["delivery_year"], row["delivery_month"])[1]
                row["delivery_date"] = datetime(row["delivery_year"], row["delivery_month"], last_day)
            except Exception:
                pass

    return row


# ── 엑셀 쓰기 ─────────────────────────────────────────────────────────────

def find_next_empty_row(ws) -> int:
    """헤더(1~2행) 이후 첫 번째 빈 행 번호 반환"""
    for r in range(3, ws.max_row + 2):
        if ws.cell(row=r, column=COL["B"]).value is None:
            return r
    return ws.max_row + 1


def already_exists(ws, company: str, contract_date: datetime | None) -> bool:
    """동일 회사 + 계약일 중복 체크"""
    if contract_date is None:
        return False
    for r in range(3, ws.max_row + 1):
        b = ws.cell(row=r, column=COL["B"]).value
        q = ws.cell(row=r, column=COL["Q"]).value
        if b == company and isinstance(q, datetime) and q.date() == contract_date.date():
            return True
    return False


def write_row(ws, row_num: int, data: dict):
    """파싱 결과를 엑셀 행에 기록"""
    r = row_num

    # A: =AC{r} 수식
    ws.cell(row=r, column=COL["A"]).value = f"=AC{r}"
    # B: 회사
    ws.cell(row=r, column=COL["B"]).value = data["company"]
    # C: 날짜
    ws.cell(row=r, column=COL["C"]).value = data["date"]
    # D: 출처
    ws.cell(row=r, column=COL["D"]).value = data["source"]
    # E: 기타
    ws.cell(row=r, column=COL["E"]).value = data["etc"] or None
    # F: 선종
    ws.cell(row=r, column=COL["F"]).value = data["vessel_type"] or None
    # G: 선주
    ws.cell(row=r, column=COL["G"]).value = data["buyer"] or None
    # H: Size
    ws.cell(row=r, column=COL["H"]).value = data["size"] or None
    # I: 척수
    ws.cell(row=r, column=COL["I"]).value = data["quantity"]
    # N: 인도년
    ws.cell(row=r, column=COL["N"]).value = data["delivery_year"]
    # O: 인도월
    ws.cell(row=r, column=COL["O"]).value = data["delivery_month"]
    # P: Builder (=B{r})
    ws.cell(row=r, column=COL["P"]).value = f"=B{r}"
    # Q: Contract Date
    ws.cell(row=r, column=COL["Q"]).value = data["contract_date"]
    # Z: Built date
    ws.cell(row=r, column=COL["Z"]).value = data["delivery_date"]
    # AA: 총계약금액(백만달러)
    ws.cell(row=r, column=COL["AA"]).value = data["amount_usd_mil"]
    # AB: 총계약금액(십억원) — 수식 또는 직접값
    if data["amount_usd_mil"] and data["exchange_rate"]:
        ws.cell(row=r, column=COL["AB"]).value = f"=AA{r}/AD{r}*1000"
    elif data["amount_krw_bil"]:
        ws.cell(row=r, column=COL["AB"]).value = data["amount_krw_bil"]
    # AC: 척당금액
    if data["quantity"] and data["amount_usd_mil"]:
        ws.cell(row=r, column=COL["AC"]).value = f"=AA{r}/I{r}"
    elif data["unit_price_usd"]:
        ws.cell(row=r, column=COL["AC"]).value = data["unit_price_usd"]
    # AD: 기준환율
    ws.cell(row=r, column=COL["AD"]).value = data["exchange_rate"]
    # AE: 확정여부
    ws.cell(row=r, column=COL["AE"]).value = data["confirmed"]
    # AF: 날짜 수식
    ws.cell(row=r, column=COL["AF"]).value = f'=YEAR(Q{r})&"."&MONTH(Q{r})'
    # AG: 상선/특수선
    ws.cell(row=r, column=COL["AG"]).value = data["vessel_category"]
    # AH: Clarksons
    ws.cell(row=r, column=COL["AH"]).value = data["clarksons"] or None


# ── 메인 ──────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="DART 단일판매·공급계약 → 엑셀 자동 입력")
    parser.add_argument("--api-key", required=True, help="DART OpenAPI 인증키")
    parser.add_argument("--days", type=int, default=7, help="조회 기간(일), 기본 7일")
    parser.add_argument("--excel", default=str(EXCEL_PATH), help="엑셀 파일 경로")
    args = parser.parse_args()

    end_date = datetime.today()
    start_date = end_date - timedelta(days=args.days)
    bgn_de = start_date.strftime("%Y%m%d")
    end_de = end_date.strftime("%Y%m%d")

    print(f"조회 기간: {bgn_de} ~ {end_de}")
    print(f"엑셀 파일: {args.excel}")

    wb = load_workbook(args.excel)
    ws = wb[SHEET_NAME]

    added = 0
    skipped = 0

    for company, corp_code in TARGET_COMPANIES.items():
        print(f"\n[{company}] 공시 조회 중...")
        disclosures = get_disclosure_list(args.api_key, corp_code, bgn_de, end_de)
        if not disclosures:
            print(f"  → 공시 없음")
            continue

        for item in disclosures:
            rcept_no = item.get("rcept_no", "")
            rcept_dt = item.get("rcept_dt", "")
            report_nm = item.get("report_nm", "")
            print(f"  공시: {report_nm} ({rcept_dt})")

            xml_text = get_disclosure_xml(args.api_key, rcept_no)
            if not xml_text:
                print(f"  → 원문 없음, 건너뜀")
                skipped += 1
                continue

            data = parse_disclosure(xml_text, company, rcept_dt)
            if data is None:
                skipped += 1
                continue

            # 중복 확인
            if already_exists(ws, company, data["contract_date"]):
                print(f"  → 이미 존재, 건너뜀")
                skipped += 1
                continue

            next_row = find_next_empty_row(ws)
            write_row(ws, next_row, data)
            added += 1
            print(f"  → {next_row}행에 추가 완료")

            time.sleep(0.5)  # API 호출 간격

    wb.save(args.excel)
    print(f"\n완료: {added}건 추가, {skipped}건 건너뜀")
    print(f"저장: {args.excel}")


if __name__ == "__main__":
    main()
