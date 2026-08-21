"""
조선사 단일판매·공급계약 공시 수집 → XLSX (기업별 시트)

사용법:
    pip install requests openpyxl

    python hd_heavy_orders.py --api-key YOUR_DART_KEY
    python hd_heavy_orders.py --api-key YOUR_DART_KEY --from-year 2020
    python hd_heavy_orders.py --api-key YOUR_DART_KEY --year 2026
    python hd_heavy_orders.py --api-key YOUR_DART_KEY --companies HD현대중공업 삼성중공업
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

DART_BASE  = "https://opendart.fss.or.kr/api"
SHEET_NAME = "뉴스수주"

# 기업명 → DART corp_code (공식 코드)
COMPANIES = {
    "HD현대중공업":   "01390344",
    "한화오션":       "00111704",
    "HD한국조선해양": "00164830",
    "삼성중공업":     "00126478",
    "HD현대미포":     "00164621",
}

DEFAULT_COMPANIES = ["HD현대중공업", "한화오션", "HD한국조선해양", "삼성중공업"]

COL = {
    "A": 1,  "B": 2,  "C": 3,  "D": 4,  "E": 5,
    "F": 6,  "G": 7,  "H": 8,  "I": 9,
    "N": 14, "O": 15, "P": 16, "Q": 17,
    "Z": 26, "AA": 27, "AB": 28, "AC": 29,
    "AD": 30, "AE": 31, "AF": 32, "AG": 33, "AH": 34,
}

VESSEL_MAP = [
    # 컨테이너
    ("컨테이너",   "컨테이너선"),
    # LNG
    ("LNG",        "LNG선"),
    ("LNGC",       "LNG선"),
    # LPG / 가스
    ("VLGC",       "LPG선"),
    ("VLAC",       "LPG선"),   # Very Large Ammonia Carrier
    ("MGC",        "LPG선"),
    ("LPG",        "LPG선"),
    ("암모니아",   "LPG선"),
    ("일반가스",   "LPG선"),
    ("가스운반",   "LPG선"),
    # 원유운반
    ("VLCC",       "원유운반선"),
    ("ULCC",       "원유운반선"),
    ("Suezmax",    "원유운반선"),
    ("suezmax",    "원유운반선"),
    ("Aframax",    "원유운반선"),
    ("aframax",    "원유운반선"),
    ("원유운반",   "원유운반선"),
    ("탱커",       "원유운반선"),
    # P/C선
    ("LR2",        "P/C선"),
    ("LR1",        "P/C선"),
    ("MR P/C",     "P/C선"),
    ("MR탱커",     "P/C선"),
    ("PC선",       "P/C선"),
    ("P/C",        "P/C선"),
    ("제품운반",   "P/C선"),
    ("석유화학",   "P/C선"),
    # 자동차운반
    ("PCTC",       "자동차운반선"),
    ("자동차운반", "자동차운반선"),
    ("PCC",        "자동차운반선"),
    # 벌크
    ("벌크",       "벌크선"),
    ("살물선",     "벌크선"),
    ("Capesize",   "벌크선"),
    ("Panamax",    "벌크선"),
    ("Handymax",   "벌크선"),
    # 해양 (구체적인 것 먼저, 마지막에 단독 "해양" 폴백)
    ("FPSO",       "해양"),
    ("FSRU",       "해양"),
    ("풍력",       "해양"),
    ("해양플랜트", "해양"),
    ("해양프로젝트","해양"),
    ("드릴십",     "해양"),
    ("drillship",  "해양"),
    ("offshore",   "해양"),
    # 특수선
    ("수상함",     "특수선"),
    ("호위함",     "특수선"),
    ("구축함",     "특수선"),
    ("잠수함",     "특수선"),
    ("군함",       "특수선"),
    ("쇄빙전용",   "특수선"),
    ("쇄빙",       "특수선"),
    ("함정",       "특수선"),
    ("특수선",     "특수선"),
    ("빙위",       "특수선"),
    # 엔진 / 기자재
    ("엔진",       "선박용엔진"),
    ("engine",     "선박용엔진"),
    ("발전기",     "선박용엔진"),
    ("발전설비",   "선박용엔진"),
    # 해양 단독 폴백 (마지막에 두어야 더 구체적인 것에 먼저 매칭)
    ("해양",       "해양"),
]


# ── DART API ──────────────────────────────────────────────────────────────

def lookup_corp_codes(api_key, names):
    """corpCode.xml에서 기업명으로 corp_code 검색. 결과 출력 후 반환."""
    import xml.etree.ElementTree as ET
    try:
        r = requests.get(f"{DART_BASE}/corpCode.xml",
                         params={"crtfc_key": api_key}, timeout=60)
        if r.content[:2] != b"PK":
            print(f"[오류] corpCode.xml 응답 이상: {r.text[:200]}")
            return {}
        import zipfile, io as _io
        with zipfile.ZipFile(_io.BytesIO(r.content)) as z:
            raw = z.read(z.namelist()[0])
        root = ET.fromstring(raw)
        result = {}
        for item in root.findall("list"):
            corp_name = item.findtext("corp_name", "")
            corp_code = item.findtext("corp_code", "")
            stock_code = item.findtext("stock_code", "")
            for tgt in names:
                if tgt in corp_name:
                    if corp_name not in result:
                        result[corp_name] = (corp_code, stock_code)
                        print(f"  {corp_name:25s}  corp_code={corp_code}  stock={stock_code}")
        return result
    except Exception as e:
        print(f"[오류] corp_code 조회 실패: {e}")
        return {}


def get_list(api_key, bgn_de, end_de, corp_code=None):
    if corp_code is None:
        corp_code = list(COMPANIES.values())[0]
    results, page = [], 1
    while True:
        try:
            r = requests.get(f"{DART_BASE}/list.json", params={
                "crtfc_key": api_key, "corp_code": corp_code,
                "bgn_de": bgn_de, "end_de": end_de,
                "page_count": 100, "page_no": page,
            }, timeout=15).json()
        except Exception as e:
            print(f"    [오류] {e}"); break
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


def _decode_html(raw_bytes):
    """UTF-8 BOM 우선, 대체문자 비율로 UTF-8/CP949 자동 선택"""
    if raw_bytes.startswith(b'\xef\xbb\xbf'):
        return raw_bytes[3:].decode('utf-8', errors='replace')
    # UTF-8로 우선 디코딩 후 대체문자(U+FFFD) 비율 확인
    utf8_text = raw_bytes.decode('utf-8', errors='replace')
    ratio = utf8_text.count('�') / max(len(utf8_text), 1)
    if ratio < 0.01:   # 1% 미만이면 UTF-8로 확정
        return utf8_text
    return raw_bytes.decode('cp949', errors='replace')


def get_html(api_key, rcept_no):
    try:
        r = requests.get(f"{DART_BASE}/document.xml",
                         params={"crtfc_key": api_key, "rcept_no": rcept_no}, timeout=20)
        if r.content[:2] == b"PK":
            with zipfile.ZipFile(io.BytesIO(r.content)) as z:
                for n in z.namelist():
                    raw = z.read(n)
                    return _decode_html(raw)
        return _decode_html(r.content)
    except Exception as e:
        print(f"    [오류] 문서 다운로드: {e}")
        return None


# ── HTML 파싱 ─────────────────────────────────────────────────────────────

def get_pairs(html):
    """
    DART HTML에서 레이블→값 매핑 추출.
    값은 class="xforms_input" span, 레이블은 직전 span.
    """
    spans = re.findall(r'<span([^>]*)>(.*?)</span>', html, re.DOTALL | re.IGNORECASE)
    items = []
    for attrs, content in spans:
        text = re.sub(r'<[^>]+>', '', content)
        text = re.sub(r'[\xa0　]', ' ', text)
        text = re.sub(r'\s+', ' ', text).strip()
        is_val = 'xforms_input' in attrs
        items.append((is_val, text))

    pairs = {}
    label_buf = []
    for is_val, text in items:
        if not is_val:
            if text:
                label_buf.append(text)
        else:
            if text and label_buf:
                # 가장 가까운 레이블과 매핑
                key = label_buf[-1]
                pairs[key] = text
                # 숫자값은 레이블 버퍼 비움, 텍스트값은 유지
            label_buf = []

    return pairs, items


def find_val(pairs, *labels):
    """매핑에서 레이블 키워드로 값 검색"""
    for label in labels:
        for k, v in pairs.items():
            if label in k and v and v not in ('-', '해당없음', 'N/A', '없음'):
                return v
    return ""


def _cell_text(html_frag):
    t = re.sub(r'<[^>]+>', ' ', html_frag)
    t = re.sub(r'[\xa0　]', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()


def parse_amendment_table(html):
    """정정공시 '항목|기존|정정' 테이블 파싱.
    Returns list of (field_name, before_text, after_text)
    """
    SKIP = {'항목', '기존', '정정', '변경', '내용', ''}
    results = []
    for row_html in re.findall(r'<tr[^>]*>(.*?)</tr>', html, re.DOTALL | re.IGNORECASE):
        cells = re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', row_html, re.DOTALL | re.IGNORECASE)
        if len(cells) < 3:
            continue
        field = _cell_text(cells[0])
        if field in SKIP or len(field) > 30:
            continue
        before = _cell_text(cells[1])
        # after: xforms_input 값 우선, 없으면 plain text
        after_inputs = re.findall(r'xforms_input[^>]*>([^<]*)<', cells[2])
        after_parts = [v.strip() for v in after_inputs if v.strip()]
        after = ' ~ '.join(after_parts) if after_parts else _cell_text(cells[2])
        if before and after and before != after and before not in SKIP:
            results.append((field, before, after))
    return results


def find_dates_from_html(html):
    """계약체결일자 시작/종료 날짜 추출 (xforms_input 날짜 패턴)"""
    # xforms_input span 안의 날짜 형식 값만 추출
    vals = re.findall(r'xforms_input[^>]*>([^<]*\d{4}-\d{2}-\d{2}[^<]*)<', html)
    dates = []
    for v in vals:
        v = v.strip()
        m = re.search(r'(\d{4}-\d{2}-\d{2})', v)
        if m:
            try:
                dates.append(datetime.strptime(m.group(1), "%Y-%m-%d"))
            except Exception:
                pass
    return dates


def extract_exchange_rate(html):
    """각주에서 'USD 1 = X,XXX.XX' 패턴으로 환율 추출"""
    m = re.search(r'USD\s*1\s*[=＝]\s*([\d,]+\.?\d*)', html)
    if m:
        try:
            return float(m.group(1).replace(',', ''))
        except Exception:
            pass
    return None


def detect_vessel(text):
    u = text.upper()
    for kw, label in VESSEL_MAP:
        if kw.upper() in u:
            return label
    return ""


def parse_krw_to_bil(text):
    """원화 텍스트 → 십억원"""
    digits = re.sub(r'[^\d]', '', text)
    if not digits:
        return None
    v = int(digits)
    if '조' in text:
        return round(v * 1000, 3)
    elif '억' in text:
        return round(v / 10, 3)
    elif '백만' in text:
        return round(v / 1000, 3)
    # 자릿수로 판단: 10자리 이상 → 원 단위
    if len(digits) >= 10:
        return round(v / 1_000_000_000, 3)
    elif len(digits) >= 7:
        return round(v / 1_000_000, 3)
    return None


def parse_usd_to_mil(text):
    """달러 텍스트 → 백만달러"""
    text2 = re.sub(r'[,\s]', '', text)
    m = re.search(r'[\d.]+', text2)
    if not m:
        return None
    v = float(m.group())
    if '억달러' in text or '억USD' in text.upper():
        return round(v * 100, 2)
    return round(v, 2)


def parse_date(text):
    """YYYY-MM-DD 또는 YYYY.MM.DD → datetime"""
    m = re.search(r'(\d{4})[-.](\d{1,2})[-.](\d{1,2})', text)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except Exception:
            pass
    return None


def find_original_dates(html):
    """기재정정 공시에서 기존(변경 전) 날짜 추출.
    1) '기존' 키워드 주변에서 날짜 범위 탐색
    2) td 셀 안에 날짜 2개 이상인 셀 탐색 (기존 계약기간)
    3) 변경 후(xforms_input) 날짜와 다른 날짜만 수집
    """
    # xforms_input(변경 후) 날짜를 제거한 plain HTML
    plain = re.sub(r'<span[^>]*xforms_input[^>]*>.*?</span>', '', html, flags=re.DOTALL)
    plain_text = re.sub(r'<[^>]+>', ' ', plain)

    new_dates = set()
    for m in re.finditer(r'(\d{4}-\d{2}-\d{2})', re.sub(r'<[^>]+>', ' ',
            re.sub(r'.*?<body', '', html, flags=re.DOTALL))):
        try:
            d = datetime.strptime(m.group(1), "%Y-%m-%d")
            if 2000 <= d.year <= 2040:
                new_dates.add(d)
        except Exception:
            pass
    # 변경 후 날짜에서 plain_text에 있는 날짜를 제거해 기존 날짜만 남김
    # (xforms_input 제거 후 남은 날짜가 진짜 "기존" 날짜)

    # 1) '기존' 키워드 뒤 날짜 우선 탐색
    m = re.search(
        r'기존[^0-9]{0,80}(\d{4}-\d{2}-\d{2})[^0-9]{0,30}(\d{4}-\d{2}-\d{2})',
        plain_text)
    if m:
        try:
            d1 = datetime.strptime(m.group(1), "%Y-%m-%d")
            d2 = datetime.strptime(m.group(2), "%Y-%m-%d")
            if 2000 <= d1.year <= 2040 and 2000 <= d2.year <= 2040:
                return [d1, d2]
        except Exception:
            pass

    # 2) plain td 셀 안에 날짜 쌍이 있는 셀 탐색
    for cell_html in re.findall(r'<td[^>]*>(.*?)</td>', plain, re.DOTALL):
        cell_text = re.sub(r'<[^>]+>', ' ', cell_html)
        ds = re.findall(r'(\d{4}-\d{2}-\d{2})', cell_text)
        if len(ds) >= 2:
            try:
                d1 = datetime.strptime(ds[0], "%Y-%m-%d")
                d2 = datetime.strptime(ds[1], "%Y-%m-%d")
                if 2000 <= d1.year <= 2040 and 2000 <= d2.year <= 2040:
                    return [d1, d2]
            except Exception:
                pass

    # 3) fallback: plain_text 전체에서 날짜 수집, xforms 날짜와 다른 것만
    dates = []
    for m in re.finditer(r'(\d{4}-\d{2}-\d{2})', plain_text):
        try:
            d = datetime.strptime(m.group(1), "%Y-%m-%d")
            if 2000 <= d.year <= 2040 and d not in new_dates:
                dates.append(d)
        except Exception:
            pass
    return dates[:2]


def parse(html, rcept_dt, report_nm, corp_name=""):
    pairs, items = get_pairs(html)

    is_amendment = "[기재정정]" in report_nm or "(정정)" in report_nm

    row = {
        "date":               None,
        "contract_date":      None,
        "start_date":         None,
        "end_date":           None,
        "orig_start_date":    None,   # 정정 전 시작일
        "orig_end_date":      None,   # 정정 전 종료일
        "amend_fields":       [],     # 정정항목 리스트 [(field, before, after)]
        "buyer":              "",
        "vessel_type":        "",
        "vessel_category":    "상선",
        "contract_nm":        "",     # 체결계약명 (선종/척수 단서)
        "etc":                "",
        "size":               "",
        "quantity":           None,
        "delivery_year":      None,
        "delivery_month":     None,
        "delivery_date":      None,
        "amount_usd_mil":     None,
        "amount_krw_bil":     None,
        "unit_price_usd":     None,
        "exchange_rate":      None,
        "confirmed":          "O",
        "clarksons":          "",
        "is_amendment":       is_amendment,
        "report_nm":          report_nm,
        "corp_name":          corp_name,
    }

    # 공시 접수일 → 기본 날짜
    if len(rcept_dt) == 8:
        try:
            dt = datetime(int(rcept_dt[:4]), int(rcept_dt[4:6]), int(rcept_dt[6:]))
            row["date"] = dt
            row["contract_date"] = dt
        except Exception:
            pass

    # ── 시작일 / 종료일: xforms_input 날짜값 순서대로 추출 ────────────
    dates = find_dates_from_html(html)
    if len(dates) >= 1:
        row["start_date"]    = dates[0]
        row["contract_date"] = dates[0]
        row["date"]          = dates[0]
    if len(dates) >= 2:
        row["end_date"]       = dates[1]
        row["delivery_year"]  = dates[1].year
        row["delivery_month"] = dates[1].month
        row["delivery_date"]  = dates[1]

    # ── 기재정정: 항목/기존/정정 테이블 파싱 ────────────────────────
    if is_amendment:
        amend_rows = parse_amendment_table(html)
        row["amend_fields"] = amend_rows

        for field, before, after in amend_rows:
            fl = field  # 필드명 소문자 비교용
            # 계약기간 변경
            if any(k in fl for k in ["계약기간", "납품기간", "이행기간"]):
                # before: "YYYY-MM-DD ~ YYYY-MM-DD" or single date
                ds = re.findall(r'(\d{4}-\d{2}-\d{2})', before)
                if len(ds) >= 2:
                    try: row["orig_start_date"] = datetime.strptime(ds[0], "%Y-%m-%d")
                    except: pass
                    try: row["orig_end_date"]   = datetime.strptime(ds[1], "%Y-%m-%d")
                    except: pass
                elif len(ds) == 1:
                    try: row["orig_end_date"]   = datetime.strptime(ds[0], "%Y-%m-%d")
                    except: pass
            # 계약종료일만 변경
            elif any(k in fl for k in ["종료일", "완료일", "납기"]):
                ds = re.findall(r'(\d{4}-\d{2}-\d{2})', before)
                if ds:
                    try: row["orig_end_date"] = datetime.strptime(ds[0], "%Y-%m-%d")
                    except: pass
            # 계약시작일만 변경
            elif any(k in fl for k in ["시작일", "착수일"]):
                ds = re.findall(r'(\d{4}-\d{2}-\d{2})', before)
                if ds:
                    try: row["orig_start_date"] = datetime.strptime(ds[0], "%Y-%m-%d")
                    except: pass
            # 계약상대방 변경
            elif any(k in fl for k in ["계약상대", "거래상대", "발주처", "상대방"]):
                if not row["buyer"]:
                    row["buyer"] = after  # 변경 후 값이 현재 값
            # 계약금액 변경
            elif any(k in fl for k in ["계약금액", "공급금액", "총계약"]):
                if "달러" in after or "USD" in after.upper():
                    v = parse_usd_to_mil(after)
                    if v: row["amount_usd_mil"] = v
                elif any(k in after for k in ["원", "억", "조"]):
                    v = parse_krw_to_bil(after)
                    if v: row["amount_krw_bil"] = v

        # fallback: xforms_input 날짜에서 orig 날짜 추출
        if not row["orig_start_date"] and not row["orig_end_date"]:
            orig_dates = find_original_dates(html)
            if len(orig_dates) >= 1:
                row["orig_start_date"] = orig_dates[0]
            if len(orig_dates) >= 2:
                row["orig_end_date"] = orig_dates[1]

    # ── 체결계약명 / 납품물목 ─────────────────────────────────────────
    contract_nm = find_val(pairs,
        "체결계약명", "계약명", "계약건명",
        "납품물목", "납품명세", "납품사항", "납품목록",
        "공급물품", "거래내용", "물품내용", "서비스내용")
    row["contract_nm"] = contract_nm
    row["etc"] = contract_nm

    # ── 척수 추출 (우선순위: pairs 레이블 > 계약명 > 전체 xforms_input > HTML)
    qty_from_label = find_val(pairs, "계약수량", "선박수", "납품수량", "수량", "척수", "건조척수",
                              "납품물목", "납품명세", "납품사항")
    if qty_from_label:
        m_qty = re.search(r'(\d+)', qty_from_label)
        if m_qty:
            row["quantity"] = int(m_qty.group(1))

    if not row["quantity"]:
        # 계약명에서 "N척" 패턴
        m_qty = re.search(r'(\d+)\s*척', contract_nm)
        if m_qty:
            row["quantity"] = int(m_qty.group(1))

    if not row["quantity"]:
        # 전체 xforms_input 값에서 "N척" 패턴
        for is_v, text in items:
            if is_v:
                m_qty = re.search(r'(\d+)\s*척', text)
                if m_qty:
                    row["quantity"] = int(m_qty.group(1))
                    break

    if not row["quantity"]:
        # HTML 전체에서 "N척" 패턴 (최후 수단, 전체 탐색)
        m_qty = re.search(r'(\d+)\s*척', html)
        if m_qty:
            row["quantity"] = int(m_qty.group(1))

    if not row["quantity"]:
        # 엔진/기자재: "N기", "N대", "N세트" 패턴 (체결계약명 우선)
        for src in [contract_nm, " ".join(t for iv, t in items if iv), html[:5000]]:
            m_qty = re.search(r'(\d+)\s*(기|대|세트)', src)
            if m_qty:
                row["quantity"] = int(m_qty.group(1))
                break

    # ── 계약상대방(선주) ───────────────────────────────────────────────
    _SKIP_VALS = {'-', '해당없음', 'N/A', '없음', '', 'N/a'}

    def _clean_buyer(v):
        v = re.sub(r'[\r\n]+', ' ', v)
        v = re.sub(r'\s+', ' ', v).strip().rstrip('.')
        return v if len(v) >= 2 else ""

    buyer = find_val(pairs, "계약상대", "계약상대방", "거래상대방", "발주처", "매수인",
                     "계약 상대방", "상대방", "거래처", "수요자", "납품처")
    buyer = _clean_buyer(buyer) if buyer else ""

    if not buyer:
        # td 기반 탐색: 레이블 td 다음 td에서 값 추출
        for lbl in ["계약상대방", "거래상대방", "발주처", "매수인"]:
            pat = (rf'>{re.escape(lbl)}</td>\s*'
                   r'(?:<td[^>]*>)(.*?)(?:</td>)')
            m_b = re.search(pat, html, re.DOTALL | re.IGNORECASE)
            if m_b:
                val = _clean_buyer(_cell_text(m_b.group(1)))
                if val and val not in _SKIP_VALS:
                    buyer = val
                    break

    if not buyer:
        # xforms_input 값 중 회사명처럼 보이는 값 (날짜·숫자 아닌 텍스트, 2자 이상)
        for is_v, text in items:
            if not is_v:
                continue
            t = text.strip()
            if (len(t) >= 2
                    and not re.match(r'^[\d,\s.%]+$', t)
                    and not re.search(r'\d{4}-\d{2}-\d{2}', t)
                    and t not in _SKIP_VALS
                    # 앞에서 이미 날짜/금액/계약명으로 매핑된 값은 제외
                    and t not in (contract_nm, row.get("contract_nm", ""))):
                # 이 값의 레이블이 "계약상대" 계열인지 items에서 역추적
                pass  # pairs 기반으로 이미 충분하므로 여기선 생략

    if not buyer:
        # HTML plain text 에서 레이블 뒤 첫 비어있지 않은 줄
        all_text_plain = re.sub(r'<[^>]+>', '\n', html)
        for lbl in ["계약상대방", "거래상대방", "발주처"]:
            m_b = re.search(
                rf'{re.escape(lbl)}\s*\n\s*([^\n]{{2,120}})', all_text_plain)
            if m_b:
                val = _clean_buyer(m_b.group(1))
                if val and val not in _SKIP_VALS:
                    buyer = val
                    break

    row["buyer"] = buyer

    # ── 계약금액 ──────────────────────────────────────────────────────
    amt_str = find_val(pairs, "계약금액", "총계약금액", "공급금액")
    if amt_str:
        if "달러" in amt_str or "USD" in amt_str.upper():
            row["amount_usd_mil"] = parse_usd_to_mil(amt_str)
        else:
            row["amount_krw_bil"] = parse_krw_to_bil(amt_str)
    else:
        # xforms_input 중 큰 숫자값 (10자리 이상) 직접 추출
        for is_val, text in items:
            if is_val and re.match(r'^[\d,]+$', text.replace(' ', '')):
                digits = re.sub(r'[^\d]', '', text)
                if len(digits) >= 9:
                    row["amount_krw_bil"] = parse_krw_to_bil(text)
                    break

    # ── 기준환율: 각주 "USD 1 = X,XXX.XX" 패턴 ──────────────────────
    row["exchange_rate"] = extract_exchange_rate(html)

    # 환율 역산: KRW/USD 둘 다 있으면
    if not row["exchange_rate"] and row["amount_krw_bil"] and row["amount_usd_mil"]:
        row["exchange_rate"] = round(row["amount_krw_bil"] * 1000 / row["amount_usd_mil"], 1)

    # ── KRW만 있고 환율 있으면 USD 역산 ──────────────────────────────
    if row["amount_krw_bil"] and not row["amount_usd_mil"] and row["exchange_rate"]:
        row["amount_usd_mil"] = round(row["amount_krw_bil"] * 1000 / row["exchange_rate"], 3)

    # USD만 있고 환율 있으면 KRW 역산
    if row["amount_usd_mil"] and not row["amount_krw_bil"] and row["exchange_rate"]:
        row["amount_krw_bil"] = round(row["amount_usd_mil"] * row["exchange_rate"] / 1000, 3)

    # ── 선종 탐지 ─────────────────────────────────────────────────────
    # pairs에서 선종/선박종류 레이블 직접 검색
    vessel_label_val = find_val(pairs, "선종", "선박종류", "선박유형", "물품명", "선박명", "공급물품", "거래내용")
    all_vals = " ".join(t for is_v, t in items if is_v)
    all_text = re.sub(r'<[^>]+>', ' ', html)

    vtype = (detect_vessel(vessel_label_val)
             or detect_vessel(contract_nm)
             or detect_vessel(row["buyer"])
             or detect_vessel(all_vals)
             or detect_vessel(all_text[:8000]))
    row["vessel_type"] = vtype
    if vtype in ("해양", "특수선", "선박용엔진"):
        row["vessel_category"] = vtype

    # ── 척당 금액 ─────────────────────────────────────────────────────
    if row["quantity"] and row["amount_usd_mil"]:
        row["unit_price_usd"] = round(row["amount_usd_mil"] / row["quantity"], 3)
    elif row["amount_usd_mil"] and not row["quantity"]:
        # 척수 없으면 총액 = 척당금액으로 처리
        row["unit_price_usd"] = row["amount_usd_mil"]

    return row


# ── 학습 데이터 로드 (수주학습용.xlsx 크로스체크) ─────────────────────

def load_reference(excel_path):
    """엑셀에서 HD현대중공업 Dart 행 추출 → {날짜: 데이터} 매핑"""
    wb = load_workbook(excel_path, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        return {}
    ws = wb[SHEET_NAME]
    ref = {}
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
        if row[1] != CORP_NAME or row[3] != "Dart":
            continue
        cdate = row[16]  # Q: Contract Date
        if not (cdate and hasattr(cdate, 'date')):
            continue
        key = cdate.date().isoformat()
        ref[key] = {
            "etc":           row[4],
            "vessel_type":   row[5],
            "buyer":         row[6],
            "quantity":      row[8],
            "delivery_year": row[13],
            "delivery_month":row[14],
            "amount_usd":    row[26],
            "amount_krw":    row[27],
            "unit_price":    row[28],
            "exchange_rate": row[29],
        }
    return ref


def apply_reference(data, ref):
    """학습 데이터로 빈 필드 보완"""
    if not data["contract_date"]:
        return data
    key = data["contract_date"].date().isoformat()
    r = ref.get(key)
    if not r:
        return data

    # 빈 필드만 학습 데이터로 채움
    if not data["vessel_type"]    and r["vessel_type"]:   data["vessel_type"]    = r["vessel_type"]
    if not data["buyer"]          and r["buyer"]:          data["buyer"]          = r["buyer"]
    if not data["etc"]            and r["etc"]:            data["etc"]            = r["etc"]
    if not data["quantity"]       and r["quantity"]:       data["quantity"]       = r["quantity"]
    if not data["delivery_year"]  and r["delivery_year"]:  data["delivery_year"]  = r["delivery_year"]
    if not data["delivery_month"] and r["delivery_month"]: data["delivery_month"] = r["delivery_month"]
    if not data["amount_usd_mil"] and r["amount_usd"]:     data["amount_usd_mil"] = r["amount_usd"]
    if not data["amount_krw_bil"] and r["amount_krw"]:     data["amount_krw_bil"] = r["amount_krw"]
    if not data["exchange_rate"]  and r["exchange_rate"]:  data["exchange_rate"]  = r["exchange_rate"]
    if not data["unit_price_usd"] and r["unit_price"]:     data["unit_price_usd"] = r["unit_price"]

    # delivery_date 재생성
    if data["delivery_year"] and data["delivery_month"] and not data["delivery_date"]:
        try:
            y, mo = data["delivery_year"], data["delivery_month"]
            last = calendar.monthrange(y, mo)[1]
            data["delivery_date"] = datetime(y, mo, last)
        except Exception:
            pass

    return data


# ── 엑셀 쓰기 ─────────────────────────────────────────────────────────────

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
    # A열: 척당금액 참조
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


# ── XLSX 저장 (시트 여러 개 유지 가능) ────────────────────────────────────

XLSX_HEADERS = [
    "정정여부", "회사", "수주일자", "계약시작일", "계약종료일",
    "정정전시작일", "정정전종료일", "정정항목", "정정전값", "정정후값",
    "체결계약명", "계약상대", "선종", "척수", "인도년", "인도월",
    "금액(원화,십억원)", "금액(달러,백만)", "척당금액(달러,백만)",
    "기준환율", "확정", "공시제목",
]
XLSX_WIDTHS = [8,14,12,12,12,12,12,24,30,30,40,22,12,6,8,8,14,14,16,10,6,45]


def _make_row_vals(d, fmt):
    afields = d.get("amend_fields") or []
    return [
        "정정" if d["is_amendment"] else "",
        d.get("corp_name") or "",
        fmt(d.get("contract_date")),
        fmt(d.get("start_date")),
        fmt(d.get("end_date")),
        fmt(d.get("orig_start_date")),
        fmt(d.get("orig_end_date")),
        ", ".join(f for f, _, _ in afields),
        " / ".join(b for _, b, _ in afields),
        " / ".join(a for _, _, a in afields),
        d.get("contract_nm") or d.get("etc") or "",
        d.get("buyer") or "",
        d.get("vessel_type") or "",
        d.get("quantity") or "",
        d.get("delivery_year") or "",
        d.get("delivery_month") or "",
        d.get("amount_krw_bil") or "",
        d.get("amount_usd_mil") or "",
        d.get("unit_price_usd") or "",
        d.get("exchange_rate") or "",
        d.get("confirmed") or "",
        d.get("report_nm") or "",
    ]


def _write_sheet(ws, records, Font, PatternFill, Alignment):
    for ci, h in enumerate(XLSX_HEADERS, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDEEFF")
        cell.alignment = Alignment(horizontal="center")

    def fmt(d):
        return d.strftime("%Y-%m-%d") if d else ""

    for ri, d in enumerate(records, 2):
        for ci, v in enumerate(_make_row_vals(d, fmt), 1):
            ws.cell(row=ri, column=ci, value=v)

    for ci, w in enumerate(XLSX_WIDTHS, 1):
        ws.column_dimensions[ws.cell(row=1, column=ci).column_letter].width = w


def save_xlsx(records, path):
    """records 안에 여러 기업 포함 가능 → 기업별 시트 + 전체 시트"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    # 기업별 시트
    from collections import defaultdict
    by_corp = defaultdict(list)
    for r in records:
        by_corp[r.get("corp_name", "기타")].append(r)

    first = True
    for corp_nm, recs in by_corp.items():
        if first:
            ws = wb.active
            ws.title = corp_nm[:31]
            first = False
        else:
            ws = wb.create_sheet(title=corp_nm[:31])
        _write_sheet(ws, recs, Font, PatternFill, Alignment)

    # 전체 통합 시트 (기업 2개 이상일 때)
    if len(by_corp) > 1:
        ws_all = wb.create_sheet(title="전체")
        _write_sheet(ws_all, records, Font, PatternFill, Alignment)

    wb.save(path)
    print(f"XLSX 저장: {path}  ({len(records)}건)")


# ── CSV ────────────────────────────────────────────────────────────────────

def save_csv(records, path, encoding="euc-kr"):
    fields = [
        "정정여부", "회사", "수주일자", "계약시작일", "계약종료일",
        "정정전시작일", "정정전종료일", "정정항목", "정정전값", "정정후값",
        "체결계약명", "계약상대", "선종", "척수",
        "인도년", "인도월",
        "금액(원화,십억원)", "금액(달러,백만)", "척당금액(달러,백만)",
        "기준환율", "확정", "공시제목",
    ]

    def fmt(d):
        return d.strftime("%Y-%m-%d") if d else ""

    with open(path, "w", newline="", encoding=encoding) as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for d in records:
            afields = d.get("amend_fields") or []
            w.writerow({
                "정정여부":         "정정" if d["is_amendment"] else "",
                "회사":            CORP_NAME,
                "수주일자":         fmt(d.get("contract_date")),
                "계약시작일":       fmt(d.get("start_date")),
                "계약종료일":       fmt(d.get("end_date")),
                "정정전시작일":     fmt(d.get("orig_start_date")),
                "정정전종료일":     fmt(d.get("orig_end_date")),
                "정정항목":         ", ".join(f for f, _, _ in afields),
                "정정전값":         " / ".join(b for _, b, _ in afields),
                "정정후값":         " / ".join(a for _, _, a in afields),
                "체결계약명":       d.get("contract_nm") or d.get("etc") or "",
                "계약상대":         d.get("buyer") or "",
                "선종":            d.get("vessel_type") or "",
                "척수":            d.get("quantity") or "",
                "인도년":           d.get("delivery_year") or "",
                "인도월":           d.get("delivery_month") or "",
                "금액(원화,십억원)": d.get("amount_krw_bil") or "",
                "금액(달러,백만)":   d.get("amount_usd_mil") or "",
                "척당금액(달러,백만)": d.get("unit_price_usd") or "",
                "기준환율":         d.get("exchange_rate") or "",
                "확정":            d.get("confirmed") or "",
                "공시제목":         d.get("report_nm") or "",
            })
    print(f"CSV 저장: {path}  ({len(records)}건)")


# ── 메인 ──────────────────────────────────────────────────────────────────

def _collect_one_company(api_key, corp_name, corp_code, ranges, ref, args):
    """단일 기업 공시 수집 → records 리스트 반환"""
    records = []
    for bgn, end_d in ranges:
        print(f"    {bgn[:4]}년 조회...")
        items = get_list(api_key, bgn, end_d, corp_code)
        print(f"      공시 {len(items)}건")

        for item in items:
            rcept_no  = item.get("rcept_no", "")
            rcept_dt  = item.get("rcept_dt", "")
            report_nm = item.get("report_nm", "").strip()

            html = get_html(api_key, rcept_no)
            if not html:
                continue

            data = parse(html, rcept_dt, report_nm, corp_name)

            if ref:
                data = apply_reference(data, ref)

            amend_tag = "[정정]" if data["is_amendment"] else ""
            cdate = data["contract_date"].strftime("%Y-%m-%d") if data["contract_date"] else "?"
            edate = data["end_date"].strftime("%Y-%m-%d") if data["end_date"] else "?"
            print(f"        {amend_tag}{cdate}~{edate} | {data['vessel_type'] or '-'} | {data['buyer'] or '-'} | {data['quantity'] or '?'}척 | USD{data['amount_usd_mil'] or ''} KRW{data['amount_krw_bil'] or ''}")

            # --debug: 공란 원인 분석
            if getattr(args, 'debug', False):
                blanks = []
                if not data['vessel_type']:   blanks.append("선종")
                if not data['quantity']:       blanks.append("척수")
                if not data['amount_usd_mil'] and not data['amount_krw_bil']:
                                               blanks.append("금액")
                if not data['exchange_rate']:  blanks.append("환율")
                if not data['buyer']:          blanks.append("선주")
                if blanks:
                    pairs, items2 = get_pairs(html)
                    print(f"        ▶ 공란필드: {', '.join(blanks)}")
                    print(f"        ▶ 계약명: {data['etc'] or '(없음)'}")
                    print(f"        ▶ 파싱된 레이블-값 쌍 (전체):")
                    for k, v in pairs.items():
                        print(f"            [{k}] = {v}")
                    if "선주" in blanks:
                        print(f"        ▶ [BUYER 디버그] pairs 키 전체:")
                        for k in pairs.keys():
                            print(f"            키: {repr(k)}")
                        # td 기반 탐색 시도 결과 출력
                        print(f"        ▶ [BUYER 디버그] td 기반 탐색:")
                        for lbl in ["계약상대방", "거래상대방", "발주처", "매수인"]:
                            pat = (rf'>{re.escape(lbl)}</td>\s*'
                                   r'(?:<td[^>]*>)(.*?)(?:</td>)')
                            m_b = re.search(pat, html, re.DOTALL | re.IGNORECASE)
                            if m_b:
                                print(f"            td[{lbl}] → {_cell_text(m_b.group(1))!r}")
                            else:
                                print(f"            td[{lbl}] → 없음")
                        # HTML plain text에서 레이블 탐색
                        print(f"        ▶ [BUYER 디버그] HTML plaintext 검색:")
                        all_plain = re.sub(r'<[^>]+>', '\n', html)
                        for lbl in ["계약상대방", "거래상대방", "발주처"]:
                            m_b = re.search(
                                rf'{re.escape(lbl)}\s*\n\s*([^\n]{{2,120}})', all_plain)
                            if m_b:
                                print(f"            plain[{lbl}] → {m_b.group(1).strip()!r}")
                            else:
                                print(f"            plain[{lbl}] → 없음")
                        # xforms_input 값 중 회사명 후보
                        print(f"        ▶ [BUYER 디버그] xforms_input 값 전체:")
                        for is_v, txt in items2:
                            if is_v and txt:
                                print(f"            val: {txt!r}")

            records.append(data)
            time.sleep(0.5)

        time.sleep(0.3)

    return records


def main():
    ap = argparse.ArgumentParser(description="조선사 단일판매·공급계약 공시 수집")
    ap.add_argument("--api-key",      required=True)
    ap.add_argument("--companies",    nargs="+", default=DEFAULT_COMPANIES,
                    metavar="COMPANY",
                    help=f"수집 기업명 (기본: {' '.join(DEFAULT_COMPANIES)})")
    ap.add_argument("--from-year",    type=int, default=2020,
                    help="수집 시작 연도 (기본: 2020)")
    ap.add_argument("--year",         type=int, default=0,
                    help="특정 연도만 수집 (예: --year 2026)")
    ap.add_argument("--reference",    default="", help="크로스체크용 학습 엑셀")
    ap.add_argument("--csv-out",      default="", help="CSV 저장 경로 (생략 시 저장 안 함)")
    ap.add_argument("--xlsx-out",     default="조선사_수주.xlsx", help="XLSX 저장 경로")
    ap.add_argument("--encoding",     default="euc-kr", help="CSV 인코딩 (기본: euc-kr)")
    ap.add_argument("--debug",        action="store_true", help="공란 원인 분석 출력")
    ap.add_argument("--lookup-corps", action="store_true",
                    help="DART에서 기업 corp_code 검색 후 종료 (코드 확인용)")
    args = ap.parse_args()

    # corp_code 조회 모드
    if args.lookup_corps:
        search_names = args.companies if args.companies else list(COMPANIES.keys())
        print(f"DART corp_code 검색: {search_names}")
        lookup_corp_codes(args.api_key, search_names)
        return

    # 학습 데이터 로드
    ref = {}
    if args.reference and Path(args.reference).exists():
        ref = load_reference(args.reference)
        print(f"학습 데이터: {len(ref)}건 로드\n")

    today = datetime.today()

    # 조회 구간 계산
    if args.year:
        ranges = [(f"{args.year}0101", f"{args.year}1231")]
        period_str = f"{args.year}년"
    else:
        from_year = args.from_year
        cur = datetime(from_year, 1, 1)
        ranges = []
        while cur < today:
            nxt = min(cur + timedelta(days=364), today)
            ranges.append((cur.strftime("%Y%m%d"), nxt.strftime("%Y%m%d")))
            cur = nxt + timedelta(days=1)
        period_str = f"{from_year}년~{today.year}년"

    # 기업 목록 확인
    target_corps = []
    for name in args.companies:
        code = COMPANIES.get(name)
        if not code:
            print(f"[경고] 알 수 없는 기업: {name}  (등록된 기업: {', '.join(COMPANIES.keys())})")
            continue
        target_corps.append((name, code))

    if not target_corps:
        print("[오류] 수집할 기업이 없습니다."); return

    print(f"수집 기간: {period_str}")
    print(f"수집 기업: {', '.join(n for n, _ in target_corps)}")
    print(f"조회 구간: {len(ranges)}개\n")

    all_records = []
    for corp_name, corp_code in target_corps:
        print(f"── {corp_name} ({corp_code}) ──────────────────")
        recs = _collect_one_company(args.api_key, corp_name, corp_code, ranges, ref, args)
        print(f"  → {corp_name} {len(recs)}건 수집\n")
        all_records.extend(recs)

    print(f"총 {len(all_records)}건 수집  (정정공시: {sum(1 for r in all_records if r['is_amendment'])}건)")

    if args.xlsx_out:
        save_xlsx(all_records, args.xlsx_out)
    if args.csv_out:
        save_csv(all_records, args.csv_out, args.encoding)


if __name__ == "__main__":
    main()
