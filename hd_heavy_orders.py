"""
HD현대중공업 단일판매·공급계약 공시 수집 → CSV/Excel

사용법:
    pip install requests openpyxl
    python hd_heavy_orders.py --api-key YOUR_DART_KEY
    python hd_heavy_orders.py --api-key YOUR_DART_KEY --years 6
"""

import argparse
import csv
import io
import re
import time
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

import requests

CORP_CODE  = "00164788"   # HD현대중공업
CORP_NAME  = "HD현대중공업"
DART_BASE  = "https://opendart.fss.or.kr/api"


def get_list(api_key, bgn_de, end_de):
    results, page = [], 1
    while True:
        r = requests.get(f"{DART_BASE}/list.json", params={
            "crtfc_key": api_key, "corp_code": CORP_CODE,
            "pblntf_detail_ty": "B002",
            "bgn_de": bgn_de, "end_de": end_de,
            "page_count": 100, "page_no": page,
        }, timeout=15).json()
        if r.get("status") not in ("000",):
            break
        results.extend(r.get("list", []))
        if page * 100 >= int(r.get("total_count", 0)):
            break
        page += 1
        time.sleep(0.3)
    return results


def get_xml(api_key, rcept_no):
    r = requests.get(f"{DART_BASE}/document.xml",
                     params={"crtfc_key": api_key, "rcept_no": rcept_no}, timeout=20)
    if r.content[:2] == b"PK":
        with zipfile.ZipFile(io.BytesIO(r.content)) as z:
            for n in z.namelist():
                if n.endswith(".xml"):
                    return z.read(n).decode("utf-8", errors="ignore")
    return r.text


def find(xml, *keys):
    for k in keys:
        m = re.search(rf"{re.escape(k)}[^<]*</[^>]+>\s*<[^>]+>([^<]+)</", xml)
        if m:
            v = re.sub(r"\s+", " ", m.group(1)).strip()
            if v and v not in ("-", "해당없음", "N/A"):
                return v
    return ""


def parse(xml, rcept_dt, report_nm):
    row = {
        "회사": CORP_NAME, "공시제목": report_nm,
        "날짜": "", "선주": "", "선종": "",
        "척수": "", "인도년": "", "인도월": "",
        "금액(원화,십억원)": "", "금액(달러,백만)": "",
        "확정": "O",
    }
    if len(rcept_dt) == 8:
        try:
            row["날짜"] = f"{rcept_dt[:4]}-{rcept_dt[4:6]}-{rcept_dt[6:]}"
        except Exception:
            pass

    row["선주"] = find(xml, "계약상대방", "거래상대방", "발주처", "매수인")

    amt = find(xml, "계약금액", "총계약금액", "공급금액")
    if amt:
        v_str = re.sub(r",", "", amt)
        v_m = re.search(r"[\d.]+", v_str)
        if v_m:
            v = float(v_m.group())
            if "달러" in amt or "USD" in amt.upper():
                if "억" in amt:
                    row["금액(달러,백만)"] = round(v * 100, 1)
                else:
                    row["금액(달러,백만)"] = round(v, 1)
            elif "조" in amt:
                row["금액(원화,십억원)"] = round(v * 1000, 1)
            elif "억" in amt:
                row["금액(원화,십억원)"] = round(v / 10, 1)
            elif "백만원" in amt:
                row["금액(원화,십억원)"] = round(v / 1000, 1)

    purpose = find(xml, "계약내용", "공급내용", "계약목적물", "품목", "선박종류")
    stx = purpose + xml[:4000]
    for kw, label in [
        ("컨테이너", "컨테이너선"), ("LNG", "LNG선"), ("LPG", "LPG선"),
        ("암모니아", "LPG선"), ("VLCC", "VLCC"), ("원유운반", "원유운반선"),
        ("PC선", "P/C선"), ("제품운반", "P/C선"), ("MR탱커", "P/C선"),
        ("벌크", "벌크선"), ("풍력", "해양"), ("FPSO", "해양"),
        ("해양플랜트", "해양"), ("특수선", "특수선"),
    ]:
        if kw.upper() in stx.upper():
            row["선종"] = label
            break

    qty = find(xml, "수량", "척수", "호선수")
    if qty:
        m = re.search(r"\d+", qty)
        if m:
            row["척수"] = int(m.group())

    dlv = find(xml, "납기", "인도예정", "납품예정", "인도일")
    if dlv:
        m = re.search(r"(\d{4})[.\-년]\s*(\d{1,2})", dlv)
        if m:
            row["인도년"] = int(m.group(1))
            row["인도월"] = int(m.group(2))

    return row


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--api-key", required=True)
    ap.add_argument("--years", type=int, default=6, help="과거 몇 년치 (기본 6)")
    ap.add_argument("--out", default="HD현대중공업_수주.csv")
    args = ap.parse_args()

    today = datetime.today()
    start = today.replace(year=today.year - args.years, month=1, day=1)

    # 1년 단위 분할
    ranges, cur = [], start
    while cur < today:
        nxt = min(cur + timedelta(days=364), today)
        ranges.append((cur.strftime("%Y%m%d"), nxt.strftime("%Y%m%d")))
        cur = nxt + timedelta(days=1)

    print(f"HD현대중공업 {args.years}년치 수주 수집 ({start.strftime('%Y-%m-%d')} ~ {today.strftime('%Y-%m-%d')})")
    print(f"조회 구간: {len(ranges)}개\n")

    records = []
    for bgn, end in ranges:
        print(f"  {bgn} ~ {end} 조회 중...")
        items = get_list(args.api_key, bgn, end)
        print(f"    공시 {len(items)}건 발견")
        for item in items:
            rcept_no  = item.get("rcept_no", "")
            rcept_dt  = item.get("rcept_dt", "")
            report_nm = item.get("report_nm", "")
            print(f"      파싱: {report_nm} ({rcept_dt})")
            xml = get_xml(args.api_key, rcept_no)
            if xml:
                records.append(parse(xml, rcept_dt, report_nm))
            time.sleep(0.5)
        time.sleep(0.3)

    if not records:
        print("\n수집된 데이터 없음")
        return

    # CSV 저장
    fields = ["회사", "날짜", "공시제목", "선주", "선종", "척수",
              "인도년", "인도월", "금액(원화,십억원)", "금액(달러,백만)", "확정"]
    with open(args.out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(records)

    print(f"\n완료: {len(records)}건 → {args.out}")


if __name__ == "__main__":
    main()
