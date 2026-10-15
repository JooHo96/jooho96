"""
DART 재무제표에서 개발비 / 연구개발비 추출

'개발비'는 회계상 두 가지가 섞여 쓰이므로 둘 다 뽑는다.
  1) 개발비 (무형자산)  - 재무상태표 계정. 자산으로 인식(자본화)한 개발비의 '잔액'.
                          기간 비용이 아니라 시점 잔액이므로 증감액으로 봐야 한다.
  2) 연구개발비 / 경상개발비 - 손익계산서(또는 판관비) 계정. 해당 기간에 비용으로
                          털어낸 금액. 보통 '연간 개발비'라고 하면 이쪽.

사용법:
    pip install requests

    # 1) 사명 검색 (정확한 DART 등록명 확인)
    python dev_cost.py --api-key KEY --search 디펜스

    # 2) 개발비 추출 (작년 연간 + 올해 1Q/2Q)
    python dev_cost.py --api-key KEY --company LIG넥스원

    # 연도 지정
    python dev_cost.py --api-key KEY --company LIG넥스원 --year 2026
"""

import argparse
import io
import json
import re
import zipfile
from pathlib import Path

import requests

DART_BASE = "https://opendart.fss.or.kr/api"
CORP_CACHE_FILE = "corp_codes_cache.json"

# 보고서 코드
REPRT = {
    "annual": "11011",   # 사업보고서 (연간)
    "q1":     "11013",   # 1분기보고서
    "h1":     "11012",   # 반기보고서 (= 1~2분기 누적)
    "q3":     "11014",   # 3분기보고서 (= 1~3분기 누적)
}

# 찾을 계정명 패턴
ASSET_PATTERNS = [
    "개발비",
]
EXPENSE_PATTERNS = [
    "연구개발비", "경상개발비", "연구개발", "경상연구개발",
]


# ── 사명 → corp_code ──────────────────────────────────────────────────────

def fetch_corp_index(api_key, cache_path=CORP_CACHE_FILE, refresh=False):
    """corpCode.xml → [(corp_name, corp_code, stock_code), ...] (디스크 캐시)"""
    import xml.etree.ElementTree as ET

    p = Path(cache_path)
    if p.exists() and not refresh:
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass

    r = requests.get(f"{DART_BASE}/corpCode.xml",
                     params={"crtfc_key": api_key}, timeout=120)
    if r.content[:2] != b"PK":
        raise RuntimeError(f"corpCode.xml 응답 이상: {r.text[:200]}")
    with zipfile.ZipFile(io.BytesIO(r.content)) as z:
        raw = z.read(z.namelist()[0])
    root = ET.fromstring(raw)

    entries = [[
        (it.findtext("corp_name")  or "").strip(),
        (it.findtext("corp_code")  or "").strip(),
        (it.findtext("stock_code") or "").strip(),
    ] for it in root.findall("list")]

    try:
        p.write_text(json.dumps(entries, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass
    return entries


def _norm(s):
    return re.sub(r'[\s()（）㈜.&·]', '', s).upper()


def search_corps(entries, keyword, listed_only=False):
    """키워드 포함 기업 목록. 상장사 우선 정렬."""
    k = _norm(keyword)
    hits = [(n, c, s) for n, c, s in entries if k in _norm(n)]
    if listed_only:
        hits = [h for h in hits if h[2]]
    hits.sort(key=lambda e: (0 if e[2] else 1, len(e[0])))
    return hits


def resolve_corp(entries, name):
    """정확 일치 → 접두 일치 → 부분 일치. 상장사·짧은 이름 우선."""
    target = _norm(name)
    exact, prefix, partial = [], [], []
    for n, c, s in entries:
        nn = _norm(n)
        if nn == target:
            exact.append((n, c, s))
        elif nn.startswith(target):
            prefix.append((n, c, s))
        elif target in nn:
            partial.append((n, c, s))
    for bucket in (exact, prefix, partial):
        if bucket:
            bucket.sort(key=lambda e: (0 if e[2] else 1, len(e[0])))
            return bucket[0], bucket
    return None, []


# ── 재무제표 조회 ─────────────────────────────────────────────────────────

def fetch_fs(api_key, corp_code, year, reprt_code, fs_div="CFS"):
    """fnlttSinglAcntAll: 전체 계정 조회.
    fs_div: CFS(연결) / OFS(별도)
    """
    try:
        r = requests.get(f"{DART_BASE}/fnlttSinglAcntAll.json", params={
            "crtfc_key": api_key, "corp_code": corp_code,
            "bsns_year": str(year), "reprt_code": reprt_code,
            "fs_div": fs_div,
        }, timeout=30).json()
    except Exception as e:
        return None, f"요청 실패: {e}"

    status = r.get("status")
    if status != "000":
        return None, f"status={status} ({r.get('message', '')})"
    return r.get("list", []), None


def _to_num(s):
    if s is None:
        return None
    s = str(s).strip().replace(",", "")
    if s in ("", "-"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def find_accounts(rows, patterns, sj_div=None):
    """계정명이 patterns 중 하나와 일치/포함하는 항목 수집.

    sj_div: 'BS'(재무상태표) / 'IS'(손익계산서) / 'CIS'(포괄손익) / 'CF' 등으로 제한
    반환: [(sj_div, account_nm, 당기금액, 당기누적금액)]
    """
    out = []
    for row in rows:
        nm = (row.get("account_nm") or "").strip()
        sj = (row.get("sj_div") or "").strip()
        if sj_div and sj not in sj_div:
            continue
        nm_n = nm.replace(" ", "")
        if not any(p in nm_n for p in patterns):
            continue
        out.append((
            sj, nm,
            _to_num(row.get("thstrm_amount")),
            _to_num(row.get("thstrm_add_amount")),
        ))
    return out


def fmt_won(v):
    """원 단위 → 읽기 쉬운 표기"""
    if v is None:
        return "-"
    sign = "-" if v < 0 else ""
    a = abs(v)
    if a >= 1e12:
        return f"{sign}{a/1e12:,.2f}조원"
    if a >= 1e8:
        return f"{sign}{a/1e8:,.0f}억원"
    return f"{sign}{a:,.0f}원"


# ── 메인 ──────────────────────────────────────────────────────────────────

def report_period(api_key, corp_code, label, year, reprt_code, fs_div):
    rows, err = fetch_fs(api_key, corp_code, year, reprt_code, fs_div)
    print(f"\n── {label}  (fs_div={fs_div}) " + "─" * 30)
    if err:
        print(f"  [조회 실패] {err}")
        return None

    assets   = find_accounts(rows, ASSET_PATTERNS,   sj_div=("BS",))
    expenses = find_accounts(rows, EXPENSE_PATTERNS, sj_div=("IS", "CIS"))

    print("  [무형자산 '개발비' — 기말 잔액]")
    if assets:
        for sj, nm, cur, add in assets:
            print(f"    {nm:28s} {fmt_won(cur):>16s}   ({cur:,.0f} 원)" if cur is not None
                  else f"    {nm:28s} {'-':>16s}")
    else:
        print("    해당 계정 없음 (개발비를 자산으로 인식하지 않거나 계정명이 다름)")

    print("  [손익 '연구개발비 / 경상개발비' — 기간 비용]")
    if expenses:
        for sj, nm, cur, add in expenses:
            cur_s = fmt_won(cur)
            add_s = fmt_won(add)
            print(f"    {nm:28s} 당기 {cur_s:>14s} | 당기누적 {add_s:>14s}")
    else:
        print("    해당 계정 없음 (판관비 주석에만 표시되는 경우가 많음)")

    return {"assets": assets, "expenses": expenses, "rows": len(rows)}


def main():
    ap = argparse.ArgumentParser(description="DART 개발비/연구개발비 추출")
    ap.add_argument("--api-key", required=True)
    ap.add_argument("--company", default="", help="기업명 (DART 등록명)")
    ap.add_argument("--search",  default="", help="사명 키워드로 검색만 하고 종료")
    ap.add_argument("--year",    type=int, default=0,
                    help="올해 연도 (기본: 현재 연도). 작년 연간 + 올해 1Q/2Q를 조회")
    ap.add_argument("--fs-div",  default="CFS", choices=["CFS", "OFS"],
                    help="CFS=연결, OFS=별도 (기본: CFS)")
    ap.add_argument("--refresh-corps", action="store_true", help="사명 캐시 갱신")
    args = ap.parse_args()

    try:
        entries = fetch_corp_index(args.api_key, refresh=args.refresh_corps)
    except Exception as e:
        print(f"[오류] 기업목록 조회 실패: {e}")
        return
    print(f"DART 등록 기업 {len(entries):,}건 로드")

    # 검색 모드
    if args.search:
        hits = search_corps(entries, args.search)
        print(f"\n'{args.search}' 포함 기업 {len(hits)}건:")
        for n, c, s in hits[:40]:
            tag = f"stock={s}" if s else "비상장"
            print(f"  {n:34s} corp_code={c}  {tag}")
        if len(hits) > 40:
            print(f"  ... 외 {len(hits)-40}건")
        return

    if not args.company:
        print("\n[안내] --company 로 기업명을 지정하거나 --search 로 먼저 검색하세요.")
        return

    hit, bucket = resolve_corp(entries, args.company)
    if not hit:
        print(f"\n[오류] '{args.company}' 를 찾지 못했습니다. --search 로 검색해 보세요.")
        return

    corp_name, corp_code, stock_code = hit
    print(f"\n대상: {corp_name}  (corp_code={corp_code}, "
          f"{'stock=' + stock_code if stock_code else '비상장'})")
    if len(bucket) > 1:
        print(f"  ※ 유사 후보 {len(bucket)-1}건 더 있음: "
              f"{', '.join(n for n, _, _ in bucket[1:6])}")
        print(f"     의도한 회사가 아니면 --company 에 정확한 등록명을 넣으세요.")

    from datetime import date
    this_year = args.year or date.today().year
    last_year = this_year - 1

    res_annual = report_period(args.api_key, corp_code,
                               f"{last_year}년 연간 (사업보고서)",
                               last_year, REPRT["annual"], args.fs_div)
    res_q1 = report_period(args.api_key, corp_code,
                           f"{this_year}년 1분기",
                           this_year, REPRT["q1"], args.fs_div)
    res_h1 = report_period(args.api_key, corp_code,
                           f"{this_year}년 반기 (1~2분기 누적)",
                           this_year, REPRT["h1"], args.fs_div)

    # 2분기 단독 = 반기누적 − 1분기
    print("\n" + "=" * 62)
    print(f"{this_year}년 2분기 단독 추정  (반기누적 − 1분기)")
    print("=" * 62)
    if res_q1 and res_h1:
        def pick(res):
            """비용 계정 중 누적값 우선으로 {계정명: 금액}"""
            d = {}
            for sj, nm, cur, add in res["expenses"]:
                d[nm.replace(" ", "")] = add if add is not None else cur
            return d

        q1d, h1d = pick(res_q1), pick(res_h1)
        common = [k for k in h1d if k in q1d
                  and h1d[k] is not None and q1d[k] is not None]
        if common:
            for k in common:
                q2 = h1d[k] - q1d[k]
                print(f"  {k:28s} {fmt_won(q2):>16s}"
                      f"   (반기 {fmt_won(h1d[k])} − 1Q {fmt_won(q1d[k])})")
        else:
            print("  비교 가능한 공통 비용계정이 없어 2분기 단독 계산을 못 했습니다.")
            print("  위 1분기/반기 출력값을 직접 비교해 주세요.")
    else:
        print("  1분기 또는 반기 조회가 실패해 계산할 수 없습니다.")

    print("""
[주의]
 · '개발비(무형자산)'는 잔액이므로 기간 비용이 아닙니다. 기간 투입액을 보려면
   연도별 잔액 증감 + 상각비를 함께 봐야 합니다.
 · 많은 기업이 연구개발비를 손익계산서 본문에 별도 표시하지 않고 판관비 주석이나
   사업보고서 '연구개발활동' 항목에만 기재합니다. 위에서 '해당 계정 없음'이 나오면
   DART 사업보고서 본문을 직접 확인해야 합니다.
 · 2분기 단독값은 누적치 차감으로 구한 추정이며, 분기 재작성이 있으면 달라집니다.
""")


if __name__ == "__main__":
    main()
