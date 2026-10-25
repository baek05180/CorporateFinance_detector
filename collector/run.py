"""수집 실행기.

  python -m collector.run --mode watch            # 오늘 신규 공시만 확인 (평일 5분 간격)
  python -m collector.run --mode daily            # 최근 7일 재수집 + 사모채 (매일 저녁)
  python -m collector.run --mode backfill --days 90   # 처음 한 번, 과거 데이터 채우기
"""
import argparse
import logging
import os
import sys
from collections import defaultdict
from datetime import date, timedelta

from . import config, normalize
from .dart import DartClient, DartQuotaExceeded
from .datagokr import BondInfoClient, DataGoKrError
from .store import Store, now_kst

log = logging.getLogger("collector")


def ymd(d):
    return d.strftime("%Y%m%d")


def classify(filing):
    """공시 목록 한 줄 → 'CB'|'BW'|'EB'|'RIGHTS'|'BOND'|None"""
    nm = (filing.get("report_nm") or "").replace(" ", "")
    if any(x in nm for x in config.MAJOR_REPORT_EXCLUDE):
        return None
    for key, kind in config.MAJOR_REPORT_KINDS.items():
        if key in nm:
            return kind
    if any(k in nm for k in config.BOND_REPORT_KEYWORDS):
        return "BOND"
    return None


def list_relevant(dart, start, end):
    """기간 내 주요사항보고서 + 채무증권 증권신고서 목록. (corp_code 없이 조회는 최대 3개월)"""
    out = []
    cur = start
    while cur <= end:
        chunk_end = min(cur + timedelta(days=60), end)
        out += dart.list_filings(ymd(cur), ymd(chunk_end), "B", "B001")
        out += dart.list_filings(ymd(cur), ymd(chunk_end), "C", "C002")
        cur = chunk_end + timedelta(days=1)
    return out


def is_financial(dart, store, corp_code, corp_name):
    info = store.corp_cache.get(corp_code)
    if info is None:
        try:
            c = dart.company(corp_code)
            info = {"induty_code": c.get("induty_code") or "", "stock_code": c.get("stock_code") or ""}
        except DartQuotaExceeded:
            raise
        except Exception as e:  # 기업개황 실패는 치명적이지 않음 → 이름으로 판단
            log.warning("기업개황 실패 %s: %s", corp_code, e)
            info = {"induty_code": "", "stock_code": ""}
        store.corp_cache[corp_code] = info
    code = info.get("induty_code", "")
    if code:
        return code.startswith(config.FINANCIAL_KSIC_PREFIX)
    return any(h in (corp_name or "") for h in config.FINANCIAL_NAME_HINTS)


def process_filings(dart, store, filings, mode):
    tasks = defaultdict(list)
    for f in filings:
        kind = classify(f)
        if not kind:
            continue
        if mode == "watch" and f["rcept_no"] in store.seen:
            continue
        tasks[(f["corp_code"], kind)].append(f)

    for (corp_code, kind), fs in sorted(tasks.items()):
        by_rno = {f["rcept_no"]: f for f in fs}
        last = max(f["rcept_dt"] for f in fs)
        first = min(f["rcept_dt"] for f in fs)
        end = date(int(last[:4]), int(last[4:6]), int(last[6:8]))
        begin = date(int(first[:4]), int(first[4:6]), int(first[6:8])) - timedelta(days=config.DETAIL_LOOKBACK_DAYS)

        if kind == "BOND":
            if is_financial(dart, store, corp_code, fs[0].get("corp_name")):
                for f in fs:
                    store.mark_seen(f["rcept_no"])
                continue
            groups = dart.bond_registration(corp_code, ymd(begin), ymd(end))
            general = groups.get("일반사항", [])
            for rno in sorted(by_rno):
                sub = {t: [x for x in lst if x.get("rcept_no") == rno] for t, lst in groups.items()}
                if not sub.get("일반사항"):
                    continue
                for rec in normalize.public_bonds(sub, by_rno[rno]):
                    name = rec.get("name") or ""
                    if any(x in name for x in config.BOND_NAME_EXCLUDE):
                        continue
                    store.upsert(rec, mode)
                store.mark_seen(rno)
            if not general:
                log.info("채무증권 상세 아직 없음: %s %s", fs[0].get("corp_name"), list(by_rno))
            continue

        items = dart.major_detail(config.DETAIL_ENDPOINT[kind], corp_code, ymd(begin), ymd(end))
        for item in sorted(items, key=lambda x: x.get("rcept_no", "")):
            rno = item.get("rcept_no")
            if rno not in by_rno:
                continue
            filing = by_rno[rno]
            if kind == "RIGHTS":
                rec = store.merge_rights_amendment(normalize.rights(item, filing))
            else:
                rec = normalize.mezzanine(item, kind, filing)
            store.upsert(rec, mode)
            store.mark_seen(rno)


def collect_private_bonds(store, mode):
    try:
        client = BondInfoClient()
    except DataGoKrError as e:
        log.warning("사모채 수집 건너뜀: %s", e)
        return 0
    today = now_kst().date()
    cutoff = (today - timedelta(days=config.BOND_ISSUE_LOOKBACK_DAYS)).isoformat()
    store.reset_diag("private_bond_kinds")
    store.reset_diag("offer_methods")
    days_done, d, kept = 0, today, 0
    tries = 0
    while days_done < config.BOND_INFO_BASDT_DAYS and tries < 14:
        d -= timedelta(days=1)
        tries += 1
        if d.weekday() >= 5:
            continue
        items = client.bonds_for_basdt(ymd(d))
        if not items:
            continue
        days_done += 1
        for it in items:
            store.note_diag("offer_methods", str(it.get("bondOffrMcdNm")))
            if "사모" not in str(it.get("bondOffrMcdNm") or ""):
                continue
            store.note_diag("private_bond_kinds", f"{it.get('scrsItmsKcdNm')} | {it.get('pclrBondKcdNm')}")
            rec = normalize.private_bond(it)
            if not rec["issue_dt"] or rec["issue_dt"] < cutoff:
                continue
            text = " ".join(str(it.get(k) or "") for k in
                            ("scrsItmsKcdNm", "sicNm", "pclrBondKcdNm", "isinCdNm"))
            if any(x in text for x in config.PRIVATE_BOND_EXCLUDE):
                continue
            if any(h in (rec["corp_name"] or "") for h in config.FINANCIAL_NAME_HINTS):
                continue
            if rec["maturity"] and rec["issue_dt"]:
                span = date.fromisoformat(rec["maturity"]) - date.fromisoformat(rec["issue_dt"])
                if span.days <= 366:  # 1년 이하 = 단기물 제외
                    continue
            store.upsert(rec, mode)
            kept += 1
        if len(items) >= 5000:
            break  # 기준일 하루치가 전체 스냅샷이면 하루만 보면 충분
    log.info("사모채 %s건 반영 (공공데이터 호출 %s회)", kept, client.calls)
    return client.calls


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["watch", "daily", "backfill"], default="watch")
    ap.add_argument("--days", type=int, default=7)
    args = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    store = Store()
    dart = DartClient()
    today = now_kst().date()
    mode = "daily" if args.mode == "backfill" else args.mode

    if args.mode == "watch":
        start = today
    elif args.mode == "daily":
        start = today - timedelta(days=args.days)
    else:
        start = today - timedelta(days=args.days)

    try:
        filings = list_relevant(dart, start, today)
        log.info("공시 목록 %s건 확인", len(filings))
        process_filings(dart, store, filings, mode)
    except DartQuotaExceeded as e:
        log.error("OpenDART 일일 한도 초과: %s (지금까지 받은 데이터는 저장)", e)

    if mode == "daily":
        try:
            collect_private_bonds(store, mode)
        except DataGoKrError as e:
            log.error("사모채 수집 실패: %s", e)

    wrote = store.save(mode)
    summary = f"mode={args.mode} 신규 {len(store.added)}건 · DART 호출 {dart.calls}회 · 파일 변경={'예' if wrote else '아니오'}"
    log.info(summary)
    if os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as f:
            f.write(summary + "\n")
            for rid in store.added[:50]:
                r = store.records[rid]
                f.write(f"- {r['type']} {r.get('corp_name')} {r.get('amount')}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
