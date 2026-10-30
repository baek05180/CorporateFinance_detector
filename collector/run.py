"""수집 실행기.

  python -m collector.run --mode watch            # 오늘 신규 공시만 확인 (평일 5분 간격)
  python -m collector.run --mode daily            # 최근 7일 재수집 + 사모채 (매일 저녁)
  python -m collector.run --mode backfill --days 1095 # 처음 한 번, 과거 3년 채우기
  python -m collector.run --mode profiles         # 기업 프로필 (매일 밤, 호출 예산 안에서 점진적으로)
"""
import argparse
import logging
import os
import sys
from collections import defaultdict
from datetime import date, timedelta

from . import config, normalize
from .dart import DartClient, DartQuotaExceeded
from .datagokr import BondInfoClient, DataGoKrError, redact
from .store import Store, dump_json, load_json, now_kst

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


def compact_bond(it):
    """잔존 회사채 스냅샷 1건 (기업 화면의 만기 구조·차환 분석용)"""
    return {k: v for k, v in {
        "isin": it.get("isinCd"), "crno": it.get("crno"), "issuer": it.get("bondIsurNm"), "name": it.get("isinCdNm"),
        "method": it.get("bondOffrMcdNm"), "link": it.get("pclrBondKcdNm") if it.get("pclrBondKcdNm") not in config.PRIVATE_BOND_PCLR_OK else None,
        "issue": normalize.iso_date(it.get("bondIssuDt")), "mat": normalize.iso_date(it.get("bondExprDt")),
        "amt": normalize.num(it.get("bondIssuAmt")), "bal": normalize.num(it.get("bondBal")),
        "cpn": normalize.pct(it.get("bondSrfcInrt")), "grn": it.get("grnDcdNm"), "opt": it.get("optnTcdNm"),
    }.items() if v not in (None, "", "None")}


def collect_private_bonds(store, mode):
    try:
        client = BondInfoClient()
    except DataGoKrError as e:
        log.warning("사모채 수집 건너뜀: %s", e)
        store.diag.setdefault("private_bond_status", {})["error"] = f"건너뜀: {e}"
        return 0
    today = now_kst().date()
    cutoff = (today - timedelta(days=config.BOND_ISSUE_LOOKBACK_DAYS)).isoformat()
    store.reset_diag("private_bond_kinds")
    store.reset_diag("offer_methods")
    days_done, d, kept = 0, today, 0
    kept_ids = set()
    outstanding = []
    tries = 0
    while days_done < config.BOND_INFO_BASDT_DAYS and tries < 14:
        d -= timedelta(days=1)
        tries += 1
        if d.weekday() >= 5:
            continue
        items = client.bonds_for_basdt(ymd(d))
        store.diag.setdefault("private_bond_status", {})["basDt_" + ymd(d)] = len(items)
        if not items:
            continue
        days_done += 1
        for it in items:
            method = str(it.get("bondOffrMcdNm") or "")
            store.note_diag("offer_methods", method)
            if str(it.get("scrsItmsKcdNm") or "") in config.PRIVATE_BOND_KINDS:
                outstanding.append(compact_bond(it))
            if method not in config.PRIVATE_BOND_METHODS:
                continue
            store.note_diag("private_bond_kinds", f"{it.get('scrsItmsKcdNm')} | {it.get('pclrBondKcdNm')}")
            if str(it.get("scrsItmsKcdNm") or "") not in config.PRIVATE_BOND_KINDS:
                continue
            if str(it.get("pclrBondKcdNm") or "") not in config.PRIVATE_BOND_PCLR_OK:
                continue
            rec = normalize.private_bond(it)
            if not rec["issue_dt"] or rec["issue_dt"] < cutoff or rec["issue_dt"] > today.isoformat():
                continue
            text = " ".join(str(it.get(k) or "") for k in ("sicNm", "isinCdNm"))
            if any(x in text for x in config.PRIVATE_BOND_EXCLUDE):
                continue
            if any(h in (rec["corp_name"] or "") for h in config.FINANCIAL_NAME_HINTS):
                continue
            if rec["maturity"] and rec["issue_dt"]:
                span = date.fromisoformat(rec["maturity"]) - date.fromisoformat(rec["issue_dt"])
                if span.days <= 366:  # 1년 이하 = 단기물 제외
                    continue
            store.upsert(rec, mode)
            kept_ids.add(rec["id"])
            kept += 1
        if len(items) >= 5000:
            # 기준일 하루치가 전체 잔존 채권 스냅샷이므로 하루만 보면 충분.
            # 필터를 바꿨을 때 예전 기준으로 들어온 건이 남지 않도록 이번에 통과 못 한 건은 지운다.
            for rid in [k for k, r in store.records.items() if k.startswith("PBOND:") and k not in kept_ids]:
                del store.records[rid]
                store.changed = True
            outstanding.sort(key=lambda b: (b.get("crno") or "", b.get("isin") or ""))
            dump_json("bonds/outstanding.json", {"basDt": ymd(d), "bonds": outstanding}, indent=None)
            store.diag["private_bond_status"]["outstanding"] = len(outstanding)
            break
    log.info("사모채 %s건 반영 (공공데이터 호출 %s회)", kept, client.calls)
    return client.calls


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["watch", "daily", "backfill", "profiles", "prices", "mezz", "bonds", "probe", "fsfix"], default="watch")
    ap.add_argument("--days", type=int, default=7)
    args = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    store = Store()
    dart = DartClient()
    today = now_kst().date()

    if args.mode == "prices":
        # 주가 API만 빠르게 확인·갱신 (DART 호출 없음)
        import json
        from .refdata import StockPriceError, stock_prices, update_price_history
        try:
            bas, prices = stock_prices(today)
            if prices:
                dump_json("ref/prices.json", {"date": bas, "prices": prices}, indent=None)
            ph, n = update_price_history(load_json("ref/price_hist.json", {}), today, (bas, prices) if prices else None,
                                         max_fetch=140)
            dump_json("ref/price_hist.json", ph, indent=None)
            summary = f"mode=prices · 기준일 {bas} · {len(prices)}종목 · 이력 {len(ph['dates'])}일 (이번 조회 {n}일)"
        except (StockPriceError, Exception) as e:  # noqa: BLE001
            summary = f"mode=prices · 실패: {redact(e)[:300]}"
        log.info(summary)
        if os.environ.get("GITHUB_STEP_SUMMARY"):
            with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as f:
                f.write(summary + "\n")
        return 0

    if args.mode == "fsfix":
        # 신용등급 트리거 대상 회사의 재무제표만 다시 받아 감가상각비(EBITDA) 보강
        import csv
        import traceback
        from .profiles import fs_refresh
        from .store import load_json as _lj
        cmap = (_lj("ref/corpcodes.json", {}) or {}).get("map", {})
        names = {}
        for r in csv.DictReader(open(os.path.join(config.DATA_DIR, "manual", "rating_triggers.csv"), encoding="utf-8-sig")):
            names[(r.get("stock_code") or "").strip()] = 1
        ccs = sorted({cmap[c] for c in names if c in cmap})
        diag = {}
        try:
            fs_refresh(dart, ccs, diag)
        except Exception:  # noqa: BLE001
            diag["error"] = redact(traceback.format_exc())[-1200:]
        dump_json("diag_fsfix.json", diag)
        return 0
    if args.mode == "probe":
        # 공모채 증권신고서 원문 파일 구성·신용평가서 트리거 문구 확인용 (시험)
        import re as _re
        from .docparse import text_of
        recs = sorted([r for r in store.records.values() if r.get("type") == "BOND" and r.get("offering") == "공모"
                       and r.get("rcept_no") and r.get("original_rcept")], key=lambda r: r["rcept_no"], reverse=True)
        picked, seen_c = [], set()
        for r in recs:
            if r["corp_code"] not in seen_c:
                seen_c.add(r["corp_code"]); picked.append(r)
            if len(picked) >= 4:
                break
        out = []
        kw = ["상향 검토", "하향 검토", "상향검토", "하향검토", "트리거", "Trigger", "등급전망", "등급 전망", "Outlook"]
        for r in picked:
            for rno in {r["rcept_no"], r["original_rcept"]}:
                try:
                    docs = dart.document(rno)
                except Exception as e:  # noqa: BLE001
                    out.append({"rcept": rno, "error": redact(e)[:200]}); continue
                files = []
                for n, x in docs.items():
                    t = text_of(x)
                    hits = {}
                    for k in kw:
                        i = t.find(k)
                        if i >= 0:
                            hits[k] = t[max(0, i - 150): i + 350]
                    title = _re.search(r"<TITLE[^>]*>(.*?)</TITLE>", x, _re.S)
                    files.append({"file": n, "len": len(t), "title": text_of(title.group(1))[:60] if title else "",
                                  "has_평가": t.count("신용평가"), "hits": hits})
                out.append({"rcept": rno, "corp": r["corp_name"], "files": files})
        dump_json("diag_probe.json", out)
        return 0
    if args.mode == "bonds":
        import json
        from .profiles import bond_doc_step, link_bond_filings
        import traceback
        diag = {}
        try:
            diag["linked"] = link_bond_filings(dart, store, today)
            store.save(args.mode)
            bond_doc_step(dart, store, today, diag)
        except Exception as e:  # noqa: BLE001
            diag["error"] = redact(traceback.format_exc())[-1500:]
        finally:
            dump_json("diag_bonds.json", diag)
        log.info("mode=bonds · %s", json.dumps({k: diag.get(k) for k in ("linked", "bond_docs")}, ensure_ascii=False))
        return 0
    if args.mode == "mezz":
        import json
        from .profiles import mezz_step
        diag = {}
        try:
            mezz_step(dart, store, today, diag)
        finally:
            dump_json("diag_mezz.json", diag)
        log.info("mode=mezz · %s", json.dumps(diag.get("mezz_docs"), ensure_ascii=False))
        return 0
    if args.mode == "profiles":
        from . import profiles
        profiles.run(dart, store, today)
        d = store.diag.pop("profiles", {})
        # 공시 감시 실행과 같은 파일을 건드리지 않도록 진단은 따로 저장 (동시 push 충돌 방지)
        import json
        dump_json("diag_profiles.json", json.loads(redact(json.dumps(d, ensure_ascii=False))))
        summary = (f"mode=profiles · DART 호출 {dart.calls}회 · 대상 {d.get('targets')}곳 · 갱신 {d.get('profiles_done')}곳"
                   f" · 예산소진={'예' if d.get('budget_exhausted') else '아니오'}")
        log.info(summary)
        if os.environ.get("GITHUB_STEP_SUMMARY"):
            with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as f:
                f.write(summary + "\n")
        return 0

    # GitHub 예약 실행은 붐비면 건너뛰어질 때가 있다. 저녁 정기 갱신을 놓치지 않도록
    # 평일 20:10 이후 감시 실행이 오늘 정기 갱신이 없었음을 보면 정기 갱신으로 바꾼다.
    if args.mode == "watch" and os.environ.get("GITHUB_EVENT_NAME") == "schedule":
        now = now_kst()
        last = store.meta.get("last_daily") or ""
        if now.weekday() < 5 and now.strftime("%H:%M") >= "20:10" and last[:10] != now.date().isoformat():
            log.info("오늘 정기 갱신이 없어 감시 실행을 정기 갱신으로 전환")
            args.mode, args.days = "daily", 7
    if args.mode == "daily" and os.environ.get("GITHUB_EVENT_NAME") == "schedule":
        last = store.meta.get("last_daily") or ""
        if last[:10] == now_kst().date().isoformat() and last[11:16] >= "20:00":
            log.info("오늘 정기 갱신을 이미 마쳐 감시로 전환")
            args.mode = "watch"
    if args.mode == "daily" and args.days < 7:
        args.days = 7
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
        store.diag["private_bond_status"] = {"run": now_kst().isoformat(timespec="minutes")}
        try:
            collect_private_bonds(store, mode)
        except DataGoKrError as e:
            log.error("사모채 수집 실패: %s", e)
            store.diag["private_bond_status"]["error"] = redact(e)[:500]

    if mode == "daily":
        try:
            from .profiles import link_bond_filings
            store.diag["bond_linked"] = link_bond_filings(dart, store, today)
        except Exception as e:  # noqa: BLE001
            store.diag["bond_linked_error"] = redact(e)[:200]
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
