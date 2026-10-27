"""기업 프로필 수집 (python -m collector.run --mode profiles).

호출 예산(config.DART_RUN_BUDGET) 안에서 우선순위 순으로 채우고, 다 못 하면 다음 실행에서 이어간다.
저장 위치
  data/ref/listed.json      KIND 상장사 목록 (업종·주요제품)
  data/ref/corpcodes.json   종목코드 → DART 고유번호
  data/ref/prices.json      종가·시가총액 (공공데이터포털, 활용신청 필요)
  data/ref/company_info.json 기업개황 (법인등록번호·업종코드·대표 등)
  data/fin/multi.json       상장사 전체 주요계정 (peer 비교용)
  data/co/{고유번호}.json    조달 대상 회사 상세 (사업내용·재무 세부·주주·자기주식)
  data/ref/bond_docs.json   공모채 증권신고서에서 뽑은 수요예측일
"""
import logging
import re
import time
from datetime import date, timedelta

from . import config, docparse
from .dart import DartBudgetExceeded, DartError, DartQuotaExceeded
from .normalize import num
from .datagokr import redact
from .refdata import StockPriceError, kind_listed, stock_prices, update_price_history
from .store import dump_json, load_json, now_kst

log = logging.getLogger("profiles")


def ymd(d):
    return d.strftime("%Y%m%d")


def _stale(stamp, days):
    if not stamp:
        return True
    try:
        return (now_kst().date() - date.fromisoformat(stamp[:10])).days >= days
    except ValueError:
        return True


# ── 보고 기간 판단 ──
def latest_periods(today):
    """(연간 사업연도, [(연도, 보고서코드), ...] 최근 분기→과거 순)"""
    y = today.year
    annual = y - 1 if today >= date(y, 4, 5) else y - 2
    interim = []
    if today >= date(y, 11, 20):
        interim.append((y, "11014"))
    if today >= date(y, 8, 20):
        interim.append((y, "11012"))
    if today >= date(y, 5, 20):
        interim.append((y, "11013"))
    return annual, interim


# ── 주요계정 (다중회사) ──
def parse_multi(rows):
    """fnlttMultiAcnt 행들 → {종목코드: {"fs": CFS|OFS, "dt": 기준일, "cur": {...}, "prev": {...}, "prev2": {...}}}"""
    by = {}
    for r in rows:
        code = (r.get("stock_code") or "").strip()
        key = config.MULTI_ACCOUNTS.get((r.get("account_nm") or "").strip())
        if not code or not key:
            continue
        fs = r.get("fs_div")
        slot = by.setdefault(code, {}).setdefault(fs, {"cur": {}, "prev": {}, "prev2": {}, "dt": r.get("thstrm_dt")})
        is_flow = r.get("sj_div") == "IS"
        cur = num(r.get("thstrm_add_amount")) if is_flow and r.get("thstrm_add_amount") not in (None, "", "-") else num(r.get("thstrm_amount"))
        prev = num(r.get("frmtrm_add_amount")) if is_flow and r.get("frmtrm_add_amount") not in (None, "", "-") else num(r.get("frmtrm_amount"))
        for k, v in (("cur", cur), ("prev", prev), ("prev2", num(r.get("bfefrmtrm_amount")))):
            if v is not None and key not in slot[k]:
                slot[k][key] = v
    out = {}
    for code, fss in by.items():
        fs = "CFS" if "CFS" in fss and fss["CFS"]["cur"] else "OFS"
        if fs in fss:
            out[code] = {"fs": fs, **fss[fs]}
    return out


def refresh_multi(dart, corpcodes, fin, today):
    annual, interim = latest_periods(today)
    codes = sorted({v for v in corpcodes.values() if v})
    stamp_key = f"{annual}|{interim[0] if interim else ''}"
    if fin.get("stamp_key") == stamp_key and not _stale(fin.get("updated"), 7):
        return fin
    fin = {"stamp_key": stamp_key, "updated": now_kst().isoformat(timespec="minutes"), "annual_year": annual,
           "annual": {}, "interim": {}, "interim_period": None}
    for i in range(0, len(codes), 100):
        chunk = codes[i:i + 100]
        fin["annual"].update(parse_multi(dart.multi_accounts(chunk, annual, "11011")))
    for (y, rc) in interim:
        got = {}
        for i in range(0, len(codes), 100):
            got.update(parse_multi(dart.multi_accounts(codes[i:i + 100], y, rc)))
        if len(got) > 200:   # 해당 분기 보고서가 충분히 나왔을 때만 채택
            fin["interim"] = got
            fin["interim_period"] = f"{y} {config.REPORT_CODES[rc]}"
            break
    log.info("주요계정: 연간 %s곳, 분기 %s곳 (%s)", len(fin["annual"]), len(fin["interim"]), fin["interim_period"])
    return fin


# ── 전체 재무제표에서 세부 항목 ──
DEBT_KEYS = ("차입금", "사채", "유동성장기부채", "상환우선주부채", "리스부채")
DEBT_SKIP = ("할인", "할증", "조정", "이자", "미지급", "파생", "전환권", "상환권", "증가", "감소", "상환", "발행")


def fs_detail(rows):
    out = {}
    debt_names = set()
    debt = 0
    debt_prev = [0, 0]      # 전기, 전전기 (연간 보고서의 비교 기간)
    has_prev = [False, False]
    for r in rows:
        sj, aid, nm = r.get("sj_div"), r.get("account_id") or "", (r.get("account_nm") or "").replace(" ", "")
        v = num(r.get("thstrm_add_amount")) if sj in ("IS", "CIS") and r.get("thstrm_add_amount") not in (None, "", "-") \
            else num(r.get("thstrm_amount"))
        if v is None:
            continue
        if sj == "BS":
            if aid == "ifrs-full_CashAndCashEquivalents" and "cash" not in out:
                out["cash"] = v
            elif ("ShorttermDeposits" in aid or nm == "단기금융상품") and "st_deposits" not in out:
                out["st_deposits"] = v
            elif aid == "ifrs-full_Assets":
                out.setdefault("assets", v)
            elif aid == "ifrs-full_Liabilities":
                out.setdefault("liab", v)
            elif aid == "ifrs-full_Equity":
                out.setdefault("equity", v)
            elif aid == "ifrs-full_IssuedCapital":
                out.setdefault("capital", v)
            elif any(k in nm for k in DEBT_KEYS) and not any(k in nm for k in DEBT_SKIP) and nm not in debt_names:
                debt_names.add(nm)
                debt += v
                for i, f in enumerate(("frmtrm_amount", "bfefrmtrm_amount")):
                    pv = num(r.get(f))
                    if pv is not None:
                        debt_prev[i] += pv
                        has_prev[i] = True
        elif sj in ("IS", "CIS"):
            if aid in ("ifrs-full_Revenue",) or nm in ("매출액", "수익(매출액)", "영업수익"):
                out.setdefault("revenue", v)
            elif aid == "dart_OperatingIncomeLoss" or nm.startswith("영업이익"):
                out.setdefault("op", v)
            elif aid == "ifrs-full_FinanceCosts" or nm in ("금융원가", "금융비용"):
                out.setdefault("fin_cost", v)
            elif nm in ("이자비용",):
                out.setdefault("interest", v)
            elif aid == "ifrs-full_ProfitLoss" or nm.startswith("당기순이익"):
                out.setdefault("ni", v)
        elif sj == "CF":
            if aid == "ifrs-full_CashFlowsFromUsedInOperatingActivities":
                out.setdefault("ocf", v)
            elif aid == "ifrs-full_AdjustmentsForDepreciationExpense" or nm.startswith("감가상각비"):
                out.setdefault("dep", v)
            elif aid == "ifrs-full_AdjustmentsForAmortisationExpense" or nm.startswith("무형자산상각비"):
                out.setdefault("amort", v)
            elif aid == "ifrs-full_PurchaseOfPropertyPlantAndEquipment" or nm.startswith("유형자산의취득"):
                out.setdefault("capex", abs(v))
            elif aid.startswith("ifrs-full_InterestPaid") or nm in ("이자의지급", "이자지급", "이자의지급(영업)", "이자지급액"):
                out["interest_paid"] = out.get("interest_paid", 0) + abs(v)
    if debt_names:
        out["debt"] = debt
        out["debt_items"] = sorted(debt_names)
        out["debt_hist"] = [debt_prev[1] if has_prev[1] else None, debt_prev[0] if has_prev[0] else None, debt]
    return out


PROFILE_VERSION = 3   # 추출 규칙이 바뀌면 올린다 → 기존 프로필을 다음 실행 때 다시 만든다


def _pct(v):
    try:
        return float(str(v or "0").replace(",", "").replace("%", "").strip() or 0)
    except ValueError:      # '-' 등
        return 0.0


def holders(rows):
    common = [r for r in rows if any(k in (r.get("stock_knd") or "") for k in ("보통", "의결권")) or not r.get("stock_knd")]
    if not common and rows:   # 주식 종류 표기가 회사마다 달라 못 찾으면 가장 많이 나온 종류로
        from collections import Counter
        kind = Counter(r.get("stock_knd") for r in rows).most_common(1)[0][0]
        common = [r for r in rows if r.get("stock_knd") == kind]
    total = next((r for r in common if (r.get("nm") or "").replace(" ", "") in ("계", "합계", "총계")), None)
    people = [r for r in common if r is not total and (r.get("nm") or "").replace(" ", "") not in ("계", "합계", "총계")]
    top = sorted(people, key=lambda r: -(num(r.get("trmend_posesn_stock_co")) or 0))[:6]
    out = {"top": [{"name": r.get("nm"), "relation": r.get("relate"),
                    "pct": _pct(r.get("trmend_posesn_stock_qota_rt")),
                    "shares": num(r.get("trmend_posesn_stock_co"))} for r in top]}
    if total and _pct(total.get("trmend_posesn_stock_qota_rt")):
        out["total_pct"] = _pct(total.get("trmend_posesn_stock_qota_rt"))
    if "total_pct" not in out and people:
        out["total_pct"] = round(sum(x["pct"] for x in out["top"]), 2)
    if rows:
        out["as_of"] = rows[0].get("stlm_dt")
    return out


def treasury_common(rows):
    for r in rows:
        if "보통" in (r.get("stock_knd") or "") and (r.get("acqs_mth1") or "").replace(" ", "") == "총계":
            return num(r.get("trmend_qy")) or 0
    return None


REPORT_RE = re.compile(r"(사업보고서|반기보고서|분기보고서)\s*\((\d{4})\.(\d{2})\)")


def classify_report(nm):
    m = REPORT_RE.search(nm or "")
    if not m:
        return None
    kind, y, mo = m.group(1), int(m.group(2)), int(m.group(3))
    if kind == "사업보고서":
        return (y, "11011")
    if kind == "반기보고서":
        return (y, "11012")
    return (y, "11013" if mo <= 4 else "11014")


def build_profile(dart, corp_code, old, today):
    """회사 1곳 프로필 갱신. 바뀐 정기보고서가 없으면 목록 조회 1회로 끝난다."""
    reports = dart.periodic_reports(corp_code, ymd(today - timedelta(days=550)), ymd(today))
    reports = [r for r in reports if classify_report(r.get("report_nm"))]
    reports.sort(key=lambda r: r["rcept_no"], reverse=True)
    prof = dict(old or {})
    if prof.get("pv") != PROFILE_VERSION:      # 규칙이 바뀌었으면 원문·재무를 다시 읽는다
        prof.pop("annual_rcept", None)
        prof.pop("latest_rcept", None)
    prof["pv"] = PROFILE_VERSION
    prof["corp_code"] = corp_code
    prof["checked"] = now_kst().isoformat(timespec="minutes")
    if not reports:
        prof["no_periodic"] = True
        return prof
    prof.pop("no_periodic", None)
    latest = reports[0]
    annual = next((r for r in reports if "사업보고서" in r["report_nm"]), None)
    ly, lrc = classify_report(latest["report_nm"])

    if annual and prof.get("annual_rcept") != annual["rcept_no"]:
        ay, _ = classify_report(annual["report_nm"])
        try:
            docs = dart.document(annual["rcept_no"])
            main = max(docs.values(), key=len) if docs else ""
            prof["business"] = docparse.annual_business(main)
            prof["biz_raw"] = docparse.business_raw(main)   # AI 요약 재료 (사이트에는 내보내지 않음)
        except (DartError, ValueError) as e:
            log.warning("사업보고서 원문 실패 %s: %s", corp_code, e)
            prof["business"] = prof.get("business") or {}
        rows, fs = dart.full_statements(corp_code, ay, "11011")
        prof["fs_annual"] = {"period": f"{ay} 연간", "fs": fs, **fs_detail(rows)} if rows else {}
        prof["annual_rcept"] = annual["rcept_no"]
        prof["annual_name"] = annual["report_nm"]

    if prof.get("latest_rcept") != latest["rcept_no"]:
        if lrc != "11011":
            rows, fs = dart.full_statements(corp_code, ly, lrc)
            prof["fs_interim"] = {"period": f"{ly} {config.REPORT_CODES[lrc]}", "fs": fs, **fs_detail(rows)} if rows else {}
        else:
            prof.pop("fs_interim", None)
        prof["holders"] = holders(dart.shareholders(corp_code, ly, lrc))
        tq = treasury_common(dart.treasury(corp_code, ly, lrc))
        if tq is None and lrc != "11011" and annual:
            ay, _ = classify_report(annual["report_nm"])
            tq = treasury_common(dart.treasury(corp_code, ay, "11011"))
        prof["treasury_shares"] = tq
        prof["latest_rcept"] = latest["rcept_no"]
        prof["latest_name"] = latest["report_nm"]
    return prof


# ── 대상 회사 고르기 ──
def read_watchlist():
    import os
    p = os.path.join(config.ROOT, config.WATCHLIST_FILE)
    try:
        with open(p, encoding="utf-8") as f:
            return [ln.split("#")[0].strip() for ln in f if ln.split("#")[0].strip()]
    except FileNotFoundError:
        return []


def resolve_watchlist(items, listed, corpcodes):
    name_to_code = {}
    for code, v in listed.items():
        name_to_code.setdefault(v["name"].replace(" ", ""), code)
    out = []
    for it in items:
        key = it.replace(" ", "")
        if re.fullmatch(r"\d{8}", key):
            out.append(key)
        elif key.upper().zfill(6) in corpcodes:
            out.append(corpcodes[key.upper().zfill(6)])
        elif key in name_to_code and name_to_code[key] in corpcodes:
            out.append(corpcodes[name_to_code[key]])
        else:
            log.warning("관심 회사 '%s'를 찾지 못함 (상장사 이름·종목코드·고유번호 8자리로 적어주세요)", it)
    return out


def pick_targets(records, info, watch, today):
    """[(우선순위키, corp_code)] — 최근 조달이 있는 회사부터"""
    cutoff = (today - timedelta(days=config.PROFILE_TARGET_DAYS)).isoformat()
    jur = {v.get("jurir_no"): k for k, v in info.items() if v.get("jurir_no")}
    last = {}
    for r in records.values():
        cc = r.get("corp_code") or jur.get(r.get("crno"))
        d = r.get("board_dt") or r.get("sub_dt") or r.get("issue_dt") or r.get("rcept_dt") or ""
        if cc and d >= cutoff and d > last.get(cc, ""):
            last[cc] = d
    for cc in watch:
        last[cc] = "9999"   # 관심 회사는 맨 앞
    return [cc for cc, _ in sorted(last.items(), key=lambda x: x[1], reverse=True)]


def parallel(items, fn, workers=None):
    """items를 여러 스레드로 처리하며 결과를 내보낸다. 예산·한도 초과가 나면 남은 일을 멈추고 그 예외를 다시 던진다."""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    stop = []

    def wrap(x):
        if stop:
            return None
        try:
            return fn(x)
        except (DartBudgetExceeded, DartQuotaExceeded) as e:
            stop.append(e)
            return None
        except Exception:  # noqa: BLE001 — 한 회사의 이상한 데이터 때문에 전체가 멈추지 않게
            log.exception("개별 처리 실패: %s", str(x)[:40])
            return None
    with ThreadPoolExecutor(max_workers=workers or config.PROFILE_WORKERS) as ex:
        futs = [ex.submit(wrap, x) for x in items]
        for f in as_completed(futs):
            res = f.result()
            if res is not None:
                yield res
    if stop:
        raise stop[0]


# ── 실행 ──
def run(dart, store, today=None):
    today = today or now_kst().date()
    dart.budget = dart.calls + config.DART_RUN_BUDGET
    dart.deadline = time.time() + config.PROFILE_TIME_BUDGET_MIN * 60
    diag = store.diag.setdefault("profiles", {})
    diag.clear()
    diag["run"] = now_kst().isoformat(timespec="minutes")

    # 1) 상장사 목록
    listed = load_json("ref/listed.json", {})
    try:
        new = kind_listed()
        if len(new) >= config.KIND_MIN_ROWS:
            listed = new
            dump_json("ref/listed.json", listed)
        diag["listed"] = len(listed)
    except Exception as e:  # noqa: BLE001
        diag["listed_error"] = redact(e)[:300]
        log.warning("KIND 목록 실패: %s", e)

    # 2) 종목코드 → 고유번호
    cc = load_json("ref/corpcodes.json", {})
    if not cc.get("map") or _stale(cc.get("updated"), 7):
        try:
            rows = dart.corp_codes()
            cc = {"updated": now_kst().isoformat(timespec="minutes"),
                  "map": {r["stock_code"]: r["corp_code"] for r in rows if r["stock_code"]}}
            dump_json("ref/corpcodes.json", cc)
        except (DartError, DartBudgetExceeded) as e:
            diag["corpcode_error"] = redact(e)[:200]
    corpcodes = cc.get("map", {})
    listed_cc = {code: corpcodes[code] for code in listed if code in corpcodes}

    # 3) 주가·시가총액 (활용신청 안 돼 있으면 건너뜀)
    try:
        bas, prices = stock_prices(today)
        if prices:
            dump_json("ref/prices.json", {"date": bas, "prices": prices}, indent=None)
        diag["prices"] = len(prices)
        ph, n = update_price_history(load_json("ref/price_hist.json", {}), today, (bas, prices) if prices else None)
        dump_json("ref/price_hist.json", ph, indent=None)
        diag["price_hist_days"], diag["price_hist_fetched"] = len(ph["dates"]), n
    except StockPriceError as e:
        diag["prices_error"] = redact(e)[:300]
    except Exception as e:  # noqa: BLE001
        diag["prices_error"] = redact(e)[:300]

    info = load_json("ref/company_info.json", {})
    try:
        # 4) 상장사 전체 주요계정 (100곳씩 묶어 조회 → 수십 회)
        fin = load_json("fin/multi.json", {})
        fin = refresh_multi(dart, listed_cc, fin, today)
        dump_json("fin/multi.json", fin, indent=None)

        # 5) 공모채 수요예측일·신용등급 (최근 것만, 병렬)
        t0 = time.time()
        bond_docs = load_json("ref/bond_docs.json", {})
        since_doc = (today - timedelta(days=config.BOND_DOC_DAYS)).isoformat()
        todo = sorted({r.get("original_rcept") or r["rcept_no"] for r in store.records.values()
                       if r.get("type") == "BOND" and r.get("offering") == "공모" and r.get("rcept_no")
                       and (r.get("rcept_dt") or "") >= since_doc}, reverse=True)
        todo = [rno for rno in todo if (bond_docs.get(rno) or {}).get("v") != docparse.BOND_DOC_VERSION]

        def bond_job(rno):
            try:
                docs = dart.document(rno)
            except DartError as e:
                return rno, {"demand_dt": None, "error": redact(e)[:100], "v": docparse.BOND_DOC_VERSION}, None
            main = max(docs.values(), key=len) if docs else ""
            got = docparse.bond_schedule(main)
            miss = None
            if not got.get("demand_dt") or not got.get("rating"):
                t = docparse.text_of(main)
                i, j = t.find("수요예측"), t.find("신용등급")
                miss = {"rcept": rno, "got": got, "demand": t[max(0, i - 80): i + 220] if i >= 0 else "",
                        "rating": t[max(0, j - 40): j + 300] if j >= 0 else ""}
            return rno, got, miss

        try:
            for rno, got, miss in parallel(todo, bond_job):
                bond_docs[rno] = got
                if miss and len(diag.setdefault("demand_miss", [])) < 6:
                    diag["demand_miss"].append(miss)
        finally:
            dump_json("ref/bond_docs.json", bond_docs)
            diag["bond_docs"] = {"todo": len(todo), "found_demand": sum(1 for v in bond_docs.values() if v.get("demand_dt")),
                                 "found_rating": sum(1 for v in bond_docs.values() if v.get("rating")),
                                 "sec": round(time.time() - t0)}

        # 5-2) 메자닌(CB·BW·EB) 원문: 투자자·콜/풋·할증률 (최근 것부터, 회당 최대 MEZZ_DOC_MAX건)
        t0 = time.time()
        mezz_docs = load_json("ref/mezz_docs.json", {})
        since_m = (today - timedelta(days=config.MEZZ_DOC_DAYS)).isoformat()
        mtodo = sorted({r["rcept_no"] for r in store.records.values()
                        if r.get("type") in ("CB", "BW", "EB") and r.get("rcept_no") and (r.get("rcept_dt") or "") >= since_m},
                       reverse=True)
        mtodo = [x for x in mtodo if (mezz_docs.get(x) or {}).get("v") != docparse.MEZZ_DOC_VERSION][:config.MEZZ_DOC_MAX]

        def mezz_job(rno):
            try:
                docs = dart.document(rno)
            except DartError as e:
                return rno, {"error": redact(e)[:100], "v": docparse.MEZZ_DOC_VERSION}, None
            main = max(docs.values(), key=len) if docs else ""
            got = docparse.mezz_doc(main)
            miss = None
            if not got["investors"] or got["call_pct"] is None or got["premium"] is None:
                t = docparse.text_of(main)
                miss = {"rcept": rno, "got": {k: got[k] for k in ("call_pct", "call_m", "put_m", "premium")},
                        "n_inv": len(got["investors"])}
                for key, w in (("대상자", 500), ("매도청구권", 700), ("할증", 200), ("기준주가", 200)):
                    i = t.find(key)
                    miss[key] = t[max(0, i - 60): i + w] if i >= 0 else ""
            return rno, got, miss

        try:
            for rno, got, miss in parallel(mtodo, mezz_job):
                mezz_docs[rno] = got
                if miss and len(diag.setdefault("mezz_miss", [])) < 8:
                    diag["mezz_miss"].append(miss)
        finally:
            dump_json("ref/mezz_docs.json", mezz_docs, indent=None)
            vals = [v for v in mezz_docs.values() if "error" not in v]
            diag["mezz_docs"] = {"todo": len(mtodo), "total": len(mezz_docs), "with_inv": sum(1 for v in vals if v["investors"]),
                                 "with_call": sum(1 for v in vals if v["call_pct"] is not None),
                                 "with_prem": sum(1 for v in vals if v["premium"] is not None), "sec": round(time.time() - t0)}

        # 6) 조달 대상 회사 상세 (병렬)
        t0 = time.time()
        watch = resolve_watchlist(read_watchlist(), listed, corpcodes)
        targets = pick_targets(store.records, info, watch, today)
        diag["targets"] = len(targets)
        todo = []
        for corp in targets:
            old = load_json(f"co/{corp}.json", None)
            if old and old.get("pv") == PROFILE_VERSION and not _stale(old.get("checked"), config.PROFILE_RECHECK_DAYS):
                continue
            todo.append((corp, old))
        diag["profiles_fresh"] = len(targets) - len(todo)

        def profile_job(item):
            corp, old = item
            if not info.get(corp) or _stale(info[corp].get("checked"), config.COMPANY_INFO_RECHECK_DAYS):
                c = dart.company(corp)
                info[corp] = {k: c.get(k) for k in ("corp_name", "stock_name", "stock_code", "ceo_nm", "corp_cls",
                                                     "jurir_no", "adres", "hm_url", "induty_code", "est_dt", "acc_mt")}
                info[corp]["checked"] = now_kst().isoformat(timespec="minutes")
            try:
                prof = build_profile(dart, corp, old, today)
            except DartError as e:
                log.warning("프로필 실패 %s: %s", corp, e)
                return False
            dump_json(f"co/{corp}.json", prof)
            return True

        done = 0
        try:
            for ok in parallel(todo, profile_job):
                done += bool(ok)
        finally:
            diag["profiles_done"] = done
            diag["profiles_left"] = len(todo) - done
            diag["profiles_sec"] = round(time.time() - t0)

        # 7) 상장사 기업개황 (사모채 발행사 연결용 법인등록번호) — 남은 예산으로 조금씩
        for code, corp in sorted(listed_cc.items()):
            if corp in info and not _stale(info[corp].get("checked"), 180):
                continue
            c = dart.company(corp)
            info[corp] = {k: c.get(k) for k in ("corp_name", "stock_name", "stock_code", "ceo_nm", "corp_cls",
                                                 "jurir_no", "adres", "hm_url", "induty_code", "est_dt", "acc_mt")}
            info[corp]["checked"] = now_kst().isoformat(timespec="minutes")

    except DartBudgetExceeded:
        diag["budget_exhausted"] = True
        log.info("호출 예산 소진 — 다음 실행에서 이어감")
    except DartQuotaExceeded as e:
        diag["quota_exceeded"] = redact(e)[:200]
    except Exception:  # noqa: BLE001 — 예상 못한 오류도 지금까지 받은 것은 저장하고 원인을 남긴다
        import traceback
        diag["crash"] = redact(traceback.format_exc())[-1500:]
        log.exception("프로필 수집 중 오류")
    finally:
        dump_json("ref/company_info.json", info)
        diag["dart_calls"] = dart.calls
