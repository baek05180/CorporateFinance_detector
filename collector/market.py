"""제안서용 시장 데이터: 한국은행 ECOS·FRED 금리, 공모채 수요예측 참여 내역."""
import os
import time
from datetime import timedelta

import requests

from . import docparse
from .dart import DartError
from .datagokr import redact
from .store import dump_json, load_json

ECOS = {  # 817Y002 시장금리(일별)
    "ktb3": "010200000", "ktb5": "010200001", "ktb10": "010210000", "aa3": "010300000", "bbb3": "010320000",
}
FRED = {"ust2": "DGS2", "ust10": "DGS10", "ffr_hi": "DFEDTARU", "ffr_lo": "DFEDTARL"}


def rates_step(today, diag, years=3):
    out = load_json("ref/mkt/rates.json", {})
    st = (today - timedelta(days=365 * years)).strftime("%Y%m%d")
    ek, fk = os.environ.get("ECOS_KEY"), os.environ.get("FRED_KEY")
    info = {}
    if ek:
        for nm, item in ECOS.items():
            try:
                u = f"https://ecos.bok.or.kr/api/StatisticSearch/{ek}/json/kr/1/2000/817Y002/D/{st}/{today:%Y%m%d}/{item}"
                rows = (requests.get(u, timeout=60).json().get("StatisticSearch") or {}).get("row") or []
                if rows:
                    out[nm] = [[f"{r['TIME'][:4]}-{r['TIME'][4:6]}-{r['TIME'][6:]}", float(r["DATA_VALUE"])] for r in rows]
                info[nm] = len(rows)
            except Exception as e:  # noqa: BLE001
                info[nm] = redact(e)[:120].replace(ek, "***")
        try:  # 기준금리 (일별)
            u = f"https://ecos.bok.or.kr/api/StatisticSearch/{ek}/json/kr/1/2000/722Y001/D/{st}/{today:%Y%m%d}/0101000"
            rows = (requests.get(u, timeout=60).json().get("StatisticSearch") or {}).get("row") or []
            if rows:
                out["base"] = [[f"{r['TIME'][:4]}-{r['TIME'][4:6]}-{r['TIME'][6:]}", float(r["DATA_VALUE"])] for r in rows]
            info["base"] = len(rows)
        except Exception as e:  # noqa: BLE001
            info["base"] = redact(e)[:120].replace(ek, "***")
    if fk:
        for nm, sid in FRED.items():
            try:
                j = requests.get("https://api.stlouisfed.org/fred/series/observations",
                                 params={"series_id": sid, "api_key": fk, "file_type": "json",
                                         "observation_start": f"{st[:4]}-{st[4:6]}-{st[6:]}"}, timeout=60).json()
                obs = [[o["date"], float(o["value"])] for o in j.get("observations", []) if o["value"] not in (".", "")]
                if obs:
                    out[nm] = obs
                info[nm] = len(obs)
            except Exception as e:  # noqa: BLE001
                info[nm] = redact(e)[:120].replace(fk, "***")
    out["updated"] = today.isoformat()
    dump_json("ref/mkt/rates.json", out, indent=None)
    diag["rates"] = info


def demand_step(dart, store, today, diag, days=200, max_docs=150):
    """최근 공모채의 [발행조건확정] 정정신고서 → 트랜치별 수요예측 참여 내역."""
    t0 = time.time()
    dem = load_json("ref/mkt/demand.json", {})
    since = (today - timedelta(days=days)).isoformat()
    todo = set()
    for r in store.records.values():
        if r.get("type") == "BOND" and r.get("offering") == "공모" and (r.get("rcept_dt") or "") >= since:
            for rno, nm in r.get("filings") or []:
                if "발행조건" in nm and (dem.get(rno) or {}).get("v") != docparse.DEMAND_DOC_VERSION:
                    todo.add(rno)
    todo = sorted(todo, reverse=True)[:max_docs]
    n_ok = 0
    for rno in todo:
        try:
            d = dart.document(rno)
        except DartError as e:
            dem[rno] = {"error": redact(e)[:100], "v": docparse.DEMAND_DOC_VERSION}
            continue
        main = max(d.values(), key=len) if d else ""
        tr = docparse.demand_doc(main)
        dem[rno] = {"tranches": tr, "v": docparse.DEMAND_DOC_VERSION}
        n_ok += bool(tr)
    dump_json("ref/mkt/demand.json", dem, indent=None)
    diag["demand"] = {"todo": len(todo), "parsed": n_ok, "total": len(dem), "sec": round(time.time() - t0)}


def mezz_events_step(dart, store, today, diag, back_days=400, max_docs=600):
    """거래소공시의 전환·행사·교환 청구, 가액 조정, 만기 전 취득 → 회차별 잔액·현재 가액 (오버행 계산용)."""
    import re as _re
    t0 = time.time()
    db = load_json("ref/mezz_events.json", {"filings": {}, "listed_to": None, "listed_from": None})
    live = {r["corp_code"] for r in store.records.values()
            if r.get("type") in ("CB", "BW", "EB") and (r.get("maturity") or "9") >= today.isoformat()}
    # 목록: 처음엔 back_days 전부(3개월씩), 이후엔 마지막으로 본 날부터
    start = date_from = today - timedelta(days=back_days)
    if db.get("listed_to") and db.get("listed_from") and db["listed_from"] <= date_from.isoformat():
        start = __import__("datetime").date.fromisoformat(db["listed_to"]) - timedelta(days=3)
    cands, d = [], start
    while d <= today:
        e = min(today, d + timedelta(days=89))
        try:
            for f in dart.list_filings(d.strftime("%Y%m%d"), e.strftime("%Y%m%d"), "I"):
                nm = (f.get("report_nm") or "").replace(" ", "")
                if docparse.MEZZ_EVENT_RE.search(nm) and "전환주식" not in nm and f.get("corp_code") in live:
                    cands.append(f)
        except DartError as ex:
            diag["mezz_events_list_error"] = redact(ex)[:120]
            break
        d = e + timedelta(days=1)
    else:
        db["listed_to"] = today.isoformat()
        db["listed_from"] = min(db.get("listed_from") or date_from.isoformat(), start.isoformat())
    cands += db.get("pending", [])  # 지난번 한도에 걸려 못 읽은 것
    todo = sorted({f["rcept_no"]: f for f in cands if (db["filings"].get(f["rcept_no"]) or {}).get("v") != docparse.MEZZ_EVENT_VERSION}.values(),
                  key=lambda f: f["rcept_no"], reverse=True)
    db["pending"] = [{k: f.get(k) for k in ("rcept_no", "corp_code", "report_nm")} for f in todo[max_docs:]]
    todo = todo[:max_docs]
    n_ok = 0
    for f in todo:
        try:
            dd = dart.document(f["rcept_no"])
        except DartError as ex:
            db["filings"][f["rcept_no"]] = {"corp": f["corp_code"], "error": redact(ex)[:80], "v": docparse.MEZZ_EVENT_VERSION}
            continue
        main = max(dd.values(), key=len) if dd else ""
        ev = docparse.mezz_event(main)
        n_ok += bool(ev)
        db["filings"][f["rcept_no"]] = {"corp": f["corp_code"], "nm": _re.sub(r"\s+", "", f.get("report_nm") or "")[:40],
                                       "ev": ev, "v": docparse.MEZZ_EVENT_VERSION}
    dump_json("ref/mezz_events.json", db, indent=None)
    diag["mezz_events"] = {"live_corps": len(live), "cands": len(cands), "todo": len(todo), "parsed": n_ok,
                           "total": len(db["filings"]), "pending": len(db["pending"]), "sec": round(time.time() - t0)}
