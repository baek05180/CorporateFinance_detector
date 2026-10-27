"""data/ → public/ (사이트 + 데이터 파일)

  python -m web.build   (프로젝트 루트에서)

public/index.html             화면 (한 페이지 앱, 주소 #/... 로 화면 전환)
public/data/meta.json         기준 시각
public/data/deals_recent.json 최근 120일 조달 (홈 화면)
public/data/deals.json        전체 조달 내역
public/data/companies.json    회사 검색용 목록
public/data/co/{key}.json     회사 화면·제안 보고서 데이터
"""
import json
import os
import shutil
import sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from collector import config  # noqa: E402
from collector.store import load_json, now_kst  # noqa: E402
from web import analysis as A  # noqa: E402

SITE_NAME = "ECM BOARD"
HERE = os.path.dirname(os.path.abspath(__file__))

KEEP = {
    "id", "type", "rcept_no", "rcept_dt", "url", "corp_name", "corp_cls", "stock_code", "amended",
    "series", "amount", "offering", "coupon", "ytm", "maturity", "board_dt", "pay_dt", "lead",
    "price", "shares", "dilution", "refix_floor", "exercise_from", "exercise_to", "exchange_target",
    "new_common", "new_other", "stock_kind", "method", "method_raw", "purposes", "name", "rating", "sub_dt",
    "underwriters", "guarantor", "coupon_raw", "isin", "kind", "issue_dt", "guarantee", "rank",
    "usage", "pcbo", "industry", "rate_type", "option", "uses", "detachable", "first_seen", "captured_by",
    "confirmed", "source", "demand_dt", "series_raw", "par_value",
}


OPTION_MATCH = {
    "public_bond": lambda d: d["type"] == "BOND" and d.get("offering") == "공모",
    "private_bond": lambda d: d["type"] == "BOND" and d.get("offering") == "사모" and not d.get("pcbo"),
    "pcbo": lambda d: d["type"] == "BOND" and bool(d.get("pcbo")),
    "cb": lambda d: d["type"] == "CB",
    "eb": lambda d: d["type"] == "EB",
    "rights": lambda d: d["type"] == "RIGHTS",
}


def write(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def jdump(obj):
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))


def main():
    today = now_kst().date()
    db = load_json("records.json", {"meta": {}, "records": {}})
    listed = load_json("ref/listed.json", {})
    corpcodes = load_json("ref/corpcodes.json", {}).get("map", {})
    info = load_json("ref/company_info.json", {})
    fin = load_json("fin/multi.json", {})
    prices_doc = load_json("ref/prices.json", {})
    prices = prices_doc.get("prices", {})
    ph = load_json("ref/price_hist.json", {})
    ph_dates, ph_close = ph.get("dates") or [], ph.get("close") or {}
    bond_docs = load_json("ref/bond_docs.json", {})
    outstanding = load_json("bonds/outstanding.json", {}).get("bonds", [])

    code_to_corp = corpcodes
    corp_to_code = {v: k for k, v in corpcodes.items()}
    jur_to_corp = {v.get("jurir_no"): k for k, v in info.items() if v.get("jurir_no")}
    multi_a = fin.get("annual", {})
    multi_i = fin.get("interim", {})

    # ── 조달 레코드 → 회사 키 ──
    def key_of(r):
        if r.get("corp_code"):
            return r["corp_code"]
        if r.get("crno") and r["crno"] in jur_to_corp:
            return jur_to_corp[r["crno"]]
        if r.get("stock_code") and r["stock_code"] in code_to_corp:
            return code_to_corp[r["stock_code"]]
        return "b" + (r.get("crno") or (r.get("corp_name") or "x").replace(" ", ""))

    rows = []
    for r in db["records"].values():
        row = {k: v for k, v in r.items() if k in KEEP and v not in (None, "", [], {})}
        row["key"] = key_of(r)
        if r.get("type") == "BOND" and r.get("offering") == "공모":
            doc = bond_docs.get(r.get("original_rcept") or r.get("rcept_no"), {})
            if doc.get("demand_dt"):
                row["demand_dt"] = doc["demand_dt"]
            if doc.get("rating") and not A.grade_of(row.get("rating")):
                row["rating"] = doc["rating"]
        if row.get("rating") in ("-", "—"):
            row.pop("rating")
        if row["key"] in info and not row.get("corp_cls"):
            row["corp_cls"] = info[row["key"]].get("corp_cls")
        rows.append(row)
    rows.sort(key=lambda x: A.deal_date(x), reverse=True)
    by_co = {}
    for r in rows:
        by_co.setdefault(r["key"], []).append(r)

    bonds_by_corp = {}
    for b in outstanding:
        k = jur_to_corp.get(b.get("crno")) or ("b" + b["crno"] if b.get("crno") else None)
        if k:
            bonds_by_corp.setdefault(k, []).append(b)

    # ── 회사 목록: 상장사 전체 + 조달 기록이 있는 모든 곳 ──
    keys = set(by_co)
    for code in listed:
        if code in code_to_corp:
            keys.add(code_to_corp[code])
    out_dir = os.path.join(config.PUBLIC_DIR, "data")
    if os.path.isdir(os.path.join(out_dir, "co")):
        shutil.rmtree(os.path.join(out_dir, "co"))

    index = []
    since_peer = (today - timedelta(days=365)).isoformat()
    pools = {ok: [] for ok in OPTION_MATCH}   # 비교 사례 후보: 최근 1년, 종류별
    for r in rows:
        if A.deal_date(r) < since_peer:
            continue
        code = r.get("stock_code") or corp_to_code.get(r["key"])
        meta = {"industry": (listed.get(code) or {}).get("industry") or r.get("industry"),
                "market": (listed.get(code) or {}).get("market") or r.get("corp_cls") or "E",
                "revenue": ((multi_a.get(code) or {}).get("cur") or {}).get("revenue"),
                "grade": A.grade_of(r.get("rating"))}
        for ok, fn in OPTION_MATCH.items():
            if fn(r):
                pools[ok].append((r, meta))

    for key in keys:
        deals = by_co.get(key, [])
        inf = info.get(key, {})
        code = inf.get("stock_code") or corp_to_code.get(key) or next((d.get("stock_code") for d in deals if d.get("stock_code")), None)
        kd = listed.get(code) if code else None
        prof = load_json(f"co/{key}.json", None) if not key.startswith("b") else None
        name = (kd or {}).get("name") or inf.get("stock_name") or inf.get("corp_name") or (deals[0].get("corp_name") if deals else key)
        market = (kd or {}).get("market") or inf.get("corp_cls") or next((d.get("corp_cls") for d in deals if d.get("corp_cls")), None) or "E"
        industry = (kd or {}).get("industry") or next((d.get("industry") for d in deals if d.get("industry")), None)
        series = A.financial_series(multi_a.get(code), fin.get("annual_year") or today.year - 1,
                                    multi_i.get(code), fin.get("interim_period"), prof)
        metrics = A.key_metrics(series, prof, prices.get(code) if code else None)
        bp = A.bond_profile(bonds_by_corp.get(key, []), today)
        rating = next((A.grade_of(d.get("rating")) for d in deals if d.get("rating") and A.grade_of(d.get("rating"))), None)
        rating_src = next((d.get("rating") for d in deals if d.get("rating") and A.grade_of(d.get("rating"))), None)
        hist = A.history_summary(deals, today)
        since3 = (today - timedelta(days=1095)).isoformat()
        options = A.propose({
            "market": market, "metrics": metrics, "rating": rating,
            "deals": [d for d in deals if A.deal_date(d) >= since3], "bonds": bp,
            "holders": (prof or {}).get("holders") or {}, "treasury_shares": (prof or {}).get("treasury_shares"),
        })
        # 비교 사례: 공모채는 같은 등급대, 나머지는 같은 업종 → 같은 시장·비슷한 매출 규모 (회사당 1건, 근거 표시)
        me = {"key": key, "industry": industry, "market": market, "grade": rating,
              "revenue": ((multi_a.get(code) or {}).get("cur") or {}).get("revenue") if code else metrics.get("revenue")}
        for o in options:
            if o["key"] not in OPTION_MATCH:
                continue
            o["comps"] = [{k: d.get(k) for k in ("id", "key", "corp_name", "type", "amount", "offering", "coupon", "price",
                                                 "rating", "pcbo", "series") if d.get(k) is not None}
                          | {"date": A.deal_date(d), "basis": basis}
                          for d, basis in A.pick_comps(o["key"], me, pools[o["key"]])]

        ai = load_json(f"summaries/{key}.json", None)
        biz = dict((prof or {}).get("business") or {})
        if biz.get("overview"):
            biz["overview_short"] = A.clip_sentences(biz["overview"], 320)
        if biz.get("customers"):
            biz["customers_short"] = A.clip_sentences(biz["customers"], 260)
        # 재무상태 그래프: 부채총계·차입금 (연간 3개년 + 최근 분기)
        debt_series = []
        fa, fi = (prof or {}).get("fs_annual") or {}, (prof or {}).get("fs_interim") or {}
        dh = fa.get("debt_hist") or [None, None, fa.get("debt")]
        ay = int(fa["period"][:4]) if fa.get("period") else None
        for r in series:
            if r["kind"] == "annual":
                d = dh[int(r["period"]) - ay + 2] if ay and 0 <= int(r["period"]) - ay + 2 <= 2 else None
                debt_series.append({"period": r["period"], "liab": r.get("liab"), "debt": d})
            elif r["kind"] == "interim":
                debt_series.append({"period": r["period"], "liab": r.get("liab"), "debt": fi.get("debt")})
        co = {
            "ai": ai, "debt_series": debt_series,
            "key": key, "name": name, "market": market, "stock_code": code, "industry": industry,
            "listing": kd, "info": {k: v for k, v in inf.items() if k not in ("checked",)} if inf else None,
            "business": biz, "profile_checked": (prof or {}).get("checked"),
            "annual_name": (prof or {}).get("annual_name"), "annual_rcept": (prof or {}).get("annual_rcept"),
            "holders": (prof or {}).get("holders"), "treasury_shares": (prof or {}).get("treasury_shares"),
            "fs_annual": (prof or {}).get("fs_annual"), "fs_interim": (prof or {}).get("fs_interim"),
            "series": series, "metrics": metrics, "rating": rating, "rating_src": rating_src,
            "bonds": bp, "deals": deals, "history": hist,
            "peers": A.peers(code, listed, multi_a, prices) if code else [],
            "proposal": options,
            "summary": A.summary_line(name, industry, metrics, hist, bp, options),
            "price_date": prices_doc.get("date"),
        }
        pc = ph_close.get(code) if code else None
        if pc and sum(v is not None for v in pc) >= 5:
            co["price_hist"] = [[d, v] for d, v in zip(ph_dates, pc) if v is not None]
        if not key.startswith("b") and not prof:
            co["pending"] = True
        write(os.path.join(out_dir, "co", f"{key}.json"), jdump(co))
        index.append([key, name, market, code or "", industry or "", len(deals), A.deal_date(deals[0]) if deals else ""])

    index.sort(key=lambda x: (-(x[5] > 0), x[1]))
    since = (today - timedelta(days=120)).isoformat()
    meta = {**db.get("meta", {}), "built": now_kst().isoformat(timespec="minutes"), "price_date": prices_doc.get("date"),
            "interim_period": fin.get("interim_period"), "annual_year": fin.get("annual_year")}
    write(os.path.join(out_dir, "meta.json"), jdump(meta))
    write(os.path.join(out_dir, "deals.json"), jdump(rows))
    write(os.path.join(out_dir, "deals_recent.json"), jdump([r for r in rows if A.deal_date(r) >= since]))
    write(os.path.join(out_dir, "companies.json"), jdump(index))

    with open(os.path.join(HERE, "template.html"), encoding="utf-8") as f:
        html = f.read().replace("__SITE_NAME__", SITE_NAME)
    write(os.path.join(config.PUBLIC_DIR, "index.html"), html)
    # Cloudflare Pages: 데이터는 자주 바뀌므로 캐시를 짧게
    write(os.path.join(config.PUBLIC_DIR, "_headers"), "/data/*\n  Cache-Control: public, max-age=300\n")
    print(f"deals={len(rows)} companies={len(index)} (프로필 {sum(1 for k in keys if os.path.exists(os.path.join(config.DATA_DIR, 'co', k + '.json')))})")


if __name__ == "__main__":
    main()
