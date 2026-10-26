"""요약을 새로 써야 하는 회사의 '재료'를 뽑는다.

  python -m tools.summary_queue --limit 40 > queue.json
  python -m tools.summary_queue --corp 01275665          # 특정 회사

대상: 사업보고서 원문(biz_raw)이 수집됐는데 요약이 없거나, 요약이 예전 사업보고서 기준인 회사.
순서: watchlist → 최근 조달일이 가까운 회사.
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from collector import config  # noqa: E402
from collector.store import load_json  # noqa: E402

PART_LIMITS = {"overview": 2600, "products": 1500, "materials": 700, "sales": 2200}


def eok(v):
    return None if v is None else round(v / 1e8, 1)


def build_packet(corp, prof, listed, corp_to_code, fin, deals):
    code = corp_to_code.get(corp)
    kd = listed.get(code) or {}
    raw = prof.get("biz_raw") or {}
    multi = (fin.get("annual", {}).get(code) or {})
    cur = multi.get("cur") or prof.get("fs_annual") or {}
    return {
        "corp_code": corp,
        "name": kd.get("name") or (deals[0]["corp_name"] if deals else corp),
        "market": kd.get("market"), "industry": kd.get("industry"), "kind_products": kd.get("products"),
        "listed": kd.get("listed"),
        "annual_rcept": prof.get("annual_rcept"), "annual_name": prof.get("annual_name"),
        "financials_eok": {"revenue": eok(cur.get("revenue")), "op": eok(cur.get("op")), "ni": eok(cur.get("ni")),
                           "year": fin.get("annual_year"), "fs": multi.get("fs")},
        "raw": {k: (raw.get(k) or "")[:n] for k, n in PART_LIMITS.items() if raw.get(k)},
    }


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=30)
    ap.add_argument("--corp", action="append")
    args = ap.parse_args(argv)

    listed = load_json("ref/listed.json", {})
    corp_to_code = {v: k for k, v in load_json("ref/corpcodes.json", {}).get("map", {}).items()}
    fin = load_json("fin/multi.json", {})
    recs = load_json("records.json", {}).get("records", {})
    last, by_corp = {}, {}
    for r in recs.values():
        cc = r.get("corp_code")
        if not cc:
            continue
        d = r.get("board_dt") or r.get("sub_dt") or r.get("issue_dt") or r.get("rcept_dt") or ""
        last[cc] = max(last.get(cc, ""), d)
        by_corp.setdefault(cc, []).append(r)
    watch = set()
    try:
        from collector.profiles import read_watchlist, resolve_watchlist
        watch = set(resolve_watchlist(read_watchlist(), listed, load_json("ref/corpcodes.json", {}).get("map", {})))
    except Exception:  # noqa: BLE001
        pass

    co_dir = os.path.join(config.DATA_DIR, "co")
    todo = []
    for fn in os.listdir(co_dir) if os.path.isdir(co_dir) else []:
        corp = fn[:-5]
        if args.corp and corp not in args.corp:
            continue
        prof = load_json(f"co/{fn}", {})
        if not prof.get("biz_raw"):
            continue
        summ = load_json(f"summaries/{corp}.json", None)
        if summ and summ.get("source_rcept") == prof.get("annual_rcept") and not args.corp:
            continue
        prio = ("1" if corp in watch else "0") + last.get(corp, "")
        todo.append((prio, corp, prof))
    todo.sort(key=lambda x: x[0], reverse=True)
    out = [build_packet(c, p, listed, corp_to_code, fin, by_corp.get(c, [])) for _, c, p in todo[:args.limit]]
    json.dump({"remaining": max(0, len(todo) - len(out)), "items": out}, sys.stdout, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
