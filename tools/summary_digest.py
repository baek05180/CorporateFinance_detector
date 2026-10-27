"""요약 재료를 짧게 추린다 (사람·AI가 빨리 읽도록).

  python -m tools.summary_digest 00105271 00545716 ...
표의 숫자 줄, 매출처·고객 문장, 개요 앞부분만 남긴다.
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from collector.store import load_json  # noqa: E402
from tools.summary_queue import build_packet  # noqa: E402

KEY = re.compile(r"매출처|거래처|고객|수출|내수|점유율|비중|비율|%|합계|합 계|소계")


def pick(text, limit, keep_all_lines=False):
    lines = [ln.strip() for ln in (text or "").split("\n") if ln.strip()]
    out, n = [], 0
    for ln in lines:
        if not keep_all_lines and not KEY.search(ln):
            continue
        ln = ln[:220]
        out.append(ln)
        n += len(ln)
        if n > limit:
            break
    return " / ".join(out)


SHORT = os.environ.get("DIGEST_SHORT") == "1"
LIM = dict(ov=300, pr=480, ma=90, sa=380) if SHORT else dict(ov=450, pr=650, ma=150, sa=650)


def digest(corp):
    listed = load_json("ref/listed.json", {})
    corp_to_code = {v: k for k, v in load_json("ref/corpcodes.json", {}).get("map", {}).items()}
    fin = load_json("fin/multi.json", {})
    prof = load_json(f"co/{corp}.json", {})
    recs = [r for r in load_json("records.json", {}).get("records", {}).values() if r.get("corp_code") == corp]
    p = build_packet(corp, prof, listed, corp_to_code, fin, recs)
    raw = prof.get("biz_raw") or {}
    return "\n".join([
        f"##### {p['corp_code']} {p['name']} | {p['market']} | {p['industry']} | KIND:{p['kind_products']} | 상장 {p['listed']} | {p['annual_rcept']} {p['annual_name']} | 재무(억) {p['financials_eok']}",
        "[개요] " + (raw.get("overview") or "")[:LIM["ov"]].replace("\n", " / "),
        "[제품] " + pick(raw.get("products"), LIM["pr"], keep_all_lines=True),
        "[원재료] " + (raw.get("materials") or "")[:LIM["ma"]].replace("\n", " / "),
        "[매출] " + pick(raw.get("sales"), LIM["sa"]),
    ])


if __name__ == "__main__":
    print("\n".join(digest(c) for c in sys.argv[1:]))
