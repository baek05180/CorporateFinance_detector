"""data/summaries/*.json 형식 검사.  python -m tools.summary_check  (문제가 있으면 종료 코드 1)"""
import glob
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from collector import config  # noqa: E402

LIMITS = {"one_liner": 220, "facts": 140, "mix_line": 80, "caption": 320, "business": 480, "customers": 260}


def check(path):
    errs = []
    try:
        d = json.load(open(path, encoding="utf-8"))
    except ValueError as e:
        return [f"JSON 오류: {e}"]
    corp = os.path.basename(path)[:-5]
    if d.get("corp_code") != corp:
        errs.append("corp_code가 파일명과 다름")
    for k in ("name", "source_rcept", "written", "one_liner"):
        if not d.get(k):
            errs.append(f"{k} 없음")
    for k, n in LIMITS.items():
        v = d.get(k)
        if v and len(v) > n:
            errs.append(f"{k} 너무 김 ({len(v)}>{n})")
        if v and v.rstrip().endswith(("…", "...")):
            errs.append(f"{k} 문장이 잘림")
    f = d.get("flow")
    if f is not None:
        if not isinstance(f.get("company"), dict) or not f["company"].get("name"):
            errs.append("flow.company 없음")
        for part, mx in (("inputs", 3), ("customers", 4), ("revenue", 6)):
            if len(f.get(part) or []) > mx:
                errs.append(f"flow.{part} {mx}개 초과")
            for x in f.get(part) or []:
                if not x.get("name"):
                    errs.append(f"flow.{part} 이름 없음")
        tot = sum((x.get("share") or 0) for x in f.get("revenue") or [])
        if tot > 105:
            errs.append(f"flow.revenue 비중 합 {tot:.1f}% > 100")
        for x in f.get("revenue") or []:
            if x.get("amount") is not None and not isinstance(x["amount"], (int, float)):
                errs.append("flow.revenue.amount는 숫자(원)")
    for p in d.get("products") or []:
        if not p.get("name"):
            errs.append("products 이름 없음")
    return errs


def main():
    bad = 0
    files = sorted(glob.glob(os.path.join(config.DATA_DIR, "summaries", "*.json")))
    for p in files:
        e = check(p)
        if e:
            bad += 1
            print(os.path.basename(p), "→", "; ".join(e))
    print(f"검사 {len(files)}건, 문제 {bad}건")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
