"""data/records.json → public/index.html, public/data.json

  python -m web.build   (프로젝트 루트에서)
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from collector import config  # noqa: E402

SITE_NAME = "ECM BOARD"
TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template.html")

KEEP = {
    "id", "type", "rcept_no", "rcept_dt", "url", "corp_name", "corp_cls", "stock_code", "amended",
    "series", "amount", "offering", "coupon", "ytm", "maturity", "board_dt", "pay_dt", "lead",
    "price", "shares", "dilution", "refix_floor", "exercise_from", "exercise_to", "exchange_target",
    "new_common", "new_other", "stock_kind", "method", "purposes", "name", "rating", "sub_dt",
    "underwriters", "guarantor", "coupon_raw", "isin", "kind", "issue_dt", "guarantee", "rank",
    "underwriter", "usage", "first_seen", "captured_by", "confirmed", "source",
}


def write_if_changed(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    try:
        with open(path, encoding="utf-8") as f:
            if f.read() == text:
                return False
    except FileNotFoundError:
        pass
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return True


def main():
    try:
        with open(os.path.join(config.DATA_DIR, "records.json"), encoding="utf-8") as f:
            db = json.load(f)
    except FileNotFoundError:
        db = {"meta": {}, "records": {}}

    rows = [{k: v for k, v in r.items() if k in KEEP and v not in (None, "", [], {})}
            for r in db["records"].values()]
    rows.sort(key=lambda r: (r.get("board_dt") or r.get("sub_dt") or r.get("issue_dt") or r.get("rcept_dt") or ""),
              reverse=True)
    meta = db.get("meta", {})
    payload = json.dumps({"meta": meta, "rows": rows}, ensure_ascii=False, separators=(",", ":"))

    with open(TEMPLATE, encoding="utf-8") as f:
        html = f.read()
    html = (html.replace("__SITE_NAME__", SITE_NAME)
                .replace("__DATA__", payload.replace("</", "<\\/")))

    a = write_if_changed(os.path.join(config.PUBLIC_DIR, "index.html"), html)
    b = write_if_changed(os.path.join(config.PUBLIC_DIR, "data.json"), payload)
    print(f"rows={len(rows)} index.html {'갱신' if a else '변경없음'} data.json {'갱신' if b else '변경없음'}")


if __name__ == "__main__":
    main()
