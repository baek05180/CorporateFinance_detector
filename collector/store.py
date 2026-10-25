"""JSON 파일 저장소. 데이터가 바뀌었을 때만 파일을 다시 쓴다 (불필요한 커밋 방지)."""
import json
import os
from datetime import datetime, timedelta

from . import config


def _path(name):
    return os.path.join(config.DATA_DIR, name)


def load_json(name, default):
    return _load(name, default)


def dump_json(name, obj, indent=1):
    """data/ 아래에 저장 (하위 폴더 자동 생성). 내용이 같으면 쓰지 않는다."""
    os.makedirs(os.path.dirname(_path(name)) or config.DATA_DIR, exist_ok=True)
    return _dump(name, obj, indent)


def _load(name, default):
    try:
        with open(_path(name), encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def _dump(name, obj, indent=1):
    os.makedirs(config.DATA_DIR, exist_ok=True)
    text = json.dumps(obj, ensure_ascii=False, indent=indent, sort_keys=True,
                      separators=(",", ":") if indent is None else None)
    p = _path(name)
    try:
        with open(p, encoding="utf-8") as f:
            if f.read() == text:
                return False
    except FileNotFoundError:
        pass
    with open(p, "w", encoding="utf-8") as f:
        f.write(text)
    return True


def now_kst():
    return datetime.now(config.KST)


class Store:
    def __init__(self):
        db = _load("records.json", {})
        self.records = db.get("records", {})
        self.meta = db.get("meta", {})
        self.seen = set(_load("seen.json", []))
        self.corp_cache = _load("corp_cache.json", {})
        self.diag = _load("diag.json", {})
        self.changed = False
        self.added = []

    # ── 레코드 반영 ──
    def upsert(self, rec, mode):
        rid = rec["id"]
        old = self.records.get(rid)
        stamp = now_kst().isoformat(timespec="minutes")
        if old:
            # 더 최근 공시(정정 포함)일 때만 덮어쓴다
            if (rec.get("rcept_no") or "") < (old.get("rcept_no") or ""):
                return
            rec["first_seen"] = old.get("first_seen", stamp)
            rec["captured_by"] = old.get("captured_by", mode)
            if old.get("original_rcept"):
                rec["original_rcept"] = old["original_rcept"]
            comparable_old = {k: v for k, v in old.items() if k not in ("updated", "confirmed")}
            comparable_new = {k: v for k, v in rec.items() if k not in ("updated", "confirmed")}
            if comparable_old == comparable_new:
                if mode == "daily" and not old.get("confirmed"):
                    old["confirmed"] = True
                    self.changed = True
                return
        else:
            rec["first_seen"] = stamp
            rec["captured_by"] = mode
            self.added.append(rid)
        rec["updated"] = stamp
        rec["confirmed"] = mode == "daily" or (old or {}).get("confirmed", False)
        self.records[rid] = rec
        self.changed = True

    def merge_rights_amendment(self, rec):
        """유상증자 정정공시 → 같은 회사의 직전 원공시를 대체한다."""
        if not rec.get("amended"):
            return rec
        cutoff = (now_kst() - timedelta(days=config.DETAIL_LOOKBACK_DAYS)).date().isoformat()
        candidates = [r for r in self.records.values()
                      if r["type"] == "RIGHTS" and r.get("corp_code") == rec.get("corp_code")
                      and (r.get("rcept_no") or "") < rec["rcept_no"] and (r.get("rcept_dt") or "") >= cutoff]
        if not candidates:
            return rec
        prev = max(candidates, key=lambda r: r["rcept_no"])
        rec["id"] = prev["id"]
        rec["original_rcept"] = prev.get("original_rcept") or prev["rcept_no"]
        return rec

    def mark_seen(self, rcept_no):
        self.seen.add(rcept_no)

    def reset_diag(self, key):
        self.diag[key] = {}

    def note_diag(self, key, value):
        bucket = self.diag.setdefault(key, {})
        bucket[value] = bucket.get(value, 0) + 1

    # ── 저장 ──
    def save(self, mode):
        wrote = False
        if self.changed:
            self.meta["data_updated"] = now_kst().isoformat(timespec="minutes")
        if mode == "daily":
            self.meta["last_daily"] = now_kst().isoformat(timespec="minutes")
        wrote |= _dump("records.json", {"meta": self.meta, "records": self.records})
        # seen은 최근 것만 유지 (파일이 끝없이 커지지 않게)
        wrote |= _dump("seen.json", sorted(self.seen)[-20000:])
        wrote |= _dump("corp_cache.json", self.corp_cache)
        wrote |= _dump("diag.json", self.diag)
        return wrote
