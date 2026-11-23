"""제안서용 시장 계산: 등급별 민평 곡선, 회사별 거래 스프레드, 같은 등급군 수요·사례·만기.

데이터
  data/ref/mkt/curves.json  금융투자협회 채권시가평가수익률(평가사 평균) — 국고채·회사채(공모 무보증) AAA~BBB-, 1~10년
  data/ref/mkt/trades.json  금융투자협회 일자별 장외 회사채 체결 (거래대금 1억원 이상)

기준 민평(만기별) 고르는 순서
  1) 팝업에 입력한 개별민평  → "개별민평"
  2) 최근 장외거래로 추정한 회사 고유 스프레드를 등급민평에 더한 값 → "추정 개별민평"
  3) 등급민평 → "등급민평"
"""
import json
import os
import re
from collections import defaultdict
from datetime import date, timedelta

DATA = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
ORDER = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-"]


def J(*p, default=None):
    f = os.path.join(DATA, *p)
    if not os.path.exists(f):
        return default
    with open(f, encoding="utf-8") as fh:
        return json.load(fh)


def gkey(g):
    """신용등급 표기 정리: 'AA0'·'AA' → 'AA', 'A0' → 'A', 공백·괄호 제거."""
    g = re.sub(r"[\s()]", "", str(g or "")).upper()
    if re.search(r"[1-9]", g):  # A1·A2+ 같은 단기(CP) 등급은 회사채 등급이 아님
        return ""
    g = g.replace("0", "")
    g = re.sub(r"[^A-Z+\-]", "", g)
    return g if g in ORDER else ""


def band(g):
    g = gkey(g)
    if not g:
        return ""
    return "AAA" if g == "AAA" else "AA" if g.startswith("AA") else "A" if g.startswith("A") else "BBB"


def in_band(g, b):
    return band(g) == b


def notch_diff(a, b):
    a, b = gkey(a), gkey(b)
    if not a or not b:
        return 99
    return abs(ORDER.index(a) - ORDER.index(b))


def tbucket(t):
    t = int(round(t or 0))
    return "2년" if t <= 2 else "3년" if t == 3 else "5년 이상"


def bucket_of(t):
    return tbucket(t)


class Curves:
    def __init__(self, raw=None):
        raw = raw or J("ref", "mkt", "curves.json", default={}) or {}
        self.tenors = raw.get("tenors") or [1, 1.5, 2, 2.5, 3, 4, 5, 7, 10]
        self.data = raw.get("data") or {}
        self.dates = sorted(self.data)
        self.source = raw.get("source", "금융투자협회 채권시가평가수익률")

    def __bool__(self):
        return bool(self.dates)

    @property
    def last(self):
        return self.dates[-1] if self.dates else None

    def asof(self, d=None):
        if not self.dates:
            return None
        if d is None:
            return self.dates[-1]
        d = d.isoformat() if isinstance(d, date) else d
        best = None
        for x in self.dates:
            if x <= d:
                best = x
            else:
                break
        return best

    def y(self, grade, tenor, d=None):
        """등급·만기 수익률(%). 만기 사이는 선형 보간, 1년 미만은 1년 값."""
        key = "KTB" if grade == "KTB" else gkey(grade)
        dd = self.asof(d)
        if not dd or not key:
            return None
        row = (self.data[dd] or {}).get(key)
        if not row:
            return None
        pts = [(t, v) for t, v in zip(self.tenors, row) if v is not None]
        if not pts:
            return None
        tenor = max(pts[0][0], min(pts[-1][0], float(tenor)))
        for (t0, v0), (t1, v1) in zip(pts, pts[1:]):
            if t0 <= tenor <= t1:
                return v0 + (v1 - v0) * (tenor - t0) / (t1 - t0) if t1 > t0 else v0
        return pts[-1][1]

    def sp(self, grade, tenor, d=None):
        a, k = self.y(grade, tenor, d), self.y("KTB", tenor, d)
        return (a - k) * 100 if a is not None and k is not None else None

    def series(self, grade, tenor, since=None):
        out = []
        for d in self.dates:
            if since and d < since:
                continue
            v = self.y(grade, tenor, d)
            if v is not None:
                out.append((d, v))
        return out

    def change(self, grade, tenor, days, d=None, spread=False):
        d1 = self.asof(d)
        if not d1:
            return None
        d0 = self.asof((date.fromisoformat(d1) - timedelta(days=days)).isoformat())
        f = self.sp if spread else (lambda g, t, x: (self.y(g, t, x) or 0) * 100)
        a, b = f(grade, tenor, d1), f(grade, tenor, d0)
        return a - b if a is not None and b is not None and d0 else None

    def pct_rank(self, grade, tenor, days=365, d=None):
        """최근 1년 스프레드 중 현재 위치 (0=최저, 100=최고)."""
        d1 = self.asof(d)
        since = (date.fromisoformat(d1) - timedelta(days=days)).isoformat()
        xs = [self.sp(grade, tenor, x) for x in self.dates if since <= x <= d1]
        xs = [x for x in xs if x is not None]
        now = self.sp(grade, tenor, d1)
        if not xs or now is None:
            return None
        return round(sum(1 for x in xs if x < now) / len(xs) * 100)

    def range(self, grade, tenor, days=365, d=None, spread=True):
        d1 = self.asof(d)
        since = (date.fromisoformat(d1) - timedelta(days=days)).isoformat()
        f = self.sp if spread else self.y
        xs = [(x, f(grade, tenor, x)) for x in self.dates if since <= x <= d1]
        xs = [(x, v) for x, v in xs if v is not None]
        if not xs:
            return None
        lo, hi = min(xs, key=lambda z: z[1]), max(xs, key=lambda z: z[1])
        return lo, hi


def wmedian(pairs):
    pairs = sorted(pairs)
    tot = sum(w for _, w in pairs)
    if not tot:
        return None
    acc = 0
    for v, w in pairs:
        acc += w
        if acc >= tot / 2:
            return v
    return pairs[-1][0]


def issuer_spread(name, grade, curves, isins=(), asof=None, days=30, raw=None):
    """최근 장외 체결 수익률 − 같은 잔존만기 등급민평 (bp). 거래대금 가중 중앙값.

    잔존 9개월 미만, 사모·후순위·영구, ±150bp 넘는 이상치는 뺀다. 3건·20억 미만이면 None."""
    raw = raw or J("ref", "mkt", "trades.json", default={}) or {}
    rows = raw.get("rows") or []
    if not rows or not curves:
        return None
    asof = asof or curves.last
    since = (date.fromisoformat(asof) - timedelta(days=days)).isoformat()
    nm = re.sub(r"\s", "", name or "")
    isins = set(isins or ())
    pick = []
    for d, bond, rem, _g, y, amt, isin in rows:
        if d < since or d > asof:
            continue
        b = re.sub(r"\s", "", bond)
        if not ((isin and isin in isins) or (nm and b.startswith(nm) and re.match(r"^\d", b[len(nm):] or "x"))):
            continue
        if rem < 0.75 or re.search(r"사모|후순위|신종|영구|CB|EB|BW", b):
            continue
        ref = curves.y(grade, rem, d)
        if ref is None:
            continue
        diff = (y - ref) * 100
        if abs(diff) > 150:
            continue
        pick.append((diff, amt, d, bond, rem, y))
    if len(pick) < 3 or sum(p[1] for p in pick) < 20:
        return None
    adj = wmedian([(p[0], p[1]) for p in pick])
    return {"bp": round(adj), "n": len(pick), "amt": round(sum(p[1] for p in pick)), "from": min(p[2] for p in pick),
            "to": max(p[2] for p in pick), "bonds": sorted({p[3] for p in pick}), "days": days}


def parse_indiv(s):
    """'2:3.95,3:4.10,5:4.40' → {2: 3.95, ...}"""
    out = {}
    for part in re.split(r"[,\s]+", str(s or "").strip()):
        m = re.match(r"^(\d{1,2})\s*[:=]\s*(\d+(?:\.\d+)?)%?$", part)
        if m:
            out[int(m.group(1))] = float(m.group(2))
    return out


def base_rates(tenors, grade, curves, indiv=None, adj=None, d=None):
    """만기별 기준 민평 → {t: (rate, 출처)}"""
    out = {}
    for t in tenors:
        if indiv and t in indiv:
            out[t] = (indiv[t], "개별민평")
        else:
            g = curves.y(grade, t, d)
            if g is None:
                out[t] = (None, "##")
            elif adj is not None:
                out[t] = (g + adj["bp"] / 100, "추정 개별민평")
            else:
                out[t] = (g, "등급민평")
    return out


def parse_band(s):
    """'개별민평 -30bp ~ +30bp' → (-30, 30). 못 읽으면 None."""
    txt = str(s or "")
    if "bp" not in txt.lower():
        return None
    nums = [int(round(float(x))) for x in re.findall(r"[-+]?\d+(?:\.\d+)?", txt)]
    if len(nums) >= 2:
        return min(nums[:2]), max(nums[:2])
    return None


def eok(v):
    """억원 단위 숫자 → '2.2조' / '420억'"""
    if v is None:
        return "-"
    return f"{v / 1e4:,.1f}조" if abs(v) >= 10000 else f"{v:,.0f}억"


# ───────── 공모채 수요·사례 (같은 등급군) ─────────
GROUPS = [("운용사", ["운용사"]), ("증권사", ["증권사"]), ("연기금·은행·보험", ["연기금·은행·보험"]),
          ("기타·외국", ["기타", "외국(실적유)", "외국(실적무)"])]


def demand_stats(pubs, grade, today, months=6, exact=False):
    """같은 등급군(또는 같은 등급) 최근 N개월 수요예측: 만기 구간별 건수·주문·경쟁률·가산금리·기관군."""
    b, g0 = band(grade), gkey(grade)
    since = (today - timedelta(days=int(months * 30.5))).isoformat()
    rows = [r for r in pubs if r.get("_dem") and r.get("_tenor") and (r.get("pay_dt") or "") >= since
            and (gkey(r.get("_grade")) == g0 if exact else in_band(r.get("_grade"), b))]
    st = defaultdict(lambda: {"n": 0, "orders": 0.0, "offer": 0.0, "ratios": [], "spreads": [], "groups": defaultdict(float), "deals": set()})
    for r in rows:
        k = tbucket(r["_tenor"])
        s = st[k]
        d = r["_dem"]
        s["n"] += 1
        s["deals"].add((r["corp_code"], r.get("original_rcept") or r.get("rcept_no")))
        s["orders"] += d.get("total") or 0
        if d.get("ratio"):
            s["ratios"].append(d["ratio"])
            s["offer"] += (d.get("total") or 0) / d["ratio"]
        if r.get("_spread") is not None:
            s["spreads"].append(r["_spread"])
        for lab, ks in GROUPS:
            s["groups"][lab] += sum((d.get("groups") or {}).get(x, 0) for x in ks)
    out = {}
    for k, s in st.items():
        sp = sorted(s["spreads"])
        out[k] = {"n": s["n"], "deals": len(s["deals"]), "orders": s["orders"], "offer": s["offer"],
                  "ratio": (s["orders"] / s["offer"]) if s["offer"] else None,
                  "avg_ratio": sum(s["ratios"]) / len(s["ratios"]) if s["ratios"] else None,
                  "under": sum(1 for x in s["ratios"] if x < 1), "sp_med": sp[len(sp) // 2] if sp else None,
                  "sp_lo": sp[0] if sp else None, "sp_hi": sp[-1] if sp else None, "groups": dict(s["groups"])}
    return out, rows


def comps(pubs, grade, today, days=150, exclude=None, limit=8):
    """같은 등급 ±1노치 최근 사례 (확정 가산금리가 있는 건)."""
    since = (today - timedelta(days=days)).isoformat()
    rs = [r for r in pubs if r.get("_spread") is not None and (r.get("pay_dt") or "") >= since
          and notch_diff(r.get("_grade"), grade) <= 1 and r.get("corp_code") != exclude]
    rs.sort(key=lambda r: (r.get("pay_dt") or "", r.get("corp_name"), r.get("series_raw") or ""), reverse=True)
    return rs[:limit]


def maturity_wall(mats, grade, start, end):
    """기간 내 같은 등급군 공모채 만기 (억원) — 기관 재투자 자금."""
    b = band(grade)
    tot, n = 0.0, 0
    for m in mats:
        if start.isoformat() <= (m.get("mat") or "") <= end.isoformat() and in_band(m.get("grade"), b):
            tot += m.get("amt") or 0
            n += 1
    return tot, n


def pipeline(pubs, grade, start, end, exclude=None):
    """기간 내 수요예측 예정인 같은 등급군 공모채 (경합 물량)."""
    b = band(grade)
    deals = defaultdict(float)
    for r in pubs:
        dd = r.get("_demand_dt") or ""
        if start.isoformat() <= dd <= end.isoformat() and in_band(r.get("_grade"), b) and r.get("corp_code") != exclude:
            deals[(r["corp_name"], dd)] += (r.get("amount") or 0) / 1e8
    return sorted(((k[0], k[1], v) for k, v in deals.items()), key=lambda x: x[1])


def last_year_same_month(pubs, grade, month, year):
    """작년 같은 달 같은 등급군 공모채 발행 (건수, 억원, 평균 경쟁률)."""
    b = band(grade)
    rs = [r for r in pubs if (r.get("pay_dt") or "").startswith(f"{year}-{month:02d}") and in_band(r.get("_grade"), b)]
    deals = {(r["corp_code"], r.get("original_rcept") or r.get("rcept_no")) for r in rs}
    rat = [r["_dem"]["ratio"] for r in rs if r.get("_dem") and r["_dem"].get("ratio")]
    return len(deals), sum((r.get("amount") or 0) for r in rs) / 1e8, (sum(rat) / len(rat) if rat else None)


def fmt_bp(x, sign=True):
    if x is None:
        return "-"
    return f"{x:+.0f}bp" if sign else f"{x:.0f}bp"


def ymd(d):
    d = date.fromisoformat(d) if isinstance(d, str) else d
    return f"{d.month}/{d.day}"
