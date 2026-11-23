"""공모사채 발행 검토 자료 v6 — 리서치형 DCM Discussion Materials.

  python -m tools.proposal6 --template 원본템플릿.pptx --inputs inputs.json --out out.pptx

구성 (docs/proposal_spec.md 의 deck_v6 에서 켜고 끔)
  Ⅰ. 예시 조건과 발행 윈도우: 요약 · Term Sheet · 만기 구조와 차환 · 발행월 수요 · 발행월 공급 · 일정 · 윈도우 비교
  Ⅱ. 기업 분석과 크레딧: 실적 동향 · 재무 레버리지와 등급 여유
  Ⅲ. 투자자 수요와 IR: 직전 회차 복기 · 투자자 구성 · IR 포인트와 Q&A
  Ⅳ. 금리 전망과 Market Update: 통화정책 · 리서치 금리전망 · 크레딧 시장 · 개별민평과 유통시장 · 조달비용 시산

모든 숫자는 데이터에서 계산한다. 확인이 필요한 값은 첫 장 노트의 '확인 목록'에 모은다.
"""
import argparse
import json
import os
import statistics
import sys
import tempfile
from collections import defaultdict
from datetime import date, timedelta
from types import SimpleNamespace

from pptx import Presentation
from pptx.util import Emu

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from tools import proposal as P3  # noqa: E402
from tools import proposal5 as P5  # noqa: E402
from tools import proposal_mkt as M  # noqa: E402
from tools.proposal_spec import load_spec  # noqa: E402
from tools.proposal_text import context, credit_story, kr_lines, ro, us_lines  # noqa: E402
from tools.proposal6_kit import (  # noqa: E402
    CW, DEEP, GAP, GREY, L, NAVY, PP_ALIGN, R, SUB, Y_BOT, MSO_ANCHOR, Page, Page5, box, chart, hline, rect, replace_tokens, tbl)

WD = "월화수목금토일"
SECS = [("Ⅰ", "예시 조건과 발행 윈도우"), ("Ⅱ", "기업 분석과 크레딧"), ("Ⅲ", "투자자 수요와 IR"), ("Ⅳ", "금리 전망과 Market Update")]


# ───────────────────────── 표기 ─────────────────────────
def md(d):
    d = date.fromisoformat(d) if isinstance(d, str) else d
    return f"{d.month}/{d.day}({WD[d.weekday()]})"


def mdd(d):
    d = date.fromisoformat(d) if isinstance(d, str) else d
    return f"{d.month}/{d.day}"


def ek(v, unit=True):
    """억원 숫자 → '1,500억원' / '2.98조원'."""
    if v is None:
        return "-"
    s = f"{v / 1e4:,.2f}조" if abs(v) >= 10000 else f"{v:,.0f}억"
    return s + ("원" if unit else "")


def won(v):
    return ek(v / 1e8) if v is not None else "-"


def bp(v, sign=True):
    if v is None:
        return "-"
    v = round(v)
    return f"{v:+d}bp" if sign else f"{v}bp"


def pc(v, d=1):
    return f"{v:.{d}f}%" if v is not None else "-"


def chg(a, b):
    return (b / a - 1) * 100 if a and b is not None and a > 0 else None


def sgn(v, d=1, unit="%"):
    return f"{v:+.{d}f}{unit}" if v is not None else "-"


def tn(t):
    return f"{int(round(t))}년"


# ───────────────────────── 계산 ─────────────────────────
def calc(F, I, X, spec):
    """페이지에 쓰는 숫자를 한곳에서 계산."""
    D = SimpleNamespace()
    D.check = []                                   # 확인 목록
    co = F.get("co") or {}
    issue, today, sch = F["issue"], F["today"], F["sch"]
    D.name, D.grade, D.band = F["name"], X.grade, X.band
    D.tenors, D.split, D.base_amt, D.max_amt = X.tenors, F["split"], F["base"], F["max"]
    D.issue, D.today, D.sch = issue, today, sch
    ev = {date.fromisoformat(k): v for k, v in (spec.get("events") or {}).items()}
    D.events = ev
    D.holi = {date.fromisoformat(k): v for k, v in (spec.get("holidays") or {}).items()}

    # 등급·전망
    ag = ((co.get("triggers") or {}).get("agencies")) or []
    D.agencies = ag
    outl = sorted({a.get("outlook") for a in ag if a.get("outlook")})
    D.rating_txt = f"{D.grade}({'/'.join(outl)})" if outl else D.grade
    D.rating_full = " · ".join(f"{a['agency']} {a.get('rating')}({a.get('outlook')})" for a in ag) or D.grade

    # 금리: 기준(개별민평/추정) + 목표 스프레드
    D.base = X.base
    D.tgt, D.exp = X.tgt, X.exp
    D.indiv_given = bool(X.indiv)
    if not D.indiv_given:
        D.check.append("개별민평: 추정치(등급민평 + 최근 30일 장외 체결 스프레드) — 민평 4사 개별민평으로 바꿔 입력")
    w = sum(D.split.values())
    D.wavg = sum(D.split[t] * (D.exp[t] or 0) for t in D.tenors) / w if all(D.exp[t] for t in D.tenors) else None
    D.wten = sum(D.split[t] * t for t in D.tenors) / w
    D.int_yr = D.base_amt * (D.wavg or 0) / 100
    D.bp1 = D.base_amt * 0.0001                     # 1bp = 연 억원
    D.cv = X.cv
    D.asof = X.asof

    # 회사 사채 잔액 · 만기
    import re as _re
    allb = [b for b in ((co.get("bonds") or {}).get("list") or []) if b.get("mat") and (b.get("bal") or b.get("amt"))]
    fx = [b for b in allb if _re.search(r"USD|EUR|JPY|CNY|HKD|SGD|AUD|외화", b.get("name") or "")]
    bl = [b for b in allb if b not in fx]
    D.fx = fx
    if fx:
        D.check.append(f"외화사채 {len(fx)}건은 원화 환산 금액이 없어 만기 표에서 제외 — " + ", ".join(b["name"] for b in fx[:3]))
    for b in bl:
        b["_eok"] = (b.get("bal") or b.get("amt")) / 1e8
    D.bonds = sorted(bl, key=lambda b: b["mat"])
    D.out_total = sum(b["_eok"] for b in bl)
    horizon = (issue + timedelta(days=150)).isoformat()
    D.targets = [b for b in D.bonds if issue.isoformat() <= b["mat"] <= horizon]
    D.target_amt = sum(b["_eok"] for b in D.targets)
    D.before = [b for b in D.bonds if today.isoformat() <= b["mat"] < issue.isoformat()]
    yr = issue.year
    D.mat_year = sum(b["_eok"] for b in D.bonds if b["mat"].startswith(str(yr)))
    D.mat_month = defaultdict(lambda: defaultdict(float))
    for b in D.bonds:
        if b["mat"].startswith(str(yr)):
            D.mat_month[int(b["mat"][5:7])][b.get("method") or "공모"] += b["_eok"]
    cp = [b.get("cpn") for b in D.targets if b.get("cpn")]
    D.target_cpn = (sum(b["cpn"] * b["_eok"] for b in D.targets if b.get("cpn")) / sum(b["_eok"] for b in D.targets if b.get("cpn"))) if cp else None
    D.check.append("CP·전자단기사채·외화사채 만기는 공시 데이터에 없음 — 회사 자금 계획으로 보완")

    # 기준금리·이벤트
    rt = F["rates"]
    D.base_rate = (rt.get("base") or [[None, None]])[-1][1]
    D.mpc = sorted(d for d, nm in ev.items() if "금통위" in nm)
    D.fomc = sorted(d for d, nm in ev.items() if "FOMC" in nm)
    d0 = sch["수요예측"]
    D.mpc_near = min(D.mpc, key=lambda d: abs((d - d0).days)) if D.mpc else None
    D.fomc_near = min(D.fomc, key=lambda d: abs((d - d0).days)) if D.fomc else None

    # 재무
    D.fin = finance(co)
    D.trig = (co.get("triggers") or {}).get("items") or []

    # 직전 회차들
    D.deals = deals(F.get("mine") or [])
    # 시장 통계
    D.dem, D.dem_x = X.dem, X.dem_x
    D.season = seasonality(F, X, spec, issue)
    D.supply = supply(F, X, issue)
    D.windows = windows(F, X, D, spec)
    D.outlook = X.outlook
    D.adj = X.adj
    D.trades = issuer_trades(D.name, D.grade, D.cv)
    fw = spec.get("fedwatch") or {}
    D.fedwatch = fw if fw.get("probs") else None
    if not D.fedwatch:
        D.check.append("FedWatch 확률: 명세 fedwatch 에 입력하면 美 통화정책 페이지에 표시")
    if not str(I.get("fee_bp") or "").strip():
        D.check.append("인수수수료 미입력 — 팝업에 입력 (v6 형식은 '협의'로 표기)")
    return D


def finance(co):
    """연간(최근 3년) + 분기 누적(당기·전년 동기) 손익과 재무 지표."""
    ser = co.get("series") or []
    fa, fi = co.get("fs_annual") or {}, co.get("fs_interim") or {}
    ann = [s for s in ser if s.get("kind") == "annual"][-3:]
    cur = next((s for s in ser if s.get("kind") == "interim"), None)
    prev = next((s for s in ser if s.get("kind") == "interim_prev"), None)
    dbt = {s["period"]: s.get("debt") for s in (co.get("debt_series") or [])}
    out = SimpleNamespace(ann=ann, cur=cur, prev=prev, rows={})

    def da(f):
        return (f.get("dep") or 0) + (f.get("amort") or 0) if f.get("dep") else None

    # EBITDA: 최근 연간(감가상각 있음) · 분기 누적
    out.ebitda_a = (fa.get("op") or 0) + da(fa) if da(fa) is not None and fa.get("op") is not None else None
    out.ebitda_i = (fi.get("op") or 0) + da(fi) if da(fi) is not None and fi.get("op") is not None else None
    # LTM: 연간 + 당기 누적 - 전년 동기 누적 (감가상각은 당기 누적을 연환산)
    out.ltm_op = (fa.get("op") + cur["op"] - prev["op"]) if fa.get("op") is not None and cur and prev and prev.get("op") is not None else None
    q = {"1Q": 1, "2Q": 2, "3Q": 3}.get((fi.get("period") or "")[-2:], 4)
    out.ltm_ebitda = (out.ltm_op + da(fi) * 4 / q) if out.ltm_op is not None and da(fi) is not None else None
    out.q = q

    def lev(debt, assets, cash, st, liab, eq, ebitda, op, intr):
        nd = debt - (cash or 0) - (st or 0) if debt is not None and cash is not None else None
        return {"debt": debt, "nd": nd,
                "dep": debt / assets * 100 if debt and assets else None,
                "ndep": nd / (assets - (cash or 0) - (st or 0)) * 100 if nd is not None and assets else None,
                "dr": liab / eq * 100 if liab and eq else None,
                "nd_eq": nd / eq * 100 if nd is not None and eq else None,
                "nd_ebitda": nd / ebitda if nd is not None and ebitda else None,
                "icr": op / intr if op and intr else None}
    yA = (fa.get("period") or "")[:4]
    a_ser = next((s for s in ann if s["period"] == yA), {})
    out.lev_a = lev(fa.get("debt"), fa.get("assets"), fa.get("cash"), fa.get("st_deposits"), a_ser.get("liab"),
                    a_ser.get("equity"), out.ebitda_a, fa.get("op"), fa.get("interest_paid"))
    ltm_int = (fi.get("interest_paid") or 0) * 4 / q if fi.get("interest_paid") else None
    out.lev_i = lev(fi.get("debt"), fi.get("assets"), fi.get("cash"), fi.get("st_deposits"), (cur or {}).get("liab"),
                    (cur or {}).get("equity"), out.ltm_ebitda, out.ltm_op, ltm_int)
    out.debt_hist = dbt
    out.ya, out.pi = yA, fi.get("period")
    out.ocf_a, out.capex_a = fa.get("ocf"), fa.get("capex")
    out.ocf_i, out.capex_i = fi.get("ocf"), fi.get("capex")
    return out


def deals(mine):
    """동사 공모채 회차별 묶음 (수요예측 결과·확정 스프레드)."""
    g = defaultdict(list)
    for r in mine:
        g[r.get("original_rcept") or r["rcept_no"]].append(r)
    out = []
    for k, rs in g.items():
        rs.sort(key=lambda r: r.get("series_raw") or "")
        d0 = rs[0].get("_demand_dt") or rs[0].get("pay_dt")
        out.append({"series": rs[0]["series_raw"].split("-")[0], "dt": d0, "pay": rs[0].get("pay_dt"), "tr": rs,
                    "amt": sum(r["amount"] for r in rs) / 1e8,
                    "orders": sum((r.get("_dem") or {}).get("total") or 0 for r in rs) if all(r.get("_dem") for r in rs) else None,
                    "offer": sum((r["_dem"]["total"] / r["_dem"]["ratio"]) for r in rs if (r.get("_dem") or {}).get("ratio"))
                    if all((r.get("_dem") or {}).get("ratio") for r in rs) else None})
    out.sort(key=lambda d: d["dt"] or "")
    return out


def spread_at(rt, d, a="aa3", b="ktb3"):
    """rates 시계열에서 d 이전 마지막 영업일 스프레드(bp)."""
    def at(key):
        v = None
        for dd, x in rt.get(key) or []:
            if dd <= d.isoformat():
                v = x
            else:
                break
        return v
    x, y = at(a), at(b)
    return (x - y) * 100 if x is not None and y is not None else None


def seasonality(F, X, spec, issue):
    """발행월 수요: 월별 발행(같은 등급군, 올해), 겨울별 스프레드 변화, 협회 집계(명세 입력)."""
    S = SimpleNamespace()
    m = issue.month
    yr0 = issue.year - 1 if m <= F["today"].month else issue.year
    S.month, S.year = m, yr0
    pubs = F["pubs"]
    b = X.band
    mon = defaultdict(lambda: {"amt": 0.0, "deals": set(), "ratios": [], "orders": 0.0})
    for r in pubs:
        dd = r.get("_demand_dt") or r.get("pay_dt") or ""
        if not dd.startswith(str(yr0)) or not M.in_band(r.get("_grade"), b):
            continue
        x = mon[int(dd[5:7])]
        x["amt"] += (r.get("amount") or 0) / 1e8
        x["deals"].add((r["corp_code"], r.get("original_rcept") or r.get("rcept_no")))
        if (r.get("_dem") or {}).get("ratio"):
            x["ratios"].append(r["_dem"]["ratio"])
    S.mon = {k: {"amt": v["amt"], "deals": len(v["deals"]), "ratio": statistics.mean(v["ratios"]) if v["ratios"] else None}
             for k, v in mon.items()}
    S.peak = max(S.mon, key=lambda k: S.mon[k]["amt"]) if S.mon else None
    tot = sum(v["amt"] for v in S.mon.values())
    S.share = S.mon.get(m, {}).get("amt", 0) / tot * 100 if tot else None
    others = [v["amt"] for k, v in S.mon.items() if k != m]
    S.mult = (S.mon.get(m, {}).get("amt", 0) / statistics.mean(others)) if others and statistics.mean(others) else None
    # 겨울별 AA- 3년 스프레드 변화 (발행월 1월 기준: 12월 말 → 1/15 → 1월 말)
    rt = F["rates"]
    S.winters = []
    if m in (1, 2):
        ys = sorted({int(d[:4]) for d, _ in (rt.get("aa3") or [])})
        for y in ys:
            a0 = spread_at(rt, date(y - 1, 12, 31))
            a1, a2 = spread_at(rt, date(y, 1, 15)), spread_at(rt, date(y, 1, 31))
            if a0 is not None and a2 is not None and (rt["aa3"][0][0] <= f"{y - 1}-12-20"):
                S.winters.append((y, a0, a1, a2))
    # 같은 등급(동사 등급) 3년 스프레드: 곡선 자료 범위 안에서
    cv = X.cv
    S.grade_w = None
    if cv and m in (1, 2):
        y = issue.year - 1
        try:
            s0, s1 = cv.sp(X.grade, 3, f"{y - 1}-12-31"), cv.sp(X.grade, 3, f"{y}-01-31")
            if s0 is not None and s1 is not None:
                S.grade_w = (y, s0, s1)
        except Exception:  # noqa: BLE001
            pass
    S.kofia = (spec.get("market_stats") or {}).get("jan_demand") if m == 1 else None
    S.wch = [(w[0], (w[2] - w[1]) if w[2] is not None else None, w[3] - w[1]) for w in S.winters]
    S.narrow = sum(1 for _, _, c in S.wch if c < 0)
    return S


def season_line(D):
    """발행월 수요를 데이터대로 한 문장으로."""
    S, m = D.season, D.season.month
    out = []
    if S.mult is not None:
        if S.mult >= 1.2:
            out.append(f"{S.year}년 {D.band}급 공모채 발행의 {S.share:.0f}%가 {m}월(다른 달 평균의 {S.mult:.1f}배)")
        elif S.peak and S.peak != m:
            pk = S.mon[S.peak]
            v_ = ek(S.mon.get(m, {}).get('amt'))
            out.append(f"{S.year}년 {D.band}급 발행은 {S.peak}월({ek(pk['amt'])})에 가장 많고 {m}월은 {v_}{ro(v_)} 경쟁 물량이 적은 달")
    if S.wch:
        lo, hi = min(c for _, _, c in S.wch), max(c for _, _, c in S.wch)
        out.append(f"AA- 3년 스프레드는 최근 {len(S.wch)}번의 겨울 중 {S.narrow}번 1월에 축소({lo:+.1f} ~ {hi:+.1f}bp)")
    return ". ".join(out)


def week_of(d):
    mon = d - timedelta(days=d.weekday())
    return mon


def supply(F, X, issue):
    """발행월 같은 등급군 만기(주별)와 작년 같은 달 주별 수요예측."""
    S = SimpleNamespace()
    b = X.band
    m0 = date(issue.year, issue.month, 1)
    m1 = date(issue.year + issue.month // 12, issue.month % 12 + 1, 1)
    wk = defaultdict(lambda: {"amt": 0.0, "names": defaultdict(float)})
    for x in F.get("mats") or []:
        d = date.fromisoformat(x["mat"])
        if m0 <= d < m1 and M.in_band(x.get("grade"), b):
            w = wk[week_of(d)]
            w["amt"] += x["amt"]
            w["names"][x["name"].split(" ")[0].rstrip("0123456789-")] += x["amt"]
    S.mat_weeks = sorted(wk.items())
    S.mat_total = sum(v["amt"] for v in wk.values())
    S.mat_issuers = len({n for _, v in wk.items() for n in v["names"]})
    S.my_week = week_of(F["sch"]["수요예측"])
    later = sum(v["amt"] for k, v in wk.items() if k > S.my_week)
    S.after_share = later / S.mat_total * 100 if S.mat_total else None
    # 작년 같은 달 주차별 수요예측 (같은 등급군)
    y = issue.year - 1
    dm = defaultdict(lambda: {"deals": set(), "amt": 0.0, "ratios": []})
    for r in F["pubs"]:
        dd = r.get("_demand_dt") or ""
        if dd.startswith(f"{y}-{issue.month:02d}") and M.in_band(r.get("_grade"), b):
            d = date.fromisoformat(dd)
            x = dm[(d.day - 1) // 7 + 1]
            x["deals"].add(r["corp_name"])
            x["amt"] += (r.get("amount") or 0) / 1e8
            if (r.get("_dem") or {}).get("ratio"):
                x["ratios"].append(r["_dem"]["ratio"])
    S.ly = {k: {"n": len(v["deals"]), "amt": v["amt"], "names": sorted(v["deals"]),
                "ratio": statistics.mean(v["ratios"]) if v["ratios"] else None} for k, v in sorted(dm.items())}
    S.ly_year = y
    S.pipe = X.pipe
    return S


def windows(F, X, D, spec):
    """발행 윈도우 후보 비교: 제안일 + 이후 달 같은 날짜 + (차환 대상 만기 직전)."""
    issue, base = F["issue"], F["base"]
    cands = [issue]
    for k in (1, 2):
        y, m = issue.year + (issue.month - 1 + k) // 12, (issue.month - 1 + k) % 12 + 1
        d = date(y, m, min(issue.day, 28))
        while d.weekday() >= 5 or d.isoformat() in P3.HOLI:
            d += timedelta(days=1)
        cands.append(d)
    if D.targets:
        mt = date.fromisoformat(D.targets[0]["mat"])
        d = P3.bday(mt, -2)
        if all(abs((d - c).days) > 10 for c in cands):
            cands.append(d)
    cands = sorted(set(cands))[:4]
    rate_park = D.base_rate                          # 선조달 자금 운용 금리 가정 = 기준금리
    out = []
    for d in cands:
        s = {"발행": d, "신고": P3.bday(d, -8), "수요예측": P3.bday(d, -6), "정정": P3.bday(d, -4)}
        mt = date.fromisoformat(D.targets[0]["mat"]) if D.targets else None
        days = (mt - d).days if mt else None
        cost = base * ((D.wavg or 0) - (rate_park or 0)) / 100 * days / 365 if days and days > 0 and D.wavg and rate_park is not None else 0
        be = cost / (base * D.wten) * 1e4 if cost else 0
        mp = [x for x in D.mpc if abs((x - s["수요예측"]).days) <= 10]
        fo = [x for x in D.fomc if abs((x - s["수요예측"]).days) <= 10]
        ly = M.last_year_same_month(F["pubs"], X.grade, d.month, d.year - 1)
        wall, _ = M.maturity_wall(F["mats"], X.grade, date(d.year, d.month, 1),
                                  date(d.year + d.month // 12, d.month % 12 + 1, 1) - timedelta(days=1))
        out.append({"s": s, "days": days, "cost": cost, "be": be, "mpc": mp, "fomc": fo, "ly": ly, "wall": wall,
                    "chosen": d == issue})
    return out


def issuer_trades(name, grade, cv, days=45):
    raw = M.J("ref", "mkt", "trades.json", default={}) or {}
    rows = [r for r in raw.get("rows") or [] if (r[1] or "").startswith(name)]
    out = []
    for d, nm, yrs, g, y, amt, isin in sorted(rows, reverse=True):
        try:
            ref = cv.y(grade, yrs, d) if cv else None
        except Exception:  # noqa: BLE001
            ref = None
        out.append({"d": d, "nm": nm, "yrs": yrs, "y": y, "amt": amt, "vs": (y - ref) * 100 if ref else None})
    return out


# ───────────────────────── 페이지 ─────────────────────────
class Deck:
    def __init__(self, prs, lay):
        self.prs, self.lay = prs, lay
        self.pages = []          # (slide, key, 섹션 번호, 제목)
        self.sec = 0
        self.n = 0

    def section(self, i):
        self.sec, self.n = i, 0

    def page(self, key, title, headline=None, bullets=None):
        self.n += 1
        sl = self.prs.slides.add_slide(self.lay)
        p = Page(sl, f"{SECS[self.sec][0]}. {SECS[self.sec][1]}", f"{self.n}. {title}", headline, bullets, key)
        self.pages.append((sl, key, self.sec, title))
        return p


def p_summary(dk, D, I, X):
    t = D.tenors
    tgt = D.targets[0] if D.targets else None
    mpc = D.mpc_near
    d0, iss = D.sch["수요예측"], D.issue
    what = (f"{md(tgt['mat'])} 만기 {ek(tgt['_eok'])}" + (f" 등 {ek(D.target_amt)}" if len(D.targets) > 1 else "")) if tgt else f"{iss.month}월 {ek(D.base_amt)} 신규 조달"
    why = []
    if (D.season.mult and D.season.mult > 1.2) or D.season.narrow:
        why.append("연초 수요" if iss.month == 1 else f"{iss.month}월 수요")
    if mpc:
        why.append("금통위 일정")
    hl = (f"{what}{' 차환' if tgt else ''}을 {md(d0)} 수요예측·{md(iss)} 납입으로 조달하는 예시: "
          f"{'와 '.join(why) or '수급과 금리 일정'}에 맞춘 윈도우")
    F_ = D.fin
    la, li = F_.lev_a, F_.lev_i
    up = next((x for x in D.trig if x.get("dir") == "up" and x.get("thr") is not None), None)
    cred = f"{D.rating_full}."
    if li.get("nd_ebitda"):
        cred += f" 순차입금/EBITDA {li['nd_ebitda']:.1f}배(9월말, LTM 근사)"
        if up:
            cred += f", {up['agency']} 상향 기준 {up['metric']} {up['op']}{up['thr']:g}{up['unit']}"
    tens = " + ".join(f"{tn(x)} {ek(D.split[x])}" for x in t)
    pri = (X.alloc or {}).get("priority") or I.get("priority")
    bl = [
        ("예시 조건", f"신고 {ek(D.base_amt)}(한도 {ek(D.max_amt)}), {tens}, 밴드 {X.band_bp[0]:+d} ~ {X.band_bp[1]:+d}bp"
                   + (f", 증액 시 {pri}년물 우선" if pri else "") + " (p.{p:terms})"),
        ("크레딧", cred + " (p.{p:credit})"),
    ]
    S = D.season
    sl_ = season_line(D)
    if sl_:
        bl.append((f"{iss.month}월 수요", sl_ + " (p.{p:season})"))
    rel = ""
    if mpc:
        rel = f"{md(mpc)} 금통위 {'다음' if mpc < d0 else '전'} {'수요예측' if mpc < d0 else '금리 확정'}"
    bl.append(("윈도우", f"{md(d0)} 수요예측·{md(iss)} 납입" + (f", {rel}" if rel else "")
               + (f". {md(tgt['mat'])} 만기({tgt['name']}) {(date.fromisoformat(tgt['mat']) - iss).days}일 전 조달" if tgt else "")
               + " (p.{p:sched})"))
    p = dk.page("summary", "Executive Summary", hl)
    # 왼쪽 요점 / 오른쪽 조건 요약
    lw = 5450000
    yb = p.bullets(bl, x=L, w=lw, size=10, y=p.y)
    x2 = L + lw + 180000
    w2 = R - x2
    y2 = p.panel(x2, p.y, w2, "조건 요약", "(단위: 억원, %)")
    rows = [["구분", "내용"],
            ["발행물", "무보증 공모사채"],
            ["등급", D.rating_txt],
            ["규모", f"신고 {D.base_amt:,} (한도 {D.max_amt:,})"],
            ["트랜치", " / ".join(f"{x}년 {D.split[x]:,}" for x in t)],
            ["밴드", f"{'개별민평' if D.indiv_given else '개별민평'} {X.band_bp[0]:+d} ~ {X.band_bp[1]:+d}bp"],
            ["기준 금리", " / ".join(f"{x}년 {D.base[x][0]:.3f}" for x in t if D.base[x][0]) + (" (추정)" if not D.indiv_given else "")],
            ["수요예측", md(d0)],
            ["발행·납입", md(iss)],
            ["자금용도", "채무상환" + (f" ({tgt['name']} 차환 등)" if tgt else "")]]
    yt = tbl(p.s, rows, x2, y2, w2, [1, 2.6], rowh=[240000] + [250000] * (len(rows) - 1), size=8.5)
    y = max(yb, yt) + 200000
    # 숫자 칸
    d6 = D.dem.get(M.tbucket(3)) or next(iter(D.dem.values()), {})
    k = [(pc(D.wavg, 2) if D.wavg else "-", f"예상 가중평균 금리\n(기준 금리 + 목표 스프레드)"),
         (f"{d6.get('avg_ratio'):.1f}배" if d6.get("avg_ratio") else "-", f"{D.band}급 최근 6개월\n평균 경쟁률({M.tbucket(3)})"),
         (ek(D.mat_year), f"{D.issue.year}년 공모·사모 사채 만기\n(사채 잔액 {ek(D.out_total)})"),
         (f"{li['nd_ebitda']:.1f}배" if li.get("nd_ebitda") else "-", f"순차입금/EBITDA\n({F_.pi}, LTM 근사)")]
    ins_h = 650000 + 230000 * 3
    y = y + max(0, (p.ybot - ins_h - y - 820000) // 2)
    y = p.kpis(y, k, h=820000) + 200000
    # 금리 여건
    O = D.outlook
    items = []
    if O:
        from tools import proposal_outlook as OL
        b_ = OL.series(O, "KR", "base")
        gp, how = OL.grade_path(O, D.grade, 3, D.cv)
        if b_:
            vals = [f"{lab} {v:.2f}%" for (lab, _), v in zip(O["cols"][:4], b_[:4]) if v is not None]
            items.append(("리서치 전망", f"기준금리 {' → '.join(vals)}"
                          + (f", {D.grade} 3년 {' → '.join(f'{lab} {v:.2f}%' for (lab, _), v in zip(O['cols'][1:5], gp[1:5]) if v is not None)}" if gp else "")
                          + " (p.{p:outlook})"))
    if D.fedwatch:
        pr = D.fedwatch.get("probs") or {}
        items.append(("FedWatch", f"({D.fedwatch.get('asof')}) " + ", ".join(f"{k_} {v}" for k_, v in list(pr.items())[:3])))
    if D.mpc_near:
        items.append(("일정", f"금통위 {md(D.mpc_near)}" + (f", FOMC {md(D.fomc_near)}" if D.fomc_near else "")
                      + f". 발행금리는 청약 1영업일 전 개별민평에 확정 스프레드를 더해 결정"))
    p.insight("금리 여건", items, size=9.5)
    p.foot("조건과 일정은 예시. 기준 금리는 " + ("입력한 개별민평" if D.indiv_given else "등급민평에 최근 30일 동사 장외 체결 스프레드를 더한 추정치")
           + ", 목표 스프레드는 동사 직전 확정값과 같은 등급군 6개월 중앙값의 평균. 순차입금/EBITDA는 LTM 근사(당사 산출).",
           "금융투자협회 채권정보센터, 전자공시시스템, 신용평가 3사, 신한투자증권 리서치센터, 신한투자증권.")
    return p


def p_terms(dk, D, I, X):
    t = D.tenors
    tgt = D.targets[0] if D.targets else None
    prev = D.deals[-1] if D.deals else None
    fee = f"{I['fee_bp']}bp" if str(I.get("fee_bp") or "").strip() else "협의"
    hl = f"신고 {ek(D.base_amt)}, {'·'.join(tn(x) for x in t)}, 개별민평 {X.band_bp[0]:+d} ~ {X.band_bp[1]:+d}bp"
    bl = [("신고 " + ek(D.base_amt) + ", 증액 한도 " + ek(D.max_amt),
           (f"{tgt['name']}({md(tgt['mat'])} 만기, {ek(tgt['_eok'])})" + (" 등 " + f"{len(D.targets)}건 {ek(D.target_amt)}" if len(D.targets) > 1 else "") + " 차환 규모" if tgt else "신규 조달 규모")
           + f". 증액 한도는 신고의 {D.max_amt / D.base_amt:.1f}배"),
          ("밴드 " + f"{X.band_bp[0]:+d} ~ {X.band_bp[1]:+d}bp",
           (f"직전 {prev['series']}회와 같은 폭. " if prev and prev["tr"][0].get("_band") else "")
           + " · ".join(f"{x}년 {D.base[x][0]:.3f}%" for x in t if D.base[x][0]) + f"({mdd(D.asof)} {'개별민평' if D.indiv_given else '추정 개별민평'})")]
    p = dk.page("terms", "예시 발행 조건: Term Sheet", hl, bl)
    lw = 5300000
    y = p.panel(L, p.y, lw, "Term Sheet", "(단위: 억원, bp)")
    rows = [["항목", "내용"],
            ["발행회사", f"(주){D.name}"],
            ["사채 종류", "무기명식 이권부 무보증 공모사채"],
            ["예상 신용등급", f"{D.rating_txt}, 신고 전 본평가"],
            ["발행 예정금액", f"신고 {D.base_amt:,}억원 (증액 한도 {D.max_amt:,}억원)"],
            ["트랜치", " / ".join(f"{x}년 {D.split[x]:,}억원" for x in t)],
            ["공모희망금리", f"개별민평 {X.band_bp[0]:+d}bp ~ {X.band_bp[1]:+d}bp (청약 1영업일 전 민평 기준)"],
            ["수요예측", f"{D.sch['수요예측'].isoformat().replace('-', '.')}({WD[D.sch['수요예측'].weekday()]}) 09:00~16:00"],
            ["발행·납입", f"{D.issue.isoformat().replace('-', '.')}({WD[D.issue.weekday()]})"],
            ["자금용도", "채무상환" + (f" ({tgt['name']} {ek(tgt['_eok'])} 차환 등)" if tgt else "")],
            ["인수수수료", fee],
            ["대표주관회사", "신한투자증권"]]
    tbl(p.s, rows, L, y, lw, [1.1, 3.2], rowh=[240000] + [262000] * (len(rows) - 1), size=8.5,
        align_cols={1: PP_ALIGN.LEFT})
    x2 = L + lw + 200000
    w2 = R - x2
    y2 = p.panel(x2, p.y, w2, "조건 근거")
    pri = (X.alloc or {}).get("priority") or I.get("priority")
    sh = {}
    for d in D.deals:
        for r in d["tr"]:
            if r.get("_tenor"):
                sh[int(round(r["_tenor"]))] = sh.get(int(round(r["_tenor"])), 0) + r["amount"] / 1e8
    tot_sh = sum(sh.values()) or 1
    main_t = max(sh, key=sh.get) if sh else None
    items = [
        ("① 금액", f"{ek(D.target_amt)} 만기 차환 + 증액 여지 {ek(D.max_amt - D.base_amt)} (p.{{p:mat}})" if tgt else f"신고 {ek(D.base_amt)}"),
        ("② 트랜치", " / ".join(f"{x}년 {D.split[x]:,}억원" for x in t)
         + (f". 동사 공모채 발행액의 {sh[main_t] / tot_sh * 100:.0f}%가 {main_t}년물" if main_t else "")
         + (" — 수요·비용·만기 분산·전망으로 자동 배분" if X.alloc else "")),
        ("③ 밴드", (f"직전 회차와 같은 수준. 확정 스프레드 " + " / ".join(bp(r.get('_spread')) for r in D.deals[-1]["tr"] if r.get('_spread') is not None)
                   if D.deals and any(r.get('_spread') is not None for r in D.deals[-1]['tr']) else "같은 등급군 관행") + " (p.{p:prev})"),
        ("④ 일정", f"{md(D.sch['수요예측'])} 수요예측, {md(D.issue)} 납입" + (f" — 금통위 {md(D.mpc_near)} {'다음' if D.mpc_near < D.sch['수요예측'] else '전'}" if D.mpc_near else "") + " (p.{p:sched})"),
        ("⑤ 증액", (f"수요예측 결과를 보고 {md(D.sch['정정'])} 조건 확정 전 판단" + (f", {pri}년물 우선" if pri else ""))
         + (f". 직전 {D.deals[-1]['series']}회는 모집 {ek(D.deals[-1]['offer'])} → 발행 {ek(D.deals[-1]['amt'])}" if D.deals and D.deals[-1].get("offer") else "")),
    ]
    ps = []
    for h, txt in items:
        ps.append({"text": h, "size": 10.5, "bold": True, "color": NAVY, "after": 1})
        ps.append({"text": txt, "size": 9.5, "after": 9})
    box(p.s, x2, y2, w2, p.ybot - y2, ps, margin=40000)
    p.foot("조건은 직전 회차 증권신고서를 참고한 예시이며, 규모·트랜치·밴드는 발행사 자금 계획과 시장 상황에 따라 조정 가능.",
           "전자공시시스템, 금융투자협회 채권정보센터, 신한투자증권.")


def p_maturity(dk, D, I, X):
    yr = D.issue.year
    tg = D.targets
    hl = (f"{md(tg[0]['mat'])} 만기 {ek(tg[0]['_eok'])}" + (f" 등 {len(tg)}건 {ek(D.target_amt)}" if len(tg) > 1 else "")
          + f"이 발행 후 150일 안에 도래: {yr}년 사채 만기 {ek(D.mat_year)}") if tg else f"{yr}년 사채 만기 {ek(D.mat_year)}"
    q = defaultdict(float)
    for mth, v in D.mat_month.items():
        q[(mth - 1) // 3 + 1] += sum(v.values())
    bl = [(f"{yr}년 사채 만기 {ek(D.mat_year)}", " / ".join(f"{k}분기 {ek(q[k])}" for k in (1, 2, 3, 4) if q[k])
           + f". 사채 잔액 {ek(D.out_total)}({len(D.bonds)}건)")]
    if D.before:
        bl.append(("발행 전 만기", ", ".join(f"{md(b['mat'])} {b['name']} {ek(b['_eok'])}" for b in D.before[:3]) + " — 보유 현금 또는 별도 재원으로 상환"))
    if tg and D.target_cpn and D.wavg:
        bl.append(("금리 차이", f"차환 대상 가중평균 쿠폰 {D.target_cpn:.3f}% 대비 이번 예상 금리 {D.wavg:.2f}%({(D.wavg - D.target_cpn) * 100:+.0f}bp), "
                   f"연 이자 {ek(abs(D.wavg - D.target_cpn) * D.target_amt / 100)} {'증가' if D.wavg > D.target_cpn else '감소'}"))
    p = dk.page("mat", f"{yr}년 만기 구조와 차환", hl, bl)
    lw = 5200000
    y = p.panel(L, p.y, lw, f"향후 12개월 사채 만기", "(단위: 억원, %)")
    hz = (D.today + timedelta(days=365)).isoformat()
    rows = [["만기일", "회차", "구분", "금액", "쿠폰", "비고"]]
    for b in [b for b in D.bonds if D.today.isoformat() <= b["mat"] <= hz][:11]:
        note = "차환 대상" if b in tg else ("발행 전 만기" if b in D.before else "")
        rows.append([md(b["mat"]), b["name"].replace(D.name, ""), b.get("method") or "-", f"{b['_eok']:,.0f}",
                     f"{b['cpn']:.3f}" if b.get("cpn") else "-", note])
    rows.append(["합계", "", "", f"{sum(float(r[3].replace(',', '')) for r in rows[1:]):,.0f}", "", ""])
    fills = {(i, c): "EAF2F8" for i, r in enumerate(rows) if r[5] == "차환 대상" for c in range(6)}
    tbl(p.s, rows, L, y, lw, [1.0, 0.9, 0.6, 0.8, 0.7, 1.0], rowh=[230000] + [215000] * (len(rows) - 1), size=8,
        fills=fills, bold_cells={(len(rows) - 1, c) for c in range(6)})
    x2 = L + lw + 200000
    w2 = R - x2
    y2 = p.panel(x2, p.y, w2, f"{yr}년 월별 사채 만기", "(단위: 억원)")
    cats = [f"{m}월" for m in range(1, 13)]
    kinds = sorted({k for v in D.mat_month.values() for k in v}, key=lambda k: {"공모": 0, "사모": 1}.get(k, 2))
    ser = [(k, [round(D.mat_month.get(m, {}).get(k, 0)) or None for m in range(1, 13)]) for k in kinds] or [("공모", [0] * 12)]
    iy, _ = p.insight("차환 재원과 일정", [
        ("이번 발행", f"{ek(D.base_amt)}(한도 {ek(D.max_amt)})은 {', '.join(b['name'] for b in tg[:3])} 만기 차환과 증액분 선조달 여지" if tg else f"{ek(D.base_amt)} 신규 조달"),
        ("확인 필요", "CP·전자단기사채·외화사채·은행 차입 만기는 공시 사채 데이터에 없음 — 회사 자금 계획으로 보완"),
    ], x=x2, w=w2)
    chart(p.s, "colstack" if len(ser) > 1 else "col", x2, y2, w2, iy - y2 - 120000, cats, ser, legend=len(ser) > 1)
    p.foot("사채는 채권 기본정보와 증권신고서 기준(공모·사모, 주식연계사채 제외). 금액은 잔액.",
           "전자공시시스템, 한국예탁결제원(채권 기본정보), 신한투자증권.")


def p_season(dk, D, I, X):
    S = D.season
    m = S.month
    hl = season_line(D).replace(f"{S.year}년 ", "", 1) or f"{m}월 수요"
    bl = []
    if S.mon:
        mx = S.mon.get(m, {})
        pk = S.mon.get(S.peak, {})
        bl.append((f"{S.year}년 {m}월 {D.band}급 발행 {ek(mx.get('amt'))}({mx.get('deals', 0)}개 딜)",
                   f"연간의 {S.share:.0f}%, 다른 달 평균의 {S.mult:.1f}배" + (f". 가장 많은 달은 {S.peak}월 {ek(pk.get('amt'))}({pk.get('deals')}개 딜)" if S.peak != m else "")))
    if S.wch:
        bl.append(("스프레드는 연초에 축소되는 경향", "AA- 3년(국고 대비) 12월 말 → 1월 말 " + ", ".join(f"{y}년 {c:+.1f}bp" for y, _, c in S.wch)
                   + (f". 1/15까지 평균 {statistics.mean([c for _, c, _ in S.wch if c is not None]):+.1f}bp" if any(c is not None for _, c, _ in S.wch) else "")))
    if S.grade_w:
        y, s0, s1 = S.grade_w
        bl.append((f"{D.grade} 3년", f"{y}년 1월 스프레드 {s0:.1f}bp → {s1:.1f}bp({s1 - s0:+.1f}bp)"))
    if S.kofia:
        rows_ = S.kofia.get("rows") or []
        if rows_:
            bl.append(("매년 반복되는 1월 수요", "금융투자협회 집계 1월 회사채 수요예측 참여금액 " + ", ".join(f"{r[0]}년 {r[2]}조원" for r in rows_[-3:])))
    p = dk.page("season", "1월 수요와 연초효과" if m == 1 else f"{m}월 수요와 계절성", hl, bl)
    lw = 5200000
    y = p.panel(L, p.y, lw, f"{S.year}년 월별 {D.band}급 공모채 발행", "(단위: 억원)")
    ks = list(range(1, 13))
    vals = [round(S.mon.get(k, {}).get("amt", 0)) or None for k in ks]
    dd_n = len([d for d in D.deals if d.get("dt")][-4:])
    reserve = (310000 + 215000 * (dd_n + 1) + 120000) if dd_n else 0
    ch_h = max(1300000, min(2500000, p.ybot - y - reserve))
    c = chart(p.s, "col", L, y, lw, ch_h, [f"{k}월" for k in ks], [("발행", vals)], legend=False)
    c_gf = p.s.shapes[-1]
    try:
        pt = c.plots[0].series[0].points[m - 1]
        pt.format.fill.solid()
        from pptx.dml.color import RGBColor
        pt.format.fill.fore_color.rgb = RGBColor.from_string("F2A900")
    except Exception:  # noqa: BLE001
        pass
    x2 = L + lw + 200000
    w2 = R - x2
    y2 = p.y
    if S.winters:
        y2 = p.panel(x2, y2, w2, "AA- 3년 신용스프레드: 겨울별 변화", "(단위: bp)")
        rows = [["겨울", "12월 말", "1/15", "1월 말", "변화"]]
        for yy, a0, a1, a2 in S.winters:
            rows.append([f"{yy - 1}~{str(yy)[2:]}", f"{a0:.1f}", f"{a1:.1f}" if a1 is not None else "-", f"{a2:.1f}", f"{a2 - a0:+.1f}"])
        y2 = tbl(p.s, rows, x2, y2, w2, [1.1, 1, 1, 1, 1], rowh=230000, size=8.5) + 150000
    if S.kofia and S.kofia.get("rows"):
        y2 = p.panel(x2, y2, w2, "1월 공모 회사채 수요예측 (전체)", "(단위: 조원, %)")
        rows = [["1월", "수요예측", "참여금액", "참여율"]] + [[f"{r[0]}년", f"{r[1]}", f"{r[2]}", f"{r[3]}"] for r in S.kofia["rows"]]
        y2 = tbl(p.s, rows, x2, y2, w2, [1, 1, 1, 1], rowh=225000, size=8.5) + 120000
    # 동사 회차 계절 비교
    dd = [d for d in D.deals if d.get("dt")][-4:]
    if dd:
        rows = [["회차", "수요예측", "발행", "경쟁률", "확정 스프레드"]]
        for d in dd:
            sp_ = " / ".join(bp(r.get("_spread")) for r in d["tr"] if r.get("_spread") is not None) or "-"
            rows.append([f"{d['series']}회", d["dt"][2:].replace("-", "."), f"{d['amt']:,.0f}",
                         f"{d['orders'] / d['offer']:.1f}배" if d.get("orders") and d.get("offer") else "-", sp_])
        yy = max(y2, y + ch_h + 120000)
        if yy + 310000 + 215000 * len(rows) <= p.ybot:
            y3 = p.panel(L, yy, CW, f"{D.name} 최근 회차: 시기별 결과", "(단위: 억원, bp)")
            tbl(p.s, rows, L, y3, CW, [0.8, 1, 1, 0.8, 2], rowh=215000, size=8.5)
        else:
            c_gf.height = Emu(p.ybot - y - 50000)
    p.foot("발행은 증권신고서 기준 공모 회사채(같은 등급군, 수요예측일 기준 월). 스프레드는 ECOS 회사채 AA- 3년 - 국고 3년. "
           "협회 집계는 명세(market_stats)에 입력한 언론 보도 인용치.",
           "전자공시시스템, 한국은행 ECOS, 금융투자협회(언론 보도), 신한투자증권.")


def p_supply(dk, D, I, X):
    S = D.supply
    m = D.issue.month
    d0 = D.sch["수요예측"]
    if S.mat_total and S.after_share is not None and S.after_share >= 50:
        hl = f"{D.issue.year}년 {m}월 {D.band}급 만기 {ek(S.mat_total)}, {S.after_share:.0f}%는 {mdd(d0)} 수요예측 주 이후: 차환 공급이 몰리기 전 일정"
    elif S.mat_total and S.after_share is not None:
        hl = f"{D.issue.year}년 {m}월 {D.band}급 만기 {ek(S.mat_total)}, {100 - S.after_share:.0f}%가 {mdd(d0)} 수요예측 주까지 도래: 차환 자금이 풀리는 시점의 수요예측"
    else:
        hl = f"{m}월 공급과 수요예측 시점"
    bl = [(f"{D.issue.year}년 {m}월 {D.band}급 공모채 만기 {S.mat_issuers}개사 {ek(S.mat_total)}",
           ", ".join(f"{mdd(k)} 주 {ek(v['amt'])}" for k, v in S.mat_weeks))] if S.mat_weeks else []
    if S.ly:
        top = max(S.ly.items(), key=lambda kv: kv[1]["amt"])
        tot = sum(v["amt"] for v in S.ly.values())
        bl.append((f"{S.ly_year}년 {m}월에도 공급은 {top[0]}주차에 집중", f"{D.band}급 발행의 {top[1]['amt'] / tot * 100:.0f}%({top[1]['n']}개사)가 {top[0]}주차"))
    if S.pipe:
        bl.append(("같은 시기 경합 물량", ", ".join(f"{nm}({mdd(dd)}, {ek(a)})" for nm, dd, a in S.pipe[:4])))
    p = dk.page("supply", f"{m}월 공급: 동일 등급 만기와 수요예측 시점", hl, bl)
    lw = 5200000
    y = p.panel(L, p.y, lw, f"{D.issue.year}년 {m}월 {D.band}급 만기도래 (만기 주 기준)", "(단위: 억원)")
    rows = [["만기 주", "발행사", "주요 발행사", "금액"]]
    for k, v in S.mat_weeks:
        nm = sorted(v["names"].items(), key=lambda kv: -kv[1])
        rows.append([f"{mdd(k)}주" + (" ◀" if k == S.my_week else ""), f"{len(nm)}社", ", ".join(n for n, _ in nm[:4]) + (f" 외 {len(nm) - 4}곳" if len(nm) > 4 else ""), f"{v['amt']:,.0f}"])
    rows.append(["합계", f"{S.mat_issuers}곳", "", f"{S.mat_total:,.0f}"])
    yl = tbl(p.s, rows, L, y, lw, [0.8, 0.5, 2.6, 0.7], rowh=[230000] + [260000] * (len(rows) - 1), size=8,
             align_cols={1: PP_ALIGN.CENTER, 2: PP_ALIGN.LEFT, 3: PP_ALIGN.CENTER}, bold_cells={(len(rows) - 1, c) for c in range(4)}) + 150000
    x2 = L + lw + 200000
    w2 = R - x2
    y2 = p.panel(x2, p.y, w2, f"{S.ly_year}년 {m}월 {D.band}급 주차별 수요예측", "(단위: 억원)")
    if S.ly:
        tot = sum(v["amt"] for v in S.ly.values())
        rows = [["주차", "딜수", "발행", "비중", "경쟁률*"]]
        for k, v in S.ly.items():
            rows.append([f"{k}주차", f"{v['n']}건", f"{v['amt']:,.0f}", f"{v['amt'] / tot * 100:.0f}%", f"{v['ratio']:.1f}배" if v["ratio"] else "-"])
        rows.append(["합계", f"{sum(v['n'] for v in S.ly.values())}건", f"{tot:,.0f}", "100%", ""])
        tbl(p.s, rows, x2, y2, w2, [1, 0.8, 1, 0.8, 0.9], rowh=230000, size=8.5, bold_cells={(len(rows) - 1, c) for c in range(5)})
    else:
        box(p.s, x2, y2, w2, 400000, [{"text": "작년 같은 달 수요예측 자료 없음", "size": 9, "color": GREY}])
    if S.pipe:
        y3 = p.panel(x2, y2 + 230000 * (len(S.ly) + 2) + 150000 if S.ly else y2 + 450000, w2, "수요예측 전후 2주 같은 등급군 일정", "(단위: 억원)")
        rows = [["발행사", "수요예측", "금액"]] + [[nm, md(dd), f"{a:,.0f}"] for nm, dd, a in S.pipe[:6]]
        tbl(p.s, rows, x2, y3, w2, [1.6, 1, 0.8], rowh=225000, size=8.5)
    if S.after_share is None:
        t0 = f"{md(d0)} 수요예측"
    elif S.after_share >= 50:
        t0 = f"{md(d0)}은 {D.band}급 만기의 {S.after_share:.0f}%가 남은 시점 — 차환 발행이 몰리는 주보다 앞선 일정"
    else:
        t0 = f"{md(d0)}까지 {D.band}급 만기의 {100 - S.after_share:.0f}%가 도래 — 상환 자금의 재투자 수요를 받는 일정"
    tips = [("수요예측 시점", t0)]
    if D.mpc_near:
        tips.append(("금통위", f"{md(D.mpc_near)} — 수요예측 {'전에 결과 확인' if D.mpc_near < d0 else '뒤 발표, 금리는 그 전에 확정'}"))
    tips.append(("만기는 차환 수요의 상한", "연말 선발행 사례가 있어 실제 발행은 만기보다 적을 수 있음"))
    iy, _ = p.insight("수요예측 시점 검토", tips)
    if len(S.mat_weeks) >= 2 and iy - yl > 1200000:
        y4 = p.panel(L, yl, lw, "주별 만기 (억원)")
        chart(p.s, "col", L, y4, lw, iy - y4 - 120000, [f"{mdd(k)}주" for k, _ in S.mat_weeks], [("만기", [round(v["amt"]) for _, v in S.mat_weeks])], legend=False)
    p.foot("만기는 증권신고서·채권 기본정보의 공모 회사채(같은 등급군). 주차는 월 1일부터 7일 단위. * 경쟁률은 수요예측 결과가 있는 트랜치 평균.",
           "전자공시시스템, 한국예탁결제원, 신한투자증권.")


def p_sched(dk, D, I, X):
    s = D.sch
    tgt = D.targets[0] if D.targets else None
    hl = f"{mdd(s['계약'])} 계약, {mdd(s['신고'])} 신고, {mdd(s['수요예측'])} 수요예측, {mdd(D.issue)} 납입" + \
         (f": {mdd(tgt['mat'])} 만기 {(date.fromisoformat(tgt['mat']) - D.issue).days}일 전 조달" if tgt else "")
    bl = [(f"{md(s['계약'])} 계약부터 {md(D.issue)} 납입까지", f"신고 후 효력발생 7영업일, 수요예측은 신고 2영업일 뒤")]
    evs = [(d, nm) for d, nm in sorted(D.events.items()) if s["계약"] <= d <= D.issue + timedelta(days=14)]
    if evs:
        bl.append(("대외 이벤트", ", ".join(f"{nm} {md(d)}" for d, nm in evs[:4])))
    p = dk.page("sched", f"발행 윈도우 예시: {mdd(D.issue)} 납입", hl, bl)
    y = p.panel(L, p.y, CW, f"발행 일정표 (대표주관계약 {mdd(s['계약'])} ~ 납입 {mdd(D.issue)})", "(단위: 월/일)")
    marks = defaultdict(list)
    for k, lab in (("계약", "대표주관계약 체결"), ("협회", "협회 신고기한"), ("신고", "증권신고서 제출"), ("수요예측", "수요예측"),
                   ("정정", "정정신고서(조건 확정)"), ("발행", "발행 및 납입")):
        if s.get(k):
            marks[s[k]].append(lab)
    for d, nm in D.events.items():
        marks[d].append(nm)
    for b in D.bonds:
        d = date.fromisoformat(b["mat"])
        marks[d].append(f"{b['name'].replace(D.name, '')} 만기")
    start = s["계약"] - timedelta(days=s["계약"].weekday())
    end = max(D.issue, date.fromisoformat(tgt["mat"]) if tgt and (date.fromisoformat(tgt["mat"]) - D.issue).days < 21 else D.issue)
    rows = [["구분"] + [f"{d}요일" for d in "월화수목금"]]
    fills, colors = {}, {}
    wk = start
    while wk <= end and len(rows) < 13:
        dr, er = ["일자"], ["주요 일정"]
        for i in range(5):
            d = wk + timedelta(days=i)
            dr.append(f"{d.month}/{d.day}")
            hol = D.holi.get(d)
            er.append("\n".join(([hol] if hol else []) + marks.get(d, [])))
            r = len(rows)
            if hol:
                colors[(r, i + 1)] = "C00000"
                fills[(r + 1, i + 1)] = "FBE9E9"
            elif d in (s["수요예측"], D.issue):
                fills[(r + 1, i + 1)] = "DCEAF5"
            elif any("금통위" in x or "FOMC" in x for x in marks.get(d, [])):
                fills[(r + 1, i + 1)] = "FFF4D6"
        rows += [dr, er]
        wk += timedelta(days=7)
    n = (len(rows) - 1) // 2
    avail = p.ybot - 650000 - y
    rh = [230000] + [200000, max(300000, (avail - 230000 - 200000 * n) // n)] * n
    for i in range(1, len(rows), 2):
        fills.update({(i, c): "F2F2F2" for c in range(6)})
    tbl(p.s, rows, L, y, CW, [0.6, 1, 1, 1, 1, 1], rowh=rh, size=8, fills=fills, colors=colors)
    p.foot("금융투자협회 규정상 대표주관계약은 증권신고서 제출 10영업일 전 체결, 체결 후 5영업일 이내 협회 신고. 채무증권 신고서 효력발생기간은 7일(공휴일·토요일 불산입). "
           "일정은 당사 산출 예시이며 발행사 내부 결의 일정에 따라 변경 가능.",
           "한국은행, 미 연준, 전자공시시스템, 신한투자증권.")


def p_windows(dk, D, I, X):
    W = D.windows
    if len(W) < 2:
        return
    ch = next(w for w in W if w["chosen"])
    alt = [w for w in W if not w["chosen"]]
    tgt = D.targets[0] if D.targets else None
    lab = lambda w: f"{w['s']['발행'].month}월" + (" (제안)" if w["chosen"] else "")  # noqa: E731
    bes = [w["be"] for w in W if w["be"]]
    least = min(W, key=lambda w: w["ly"][1])
    hl = (f"{lab(ch).split()[0]} 납입은 " + (f"{ch['days']}일 선조달(비용 {ch['cost']:.1f}억원)" if ch["cost"] else "선조달 비용 없음")
          + (f", 작년 같은 달 {D.band}급 발행이 가장 적었던 달" if least is ch else "")
          + (f": 손익분기 금리 차 {min(bes):.1f}~{max(bes):.1f}bp" if bes else ""))
    bl = []
    for w in W:
        s = w["s"]
        t = f"{md(s['수요예측'])} 수요예측, {md(s['발행'])} 납입"
        if w["days"] is not None:
            t += f", {tgt['name']} 만기 {w['days']}일 전" if w["days"] > 0 else f", 만기 {abs(w['days'])}일 뒤(단기 차입으로 연결)"
        if w["mpc"]:
            t += f". 금통위 {md(w['mpc'][0])} {'뒤' if w['mpc'][0] < s['수요예측'] else '전'}"
        n, a, rr = w["ly"]
        if n:
            t += f". 작년 같은 달 {D.band}급 {n}건 {ek(a)}" + (f"(평균 {rr:.1f}배)" if rr else "")
        bl.append((lab(w), t))
    p = dk.page("windows", "발행 윈도우 비교", hl, bl)
    lw = 5900000
    y = p.panel(L, p.y, lw, "발행 윈도우 비교 (예시)", "(단위: 월/일, 억원)")
    rows = [["구분"] + [lab(w) for w in W],
            ["증권신고서"] + [md(w["s"]["신고"]) for w in W],
            ["수요예측"] + [md(w["s"]["수요예측"]) for w in W],
            ["발행·납입"] + [md(w["s"]["발행"]) for w in W],
            ["만기 대비"] + [(f"{w['days']}일 전" if w["days"] > 0 else f"{-w['days']}일 뒤") if w["days"] is not None else "-" for w in W],
            ["금통위·FOMC"] + ["\n".join([f"금통위 {mdd(x)}" for x in w["mpc"]] + [f"FOMC {mdd(x)}" for x in w["fomc"]]) or "-" for w in W],
            [f"같은 달 {D.band}급 만기"] + [f"{w['wall']:,.0f}" for w in W],
            ["작년 같은 달"] + [f"{w['ly'][0]}건 / {w['ly'][1]:,.0f}" for w in W]]
    fills = {(r, 1 + W.index(ch)): "EAF2F8" for r in range(1, len(rows))}
    tbl(p.s, rows, L, y, lw, [1.2] + [1] * len(W), rowh=[230000] + [250000] * 4 + [380000] + [250000] * 2, size=8, fills=fills)
    x2 = L + lw + 200000
    w2 = R - x2
    y2 = p.panel(x2, p.y, w2, "선조달 비용 시산", "(단위: 억원, %, bp)")
    rows = [["항목", "내용"],
            ["조달액", f"{D.base_amt:,}억원 ({' / '.join(f'{t}년 {D.split[t]:,}' for t in D.tenors)})"],
            ["조달 금리", f"{D.wavg:.3f}% (예상 가중평균)" if D.wavg else "-"],
            ["운용 금리 가정", f"{D.base_rate:.2f}% (기준금리)" if D.base_rate is not None else "-"]]
    for w in W:
        if w["cost"]:
            rows.append([f"{lab(w)} 납입 ({w['days']}일)", f"{w['cost']:.1f}억원 / 손익분기 {w['be']:.1f}bp"])
    tbl(p.s, rows, x2, y2, w2, [1.2, 2], rowh=240000, size=8, align_cols={1: PP_ALIGN.LEFT})
    p.insight("판단 기준", [
        ("선조달 손익분기", "먼저 발행해 생기는 이자 비용(조달 금리 - 운용 금리)을 발행 금리 차이(bp)로 환산 — 늦은 윈도우에서 금리가 이 이상 오르면 먼저 발행이 유리"),
        ("수요", f"{lab(ch).split()[0]}은 연초 기관 자금 집행기로 물량 소화력이 큼 (p.{{p:season}})" if ch["s"]["발행"].month == 1 else "같은 달 만기·발행 물량 비교 (p.{p:supply})"),
    ])
    p.foot("만기 대비는 차환 대상 첫 만기 기준. 선조달 비용 = 조달액 × (조달 금리 - 운용 금리) × 일수/365, 손익분기 = 비용 ÷ (조달액 × 가중평균 만기). "
           "운용 금리는 기준금리로 가정(세전, 인수수수료 제외).", "한국은행, 전자공시시스템, 신한투자증권.")


def p_results(dk, D, I, X):
    F_ = D.fin
    ann, cur, prev = F_.ann, F_.cur, F_.prev
    if not ann:
        return
    a1 = ann[-1]
    a0 = ann[-2] if len(ann) > 1 else None
    hl = []
    if cur and prev:
        hl.append(f"{cur['period'][:4]} 누적 매출 {won(cur['revenue'])}({sgn(chg(prev['revenue'], cur['revenue']))}), 영업이익 {won(cur['op'])}({sgn(chg(prev['op'], cur['op']))})")
    p_hl = ": ".join(hl) if hl else f"{a1['period']} 매출 {won(a1['revenue'])}"
    q = {1: "1분기", 2: "상반기", 3: "1~3분기"}.get(F_.q, "")
    bl = []
    if a0:
        bl.append((f"{a1['period']} 매출 {won(a1['revenue'])}({sgn(chg(a0['revenue'], a1['revenue']))}), 영업이익 {won(a1['op'])}({sgn(chg(a0['op'], a1['op']))})",
                   f"영업이익률 {a1.get('op_margin'):.1f}%(전년 {a0.get('op_margin'):.1f}%)" + (f", EBITDA {won(F_.ebitda_a)}" if F_.ebitda_a else "")))
    if cur and prev:
        om_c = cur["op"] / cur["revenue"] * 100 if cur.get("revenue") else None
        om_p = prev["op"] / prev["revenue"] * 100 if prev.get("revenue") else None
        bl.append((f"{q} 매출 {won(cur['revenue'])}({sgn(chg(prev['revenue'], cur['revenue']))}), 영업이익 {won(cur['op'])}({sgn(chg(prev['op'], cur['op']))})",
                   f"영업이익률 {pc(om_c)}(전년 동기 {pc(om_p)})" + (f", EBITDA {won(F_.ebitda_i)}" if F_.ebitda_i else "")
                   + (f", 순이익 {won(cur['ni'])}" if cur.get("ni") is not None else "")))
    if F_.ocf_i and F_.capex_i:
        bl.append(("현금흐름", f"{q} 영업현금흐름 {won(F_.ocf_i)}, 유형자산 투자 {won(F_.capex_i)} — "
                   + ("투자가 영업현금흐름을 상회" if F_.capex_i > F_.ocf_i else "영업현금흐름 안에서 투자")))
    p = dk.page("results", "실적과 현금흐름", p_hl, bl)
    lw = 5600000
    y = p.panel(L, p.y, lw, "실적 요약 (연결)", "(단위: 억원, %)")
    cols = [s["period"] for s in ann] + ([f"{prev['period'][:4]} {q}", f"{cur['period'][:4]} {q}", "증감"] if cur and prev else [])
    def row(lab, key, fmt="n"):
        vs = []
        for s in ann + ([prev, cur] if cur and prev else []):
            v = s.get(key)
            if key == "op_margin" and v is None and s.get("revenue"):
                v = s["op"] / s["revenue"] * 100
            vs.append("-" if v is None else (f"{v / 1e8:,.0f}" if fmt == "n" else f"{v:.1f}%"))
        if cur and prev:
            a, b = prev.get(key), cur.get(key)
            if key == "op_margin":
                a = prev["op"] / prev["revenue"] * 100
                b = cur["op"] / cur["revenue"] * 100
                vs.append(f"{b - a:+.1f}%p")
            else:
                vs.append(sgn(chg(a, b)) if a and b is not None and a > 0 else ("흑자 전환" if a is not None and a < 0 < (b or 0) else "-"))
        return [lab] + vs
    rows = [["구분"] + cols, row("매출", "revenue"), row("영업이익", "op"), row("영업이익률", "op_margin", "p"), row("순이익", "ni")]
    tbl(p.s, rows, L, y, lw, [1] + [1] * (len(cols)), rowh=[240000] + [250000] * 4, size=8.5)
    # 오른쪽: 사업 구성
    ai = (D_co := (X.co or {})).get("ai") or {}
    rev = (ai.get("flow") or {}).get("revenue") or []
    x2 = L + lw + 200000
    w2 = R - x2
    if rev:
        y2 = p.panel(x2, p.y, w2, f"매출 구성 ({(ai.get('flow') or {}).get('basis') or ''})", "(단위: %)")
        chart(p.s, "bar", x2, y2, w2, 1900000, [r["name"] for r in rev][::-1], [("비중", [r["share"] for r in rev][::-1])],
              fmt='0.0"%"', legend=False)
    # 연간 추이 차트
    y3 = y + 240000 + 250000 * 4 + 150000
    y3 = p.panel(L, y3, lw, "연간 매출·영업이익률", "(단위: 억원, %)")
    chart(p.s, "col", L, y3, lw, 1500000, [s["period"] for s in ann], [("매출", [round(s["revenue"] / 1e8) for s in ann])], legend=False)
    pts = []
    if a0 and a1.get("op_margin") is not None and a0.get("op_margin") is not None and a1["op_margin"] < a0["op_margin"]:
        pts.append(("수익성", f"매출 증가에도 영업이익률이 {a0['op_margin']:.1f}% → {a1['op_margin']:.1f}%로 하락 — IR 질의 대상"))
    if cur and prev and cur.get("op") is not None and prev.get("op") and cur["op"] < prev["op"]:
        pts.append(("올해 이익", f"{q} 영업이익 {sgn(chg(prev['op'], cur['op']))} — 4분기 전망과 원가 요인 설명 필요"))
    one = ai.get("caption") or ai.get("one_liner")
    if one:
        pts.append(("사업", one[:140]))
    if pts:
        p.insight("확인 포인트", pts, x=x2, w=w2)
    p.foot(f"연결 기준. {q}는 분기보고서 누적, EBITDA는 영업이익에 감가상각비·무형자산상각비를 더한 값(현금흐름표 기준).",
           "전자공시시스템(사업·분기보고서), 신한투자증권.")


def p_headroom(dk, D, I, X):
    F_ = D.fin
    la, li = F_.lev_a, F_.lev_i
    trig = [t for t in D.trig if t.get("thr") is not None and t.get("cur") is not None]
    met_down = [t for t in trig if t["dir"] == "down" and t.get("met")]
    g_ = lambda t: abs(t["gap"]) if t.get("gap") is not None else abs(t["thr"] - t["cur"])  # noqa: E731
    near_up = [t for t in trig if t["dir"] == "up"]
    if met_down:
        t0 = met_down[0]
        hl = f"{t0['agency']} 하향 검토 기준({t0['metric']} {t0['thr']:g}{t0['unit']} 지속)을 현재 {t0['cur']:g}{t0['unit']}로 상회: 차입 감축 경로가 IR의 핵심"
    elif near_up:
        t = near_up[0]
        hl = f"{t['agency']} 상향 기준 {t['metric']} {t['op']}{t['thr']:g}{t['unit']}까지 {g_(t):.1f}{t['unit']}: 하향 기준까지는 여유"
    else:
        hl = "재무 레버리지와 등급 Headroom"
    bl = [(D.rating_full, ", ".join(f"{a['agency']} {a.get('eval_date') or ''}" for a in D.agencies))] if D.agencies else []
    if li.get("dep") is not None:
        bl.append((f"{F_.pi} 차입금의존도 {li['dep']:.1f}%, 부채비율 {pc(li.get('dr'), 0)}",
                   f"순차입금 {won(li['nd'])}" + (f", 순차입금/EBITDA {li['nd_ebitda']:.1f}배(LTM 근사)" if li.get("nd_ebitda") else "")
                   + (f". {F_.ya}년 말 {la['dep']:.1f}%·{la['nd_ebitda']:.1f}배" if la.get("dep") and la.get("nd_ebitda") else "")))
    for t in trig[:3]:
        st = "기준 도달" if t.get("met") else f"기준까지 {g_(t):.1f}{t['unit']}"
        bl.append((f"{t['agency']} {'상향' if t['dir'] == 'up' else '하향'}", f"{t['text']} — 현재 {t['cur']:g}{t['unit']}, {st}"))
    p = dk.page("credit", "재무 레버리지와 등급 Headroom", hl, bl)
    y = p.panel(L, p.y, CW, "평가사 등급변동요인 대비 지표", "(단위: %, 배)")
    rows = [["지표 (연결)", "등급변동요인", F_.ya, F_.pi, "판정"]]
    keymap = {"차입금의존도": "dep", "순차입금의존도": "ndep", "순차입금/EBITDA": "nd_ebitda", "부채비율": "dr", "순차입금/자본": "nd_eq",
              "EBITDA/이자비용": "icr"}
    for t in trig:
        k = keymap.get(t["metric"])
        fa = la.get(k) if k else None
        fi = li.get(k) if k else None
        unit = t["unit"]
        f = (lambda v: "-" if v is None else (f"{v:.1f}{unit}"))  # noqa: E731
        judge = ("하향 기준 도달" if t["dir"] == "down" else "상향 기준 충족") if t.get("met") else \
                f"{'하향' if t['dir'] == 'down' else '상향'}까지 {g_(t):.1f}{unit}"
        rows.append([f"{t['agency']}: {t['metric']}", f"{'상향' if t['dir'] == 'up' else '하향'} {t['op']} {t['thr']:g}{unit}", f(fa), f(fi) if fi is not None else f"{t.get('cur'):g}{unit}", judge])
    for lab, k, u in (("차입금의존도", "dep", "%"), ("순차입금/EBITDA", "nd_ebitda", "배"), ("부채비율", "dr", "%")):
        if not any(keymap.get(t["metric"]) == k for t in trig):
            rows.append([lab, "-", f"{la[k]:.1f}{u}" if la.get(k) is not None else "-", f"{li[k]:.1f}{u}" if li.get(k) is not None else "-", "참고"])
    fills = {(i, 4): ("FBE9E9" if "하향 기준 도달" in r[4] else "EAF2F8") for i, r in enumerate(rows) if i and r[4] not in ("참고", "-")}
    y = tbl(p.s, rows, L, y, CW, [2, 2, 1, 1, 1.5], rowh=[240000] + [250000] * (len(rows) - 1), size=8.5, fills=fills) + 160000
    ql = defaultdict(list)
    for t in D.trig:
        if t.get("thr") is None:
            ql[(t["agency"], t["dir"])].append(t["text"].rstrip(","))
    items = [(f"{a} {'상향' if d == 'up' else '하향'}", " ".join(v)) for (a, d), v in sorted(ql.items())][:5]
    dh = [(k, v) for k, v in (F_.debt_hist or {}).items() if v]
    if dh:
        cw2 = 3800000
        yc = p.panel(L, y, cw2, "총차입금 추이 (리스 포함)", "(단위: 억원)")
        chart(p.s, "col", L, yc, cw2, p.ybot - yc - 50000, [k for k, _ in dh], [("총차입금", [round(v / 1e8) for _, v in dh])], legend=False)
        if items:
            p.insight("등급변동요인 (정성)", items, y=y, x=L + cw2 + 200000, w=CW - cw2 - 200000)
    elif items:
        p.insight("등급변동요인 (정성)", items, y=y)
    p.foot(f"{F_.ya}은 사업보고서, {F_.pi}는 분기보고서 기준 당사 산출(차입금은 리스부채 포함, 순차입금은 현금·단기금융상품 차감). "
           "순차입금/EBITDA의 분기 EBITDA는 연간 + 당기 누적 - 전년 동기 누적(감가상각은 연환산) 근사. 판정 현재값은 평가사 산식과 다를 수 있음.",
           "신용평가 3사 평가보고서, 전자공시시스템, 신한투자증권.")


def p_prev(dk, D, I, X):
    dd = [d for d in D.deals if d.get("orders")][-3:]
    if not dd:
        return
    last = dd[-1]
    sp_ = [r.get("_spread") for r in last["tr"] if r.get("_spread") is not None]
    hl = f"직전 {last['series']}회는 모집 {ek(last['offer'])}에 {ek(last['orders'])} 주문({last['orders'] / last['offer']:.1f}배)" + \
         (f", 확정 스프레드 {' / '.join(bp(x) for x in sp_)}" if sp_ else "")
    bl = []
    for d in dd:
        tr = d["tr"]
        bl.append((f"{d['series']}회 ({d['dt'][2:].replace('-', '.')})",
                   " / ".join(f"{tn(r['_tenor'])} 모집 {r['_dem']['total'] / r['_dem']['ratio']:,.0f} → 주문 {r['_dem']['total']:,.0f}({r['_dem']['ratio']:.1f}배) → 발행 {r['amount'] / 1e8:,.0f}"
                              + (f", {bp(r['_spread'])}" if r.get("_spread") is not None else "") for r in tr if r.get("_dem") and r.get("_tenor"))))
    p = dk.page("prev", "직전 회차 수요예측 복기", hl, bl)
    lw = 5600000
    y = p.panel(L, p.y, lw, f"{D.name} 최근 회차", "(단위: 억원, 배, bp, %)")
    rows = [["회차", "수요예측", "만기", "모집", "주문", "경쟁률", "발행", "확정", "쿠폰"]]
    for d in dd:
        for r in d["tr"]:
            dm = r.get("_dem") or {}
            rows.append([r["series_raw"], (d["dt"] or "")[2:].replace("-", "."), tn(r["_tenor"]) if r.get("_tenor") else "-",
                         f"{dm['total'] / dm['ratio']:,.0f}" if dm.get("ratio") else "-", f"{dm.get('total', 0):,.0f}",
                         f"{dm['ratio']:.2f}" if dm.get("ratio") else "-", f"{r['amount'] / 1e8:,.0f}",
                         bp(r.get("_spread")), f"{r['_rate']:.3f}" if r.get("_rate") else "-"])
    yy = tbl(p.s, rows, L, y, lw, [0.8, 0.9, 0.6, 0.8, 0.8, 0.7, 0.8, 0.7, 0.8], rowh=[230000] + [225000] * (len(rows) - 1), size=8) + 180000
    hist = [d for d in D.deals if any(r.get("_spread") is not None for r in d["tr"])][-6:]
    if hist:
        yy = p.panel(L, yy, lw, "회차별 확정 스프레드 추이", "(단위: 억원, bp)")
        rows = [["회차", "수요예측", "만기 구성", "발행", "확정 스프레드"]]
        for d in hist:
            rows.append([f"{d['series']}회", (d["dt"] or "")[2:].replace("-", "."), " / ".join(tn(r["_tenor"]) for r in d["tr"] if r.get("_tenor")),
                         f"{d['amt']:,.0f}", " / ".join(bp(r.get("_spread")) for r in d["tr"] if r.get("_spread") is not None)])
        tbl(p.s, rows, L, yy, lw, [0.7, 0.9, 1.3, 0.8, 1.6], rowh=225000, size=8)
    x2 = L + lw + 200000
    w2 = R - x2
    y2 = p.panel(x2, p.y, w2, f"{last['series']}회 기관군별 주문", "(단위: 억원)")
    groups = [("운용사", ["운용사"]), ("증권사", ["증권사"]), ("연기금·은행·보험", ["연기금·은행·보험"]), ("기타·외국", ["기타", "외국(실적유)", "외국(실적무)"])]
    tl = [tn(r["_tenor"]) for r in last["tr"] if r.get("_tenor")]
    ser = [(g, [round(sum((r["_dem"]["groups"] or {}).get(k, 0) for k in ks)) for r in last["tr"] if r.get("_tenor")]) for g, ks in groups]
    ser = [s for s in ser if sum(s[1])]
    chart(p.s, "colstack", x2, y2, w2, 2300000, tl, ser)
    tot = sum(sum(v) for _, v in ser) or 1
    top = max(ser, key=lambda s: sum(s[1]))
    pts = [("주요 수요", f"{top[0]}가 주문의 {sum(top[1]) / tot * 100:.0f}% — 같은 투자자군을 우선 접촉 (p.{{p:investors}})")]
    if len(dd) > 1 and dd[-2].get("offer"):
        a, b = dd[-2], dd[-1]
        pts.append(("회차 간 비교", f"{a['series']}회 {a['orders'] / a['offer']:.1f}배 → {b['series']}회 {b['orders'] / b['offer']:.1f}배"))
    sec = [r for r in last["tr"] if r.get("_dem")]
    ss = sum((r["_dem"]["groups"] or {}).get("증권사", 0) for r in sec)
    if ss / tot > 0.3:
        pts.append(("증권사 비중", f"주문의 {ss / tot * 100:.0f}%가 증권사 — 실수요 기관 확대가 결정 수준을 낮추는 관건"))
    p.insight("시사점", pts, x=x2, w=w2)
    p.foot("모집은 주문 ÷ 경쟁률, 확정 스프레드는 개별민평 대비. 수요예측 결과는 발행조건 확정 정정신고서 기준.",
           "전자공시시스템(증권신고서·정정신고서), 신한투자증권.")


def p_investors(dk, D, I, X):
    dem = D.dem
    if not dem:
        return
    groups = ["운용사", "증권사", "연기금·은행·보험 등", "기타·외국"]
    tot = {k: sum(v["groups"].values()) or 1 for k, v in dem.items()}
    keys = [k for k in ("2년", "3년", "5년 이상") if k in dem]
    hl = f"최근 6개월 {D.band}급 수요예측 주문 {ek(sum(v['orders'] for v in dem.values()))}: " + \
         ", ".join(f"{k} 평균 {dem[k]['avg_ratio']:.1f}배" for k in keys if dem[k].get("avg_ratio"))
    bl = []
    for k in keys:
        v = dem[k]
        g = sorted(v["groups"].items(), key=lambda kv: -kv[1])
        bl.append((f"{k} {v['n']}건(딜 {v['deals']}개)", f"주문 {ek(v['orders'])}, 평균 경쟁률 {v['avg_ratio']:.1f}배, 확정 스프레드 중앙값 {bp(v['sp_med'])}({bp(v['sp_lo'])} ~ {bp(v['sp_hi'])})"
                   + (f". 주문의 {g[0][1] / tot[k] * 100:.0f}%가 {g[0][0]}" if g else "")))
    p = dk.page("investors", "투자자 구성과 접촉 대상", hl, bl)
    lw = 5300000
    y = p.panel(L, p.y, lw, f"{D.band}급 만기별 기관군 주문 비중 (최근 6개월)", "(단위: %)")
    ser = [(g, [round(dem[k]["groups"].get(g, 0) / tot[k] * 100, 1) for k in keys]) for g in groups]
    ser = [s for s in ser if sum(s[1])]
    chart(p.s, "colstack", L, y, lw, 2500000, keys, ser, fmt='0"%"')
    x2 = L + lw + 200000
    w2 = R - x2
    y2 = p.panel(x2, p.y, w2, "같은 등급 최근 사례", "(단위: 억원, bp)")
    rows = [["발행사", "수요예측", "만기", "경쟁률", "확정"]]
    for r in X.comps[:8]:
        rows.append([r["corp_name"], (r.get("_demand_dt") or "")[5:].replace("-", "/"), tn(r["_tenor"]) if r.get("_tenor") else "-",
                     f"{r['_dem']['ratio']:.1f}배" if (r.get("_dem") or {}).get("ratio") else "-", bp(r.get("_spread"))])
    tbl(p.s, rows, x2, y2, w2, [1.5, 0.8, 0.6, 0.8, 0.7], rowh=225000, size=8)
    p.insight("접촉 대상과 방식", [
        ("실수요 기관", "연기금·은행·보험은 확정 스프레드 이하 주문 비중이 높은 기관군 — 신고 전 NDR로 질의를 먼저 받음"),
        ("운용사", "만기별 주문 규모가 가장 큰 기관군 — 밴드 하단 주문 유도가 결정 수준을 좌우"),
        ("증권사", "인수 목적 주문 비중이 높아 최종 배정은 실수요 기관 중심으로 설계"),
    ])
    p.foot("같은 등급군(±등급 구간) 무보증 공모사채 수요예측 결과. 기관별 참여 내역은 포함하지 않음(기관군 합계).",
           "전자공시시스템(정정신고서 수요예측 결과), 신한투자증권.")


def p_ir(dk, D, I, X, F):
    pts, qa, msg = credit_story(F, X, I)
    hl = f"IR 메시지: {msg}" if msg else "IR 포인트와 투자자 Q&A"
    p = dk.page("ir", "IR 포인트와 투자자 Q&A", hl, [(a, b) for a, b in pts])
    y = p.panel(L, p.y, CW, "예상 질문과 답변 방향", None)
    rows = [["예상 질문", "답변 방향", "근거"]]
    refs = ["p.{p:credit}", "p.{p:results}", "p.{p:credit}", "p.{p:results}", "p.{p:credit}"]
    for i, (q, a) in enumerate(qa):
        rows.append([q, a, refs[i % len(refs)]])
    tbl(p.s, rows, L, y, CW, [2.2, 4.6, 0.6], rowh=[240000] + [330000] * (len(rows) - 1), size=8.5,
        align_cols={1: PP_ALIGN.LEFT, 2: PP_ALIGN.CENTER})
    p.foot("답변 방향은 공시 수치 기준 초안 — '##'은 회사 확인이 필요한 항목.", "전자공시시스템, 신용평가 3사, 신한투자증권.")


def p_policy(dk, D, I, X, F, spec):
    rt = F["rates"]
    kr, _ = kr_lines(rt, D.cv)
    us, _ = us_lines(rt)
    mb = (spec.get("members") or {})
    nxt = [d for d in D.mpc if d >= D.today]
    nf = [d for d in D.fomc if d >= D.today]
    hl = f"기준금리 {D.base_rate:.2f}%" + (f", 다음 금통위 {md(nxt[0])}" if nxt else "") + \
         (f" · FOMC {md(nf[0])}" if nf else "") + f": {md(D.sch['수요예측'])} 수요예측은 " + \
         (f"{md(D.mpc_near)} 금통위 {'다음' if D.mpc_near < D.sch['수요예측'] else '전'}" if D.mpc_near else "")
    p = dk.page("policy", "통화정책: 금통위와 FOMC", hl)
    lw = (CW - 200000) // 2
    y = p.panel(L, p.y, lw, "韓 금통위", None)
    ps = [{"text": x.lstrip("- "), "size": 9, "bul": "·" if not x.startswith("-") else "-", "after": 2} for x in kr]
    if mb.get("mpc", {}).get("summary"):
        ps += [{"text": "위원 성향", "size": 9.5, "bold": True, "color": NAVY, "after": 1}] + \
              [{"text": s, "size": 9, "bul": "·", "after": 2} for s in mb["mpc"]["summary"].split("\n")]
    box(p.s, L, y, lw, 2600000, ps, margin=20000)
    x2 = L + lw + 200000
    y2 = p.panel(x2, p.y, lw, "美 FOMC", None)
    ps = [{"text": x.lstrip("- "), "size": 9, "bul": "·" if not x.startswith("-") else "-", "after": 2} for x in us]
    if mb.get("fed", {}).get("summary"):
        ps += [{"text": "연준 성향", "size": 9.5, "bold": True, "color": NAVY, "after": 1}] + \
              [{"text": s, "size": 9, "bul": "·", "after": 2} for s in mb["fed"]["summary"].split("\n")]
    if D.fedwatch:
        ps += [{"text": f"FedWatch ({D.fedwatch.get('asof')})", "size": 9.5, "bold": True, "color": NAVY, "after": 1}] + \
              [{"text": f"{k} {v}", "size": 9, "bul": "·", "after": 2} for k, v in (D.fedwatch.get("probs") or {}).items()]
    box(p.s, x2, y2, lw, 2600000, ps, margin=20000)
    ys = y + 2700000
    ys = p.panel(L, ys, CW, "기준금리와 국고 3년", "(단위: %)")
    b = rt.get("base") or []
    k3 = rt.get("ktb3") or []
    since = (D.today - timedelta(days=365)).isoformat()
    kk = [(d, v) for d, v in k3 if d >= since][::5]
    bb = dict(b)
    cats, vb, vk = [], [], []
    last_b = None
    for d, v in kk:
        for dd_, x in b:
            if dd_ <= d:
                last_b = x
        cats.append(d[2:7].replace("-", "."))
        vb.append(last_b)
        vk.append(v)
    if cats:
        lo_ = min(v for v in vb + vk if v is not None)
        chart(p.s, "line", L, ys, CW, p.ybot - ys - 100000, cats, [("기준금리", vb), ("국고 3년", vk)], fmt="0.00", ymin=int(lo_ * 2) / 2 - 0.5)
    p.foot("금리는 일별 종가(국고 3년은 5영업일 간격 표시). 위원 성향은 명세(members)에 정리한 공개 발언·의결 기준.",
           "한국은행, 미 연준, 금융투자협회, CME FedWatch(명세 입력), 신한투자증권.")


def p_outlook(dk, D, I, X):
    O = D.outlook
    if not O:
        return
    from tools import proposal_outlook as OL
    cols = O["cols"][:7]
    b_ = OL.series(O, "KR", "base") or []
    k3 = OL.ktb_path(O, 3) or []
    gp, how = OL.grade_path(O, D.grade, 3, D.cv)
    aa = OL.series(O, "CR", "AA-", 3) or []
    now_g = D.cv.y(D.grade, 3) if D.cv else None
    lab_q, v_q = OL.at(O, gp, D.issue) if gp else (None, None)
    hl = f"리서치({str(O['vintage'])[5:].replace('-', '/')}): {D.grade} 3년 현재 {now_g:.2f}% → {lab_q} {v_q:.2f}%" if (now_g and v_q) else "신한 리서치 금리전망"
    bl = []
    if b_:
        bl.append(("기준금리", " → ".join(f"{lab} {v:.2f}%" for (lab, _), v in zip(cols, b_) if v is not None)))
    if gp:
        bl.append((f"{D.grade} 3년", " → ".join(f"{lab} {v:.2f}%" for (lab, _), v in zip(cols, gp) if v is not None) + f" ({how})"))
    com = (O.get("comments") or {})
    if com.get("kr"):
        bl.append(("국내", OL.first_sentences(com["kr"], 160)))
    if com.get("credit"):
        bl.append(("크레딧", OL.first_sentences(com["credit"], 160)))
    p = dk.page("outlook", "신한 리서치 금리전망", hl, bl)
    y = p.panel(L, p.y, CW, "금리 전망 경로", "(단위: %, 기말 기준)")
    rows = [["구분"] + [lab for lab, _ in cols]]
    for lab, s in (("기준금리", b_), ("국고 3년", k3), ("AA- 3년", aa), (f"{D.grade} 3년", gp)):
        if s:
            rows.append([lab] + [f"{v:.2f}" if v is not None else "-" for v in s[:len(cols)]])
    y = tbl(p.s, rows, L, y, CW, [1.2] + [1] * len(cols), rowh=235000, size=8.5) + 140000
    if gp and k3:
        y = p.panel(L, y, CW, f"{D.grade} 3년 스프레드 전망 (등급 금리 - 국고)", "(단위: bp)")
        sp = [round((a - b) * 100) if a is not None and b is not None else None for a, b in zip(gp[:len(cols)], k3[:len(cols)])]
        chart(p.s, "col", L, y, CW, max(900000, p.ybot - y - 100000), [lab for lab, _ in cols], [("스프레드", sp)], legend=False)
    p.foot(f"신한투자증권 리서치센터 금리전망({O['vintage']}, 시트 {O.get('sheet')}). {D.grade} 경로는 {how}.",
           "신한투자증권 리서치센터.")


def p_credit_mkt(dk, D, I, X, F):
    cv = D.cv
    g = D.grade
    if not cv:
        return
    s_now = cv.sp(g, 3)
    rg = cv.range(g, 3, 365)
    lo, hi = (rg[0][1], rg[1][1]) if rg else (None, None)
    rk = cv.pct_rank(g, 3, 365)
    c1 = cv.change(g, 3, 30, spread=True)
    hl = f"{g} 3년 스프레드 {s_now:.0f}bp, 1년 범위 {lo:.0f}~{hi:.0f}bp 중 {rk:.0f}% 위치" if s_now is not None and lo is not None and rk is not None else "크레딧 시장"
    bl = [(f"{g} 3년 {cv.y(g, 3):.3f}%({mdd(D.asof)})", f"국고 3년 대비 {s_now:.0f}bp, 1개월 {bp(c1)}")] if s_now is not None else []
    mv = ((I.get("_mv") or {}).get("credit")) or []
    for x in mv[:2]:
        bl.append(x.lstrip("- "))
    p = dk.page("credit_mkt", "크레딧 시장: 수급과 스프레드", hl, bl)
    lw = 5600000
    y = p.panel(L, p.y, lw, f"{g}·AA- 3년 신용스프레드 (최근 1년)", "(단위: bp)")
    since = (D.today - timedelta(days=365)).isoformat()
    s1 = cv.series(g, 3, since) if hasattr(cv, "series") else []
    pts = [(d, v) for d, v in s1][::3]
    if pts:
        aa = []
        for d, _ in pts:
            try:
                aa.append(round(cv.sp("AA-", 3, d), 1))
            except Exception:  # noqa: BLE001
                aa.append(None)
        sp = []
        for d, v in pts:
            try:
                sp.append(round(cv.sp(g, 3, d), 1))
            except Exception:  # noqa: BLE001
                sp.append(None)
        chart(p.s, "line", L, y, lw, 2500000, [d[2:].replace("-", ".") for d, _ in pts], [(f"{g} 3년", sp), ("AA- 3년", aa)], fmt="0")
        gap_now = (cv.sp(g, 3) or 0) - (cv.sp("AA-", 3) or 0)
        gaps = [a_ - b_ for a_, b_ in zip(sp, aa) if a_ is not None and b_ is not None]
        items = [("위치", f"{g} 3년 스프레드는 1년 범위의 {rk}% 지점 — {'높은 편' if rk >= 70 else '낮은 편' if rk <= 30 else '중간'}"),
                 ("등급 간 격차", f"{g}-AA- 3년 {gap_now:.0f}bp(1년 {min(gaps):.0f}~{max(gaps):.0f}bp)" if gaps else None),
                 ("1개월", f"{g} 3년 스프레드 {bp(c1)}, 국고 3년 {bp(cv.change('KTB', 3, 30))}")]
        p.insight("시사점", items, x=L, w=lw)
    x2 = L + lw + 200000
    w2 = R - x2
    y2 = p.panel(x2, p.y, w2, f"{D.band}급 수요예측 (최근 6개월)", "(단위: 억원, 배, bp)")
    rows = [["만기", "건수", "주문", "경쟁률", "스프레드"]]
    for k in ("2년", "3년", "5년 이상"):
        v = D.dem.get(k)
        if v:
            rows.append([k, f"{v['n']}", f"{v['orders']:,.0f}", f"{v['avg_ratio']:.1f}" if v.get("avg_ratio") else "-", bp(v.get("sp_med"))])
    y2 = tbl(p.s, rows, x2, y2, w2, [0.9, 0.6, 1, 0.8, 0.9], rowh=230000, size=8.5) + 150000
    mat_n = defaultdict(float)
    for m in F.get("mats") or []:
        if m["mat"].startswith(str(D.issue.year)) and M.in_band(m.get("grade"), D.band):
            mat_n[int(m["mat"][5:7])] += m["amt"]
    if mat_n:
        y2 = p.panel(x2, y2, w2, f"{D.issue.year}년 {D.band}급 월별 만기", "(단위: 억원)")
        chart(p.s, "col", x2, y2, w2, max(900000, p.ybot - y2 - 50000), [f"{k}월" for k in range(1, 13)],
              [("만기", [round(mat_n.get(k, 0)) or None for k in range(1, 13)])], legend=False, label_size=7)
    p.foot("신용스프레드 = 등급별 시가평가 수익률(평가사 평균) - 국고. 수요예측은 같은 등급군 공모 회사채.",
           "금융투자협회 채권정보센터, 전자공시시스템, 신한투자증권.")


def p_indiv(dk, D, I, X):
    cv = D.cv
    t = D.tenors
    a = D.adj
    if D.indiv_given:
        hl = f"입력 개별민평 " + " / ".join(f"{x}년 {D.base[x][0]:.3f}%" for x in t)
    elif a:
        hl = f"최근 {a['days']}일 동사 채권은 {D.grade} 등급민평 대비 {bp(a['bp'])}에서 거래: 추정 개별민평 " + " / ".join(f"{x}년 {D.base[x][0]:.3f}%" for x in t if D.base[x][0])
    else:
        hl = "개별민평과 유통시장"
    bl = [(f"{D.grade} 등급민평({mdd(D.asof)})", " / ".join(f"{x}년 {cv.y(D.grade, x):.3f}%" for x in t))]
    if a:
        bl.append(("장외 거래", f"{a['from'][5:].replace('-', '/')}~{a['to'][5:].replace('-', '/')} {a['n']}건, 거래대금 {a['amt']:,}억원, 등급민평 대비 가중 중앙값 {bp(a['bp'])}"))
    if not D.indiv_given:
        bl.append(("확인 필요", "민평 4사 개별민평으로 바꾸면 밴드 기준점과 예상 금리가 자동으로 다시 계산됨"))
    p = dk.page("indiv", f"{D.name} 개별민평과 유통시장", hl, bl)
    lw = 4600000
    y = p.panel(L, p.y, lw, "만기별 기준 금리", "(단위: %, bp)")
    rows = [["만기", "등급민평", "조정", "기준 금리", "국고 대비"]]
    for x in t:
        g = cv.y(D.grade, x)
        k = cv.y("KTB", x)
        rows.append([f"{x}년", f"{g:.3f}", bp((D.base[x][0] - g) * 100) if D.base[x][0] else "-", f"{D.base[x][0]:.3f}" if D.base[x][0] else "-",
                     bp((D.base[x][0] - k) * 100, False) if D.base[x][0] and k else "-"])
    tbl(p.s, rows, L, y, lw, [0.7, 1, 0.8, 1, 0.9], rowh=240000, size=8.5)
    x2 = L + lw + 200000
    w2 = R - x2
    y2 = p.panel(x2, p.y, w2, "동사 채권 장외 거래 (최근)", "(단위: 년, %, 억원, bp)")
    rows = [["일자", "종목", "잔존", "수익률", "거래", "민평 대비"]]
    for r in D.trades[:10]:
        rows.append([r["d"][5:].replace("-", "/"), r["nm"].replace(D.name, ""), f"{r['yrs']:.1f}", f"{r['y']:.3f}", f"{r['amt']:,.0f}", bp(r["vs"])])
    if len(rows) > 1:
        tbl(p.s, rows, x2, y2, w2, [0.7, 0.9, 0.6, 0.8, 0.7, 0.9], rowh=220000, size=8)
    else:
        box(p.s, x2, y2, w2, 300000, [{"text": "최근 장외 거래 없음", "size": 9, "color": GREY}])
    p.insight("시사점", [
        ("밴드 기준점", "밴드는 청약 1영업일 전 개별민평 기준 — 기준점 변화는 수요예측 결과보다 먼저 금리에 반영"),
        ("유통 수급", f"민평보다 {'낮은' if a and a['bp'] < 0 else '높은'} 금리 거래가 많으면 결정 스프레드도 같은 방향으로 형성되는 경향" if a else "거래 표본이 적어 참고 지표"),
    ])
    p.foot("등급민평은 금융투자협회 시가평가 수익률(평가사 평균). 추정 개별민평 = 등급민평 + 최근 30일 동사 장외 체결의 등급민평 대비 가중 중앙값(체결 3건·20억 이상일 때). "
           "민평 대비는 잔존 만기에 맞춘 등급민평과 비교.", "금융투자협회 채권정보센터, 신한투자증권.")


def p_cost(dk, D, I, X):
    t = D.tenors
    if not D.wavg:
        return
    hl = f"예상 가중평균 {D.wavg:.3f}%, 연 이자 {ek(D.int_yr)}: 스프레드 1bp는 연 {D.bp1:.2f}억원, 밴드 폭 {X.band_bp[1] - X.band_bp[0]}bp는 연 {D.bp1 * (X.band_bp[1] - X.band_bp[0]):.1f}억원"
    bl = [("만기별 예상 금리", " / ".join(f"{x}년 {D.exp[x]:.3f}%(기준 {D.base[x][0]:.3f}% {bp(D.tgt[x])})" for x in t))]
    if D.target_cpn:
        bl.append(("차환 대상 쿠폰", f"{D.target_cpn:.3f}% 대비 {(D.wavg - D.target_cpn) * 100:+.0f}bp — 같은 금액 기준 연 이자 {ek(abs(D.wavg - D.target_cpn) * D.target_amt / 100)} {'증가' if D.wavg > D.target_cpn else '감소'}"))
    p = dk.page("cost", "조달비용 시산", hl, bl)
    y = p.panel(L, p.y, CW, "만기별 조달비용", "(단위: 억원, %, bp)")
    rows = [["만기", "금액", "기준 금리", "목표 스프레드", "예상 금리", "연 이자", "-10bp", "+10bp"]]
    for x in t:
        a = D.split[x]
        rows.append([f"{x}년", f"{a:,}", f"{D.base[x][0]:.3f}", bp(D.tgt[x]), f"{D.exp[x]:.3f}", f"{a * D.exp[x] / 100:,.1f}",
                     f"{a * (D.exp[x] - 0.1) / 100:,.1f}", f"{a * (D.exp[x] + 0.1) / 100:,.1f}"])
    rows.append(["합계·가중", f"{D.base_amt:,}", "", "", f"{D.wavg:.3f}", f"{D.int_yr:,.1f}", f"{D.int_yr - D.bp1 * 10:,.1f}", f"{D.int_yr + D.bp1 * 10:,.1f}"])
    y = tbl(p.s, rows, L, y, CW, [1, 1, 1, 1, 1, 1, 1, 1], rowh=245000, size=8.5, bold_cells={(len(rows) - 1, c) for c in range(8)}) + 160000
    if D.outlook:
        from tools import proposal_outlook as OL
        gp, _ = OL.grade_path(D.outlook, D.grade, 3, D.cv)
        if gp:
            y = p.panel(L, y, CW, f"발행 시점별 {D.grade} 3년 금리 (리서치 전망 기준)", "(단위: %)")
            cols = D.outlook["cols"][:6]
            chart(p.s, "col", L, y, CW, max(900000, p.ybot - y - 600000), [lab for lab, _ in cols], [(f"{D.grade} 3년", [round(v, 2) if v else None for v in gp[:6]])],
                  fmt="0.00", legend=False, ymin=min(v for v in gp[:6] if v) - 0.3)
    p.foot("연 이자는 표면금리 기준(인수수수료·발행 비용 제외). 예상 금리 = 기준 금리 + 목표 스프레드(동사 직전 확정값과 같은 등급군 중앙값의 평균, 밴드 안).",
           "금융투자협회 채권정보센터, 신한투자증권 리서치센터, 신한투자증권.")


# ───────────────────────── 표지·목차·간지 ─────────────────────────
def contents(prs, lay, pages):
    s = prs.slides.add_slide(lay)
    box(s, L, 700000, 5000000, 600000, [{"text": "Contents", "size": 28, "bold": True, "color": NAVY}], margin=0)
    rect(s, L, 1350000, 900000, 50000, fill="F2A900")
    ps = []
    for i, (n, t) in enumerate(SECS):
        ps.append({"text": f"{n}.  {t}", "size": 15, "bold": True, "color": NAVY, "after": 4})
        for sl, key, sec, title in pages:
            if sec == i:
                ps.append({"text": title, "size": 10, "color": SUB, "bul": "·", "after": 1})
        ps[-1]["after"] = 12
    box(s, L, 1650000, 8000000, 5200000, ps, margin=0)
    return s


def divider(prs, lay, i, pages):
    s = prs.slides.add_slide(lay)
    rect(s, 0, 0, 3600000, 7562850, fill=DEEP)
    box(s, 400000, 2600000, 3000000, 900000, [{"text": SECS[i][0], "size": 44, "bold": True, "color": "FFFFFF"}], margin=0)
    box(s, 400000, 3500000, 3000000, 900000, [{"text": SECS[i][1], "size": 18, "bold": True, "color": "FFFFFF"}], margin=0)
    ps = [{"text": f"{k + 1}. {t}", "size": 13, "color": NAVY, "after": 8} for k, t in enumerate([t for _, _, sec, t in pages if sec == i])]
    box(s, 4100000, 2300000, 5800000, 4000000, ps, margin=0)
    return s


# ───────────────────────── 원래 제안서(v5) 디자인에 넣기 ─────────────────────────
LAY5 = {"summary": "SUM", "season": "I", "supply": "I", "windows": "I", "mat": "I", "terms": "I", "sched": "I",
        "results": "II", "credit": "II", "prev": "II", "investors": "II", "ir": "II",
        "indiv": "III", "cost": "III", "policy": "III", "outlook": "III", "credit_mkt": "III"}
TITLE5 = {"summary": "권고안 요약", "credit": "재무 레버리지와 등급 Headroom", "prev": "前발행 수요예측 복기"}
MIX = [("summary", p_summary), ("season", p_season), ("supply", p_supply), ("windows", p_windows),
       ("results", p_results), ("credit", p_headroom), ("prev", p_prev), ("investors", p_investors),
       ("indiv", p_indiv), ("cost", p_cost)]


class Deck5(Deck):
    def __init__(self, prs, PR, lays):
        super().__init__(prs, None)
        self.PR, self.lays = PR, lays

    def page(self, key, title, headline=None, bullets=None):
        sl = P5.new_slide(self.prs, self.lays[LAY5.get(key, "I")])
        p = Page5(sl, self.PR, TITLE5.get(key, title), headline, bullets, key)
        self.pages.append((sl, key, self.sec, title))
        return p


def extra_pages(prs, PR, lays, F, I, X, spec):
    """v6 의 내용 장들을 원래 제안서 디자인으로 만든다 → ({key: slide}, D)."""
    D = calc(F, I, X, spec)
    I2 = dict(I, _mv={"credit": P5.fresh_view(spec, "credit", D.today, X.rules)})
    dk = Deck5(prs, PR, lays)
    on = (spec.get("deck_v6") or {}).get("pages") or {}
    for key, fn in MIX:
        if on.get(key, True) is False:
            continue
        try:
            fn(dk, D, I2, X)
        except Exception as e:  # noqa: BLE001
            import traceback
            print(f"page {key} skipped: {type(e).__name__}: {e}")
            traceback.print_exc()
            D.check.append(f"{key} 장 생성 실패 — 데이터 확인 필요")
    return {key: sl for sl, key, _, _ in dk.pages}, D


def finish_checks(prs, D, first):
    for i, s_ in enumerate(prs.slides):
        if any(sh.has_text_frame and "##" in sh.text_frame.text for sh in s_.shapes) or \
                any(sh.has_table and any("##" in c.text for r in sh.table.rows for c in r.cells) for sh in s_.shapes):
            D.check.append(f"p.{i + 1}: '##' 표시 항목은 회사 확인 후 채움")
    if first is not None:
        nt = first.notes_slide.notes_text_frame
        nt.text = "[확인 목록]\n" + "\n".join(f"- {c}" for c in D.check) + (("\n\n" + nt.text) if nt.text else "")
    return D.check


def build(template, F, I, X, spec, out_path):
    D = calc(F, I, X, spec)
    prs = Presentation(template)
    S = list(prs.slides)
    cover, disc, end = S[0], S[1], S[-1]
    lay = next(l for l in prs.slide_layouts if l.name == "1_빈 화면")
    lay_d = next(l for l in prs.slide_layouts if l.name == "빈 화면")
    name = F["name"]
    import re as _re
    for sh in list(cover.shapes):
        if sh.has_text_frame and "위한" in sh.text_frame.text:
            sh._element.getparent().remove(sh._element)
        elif sh.has_text_frame:
            for para_ in sh.text_frame.paragraphs:
                for r_ in para_.runs:
                    r_.text = _re.sub(r"20\d\d\.\d\d", f"{D.today.year}.{D.today.month:02d}", r_.text)
    box(cover, 2600000, 3350000, 7700000, 1900000, [
        {"text": f"{name}", "size": 30, "bold": True, "color": "FFFFFF", "align": PP_ALIGN.RIGHT, "after": 6},
        {"text": f"{D.issue.year}년 회사채 발행 검토 자료", "size": 26, "bold": True, "color": "FFFFFF", "align": PP_ALIGN.RIGHT, "after": 4},
        {"text": "DCM Discussion Materials", "size": 14, "color": "C9D6E2", "align": PP_ALIGN.RIGHT}], margin=0,
        anchor=MSO_ANCHOR.BOTTOM)
    for sh in disc.shapes:
        if sh.has_text_frame:
            for para_ in sh.text_frame.paragraphs:
                for r_ in para_.runs:
                    r_.text = _re.sub(r"COPYRIGHT©\d{4}", f"COPYRIGHT©{D.today.year}", r_.text)
    mv = {"credit": P5.fresh_view(spec, "credit", D.today, X.rules)}
    I = dict(I, _mv=mv)
    dk = Deck(prs, lay)
    on = (spec.get("deck_v6") or {}).get("pages") or {}
    plan = [
        (0, [("summary", p_summary), ("terms", p_terms), ("mat", p_maturity), ("season", p_season), ("supply", p_supply),
             ("sched", p_sched), ("windows", p_windows)]),
        (1, [("results", p_results), ("credit", p_headroom)]),
        (2, [("prev", p_prev), ("investors", p_investors), ("ir", None)]),
        (3, [("policy", None), ("outlook", p_outlook), ("credit_mkt", None), ("indiv", p_indiv), ("cost", p_cost)]),
    ]
    for sec, items in plan:
        dk.section(sec)
        for key, fn in items:
            if on.get(key, True) is False:
                continue
            try:
                if key == "ir":
                    p_ir(dk, D, I, X, F)
                elif key == "policy":
                    p_policy(dk, D, I, X, F, spec)
                elif key == "credit_mkt":
                    p_credit_mkt(dk, D, I, X, F)
                else:
                    fn(dk, D, I, X)
            except Exception as e:  # noqa: BLE001 — 한 장의 문제로 전체가 멈추지 않게
                import traceback
                print(f"page {key} skipped: {type(e).__name__}: {e}")
                traceback.print_exc()
                D.check.append(f"{key} 페이지 생성 실패 — 데이터 확인 필요")
    built = list(dk.pages)
    toc = contents(prs, lay, built)
    divs = [divider(prs, lay_d, i, built) for i in range(len(SECS))]
    # 순서: 표지, 목차, (간지, 페이지...)×4, Disclaimer, 끝
    order = [cover, toc]
    for i in range(len(SECS)):
        order.append(divs[i])
        order += [sl for sl, _, sec, _ in built if sec == i]
    order += [disc, end]
    keep = {id(s) for s in order}
    for s in list(prs.slides):
        if id(s) not in keep:
            P5.remove_slide(prs, s)
    P5.reorder(prs, order)
    nums = {key: list(prs.slides).index(sl) + 1 for sl, key, _, _ in built}
    replace_tokens(prs, nums)
    # 쪽번호
    for i, s in enumerate(prs.slides):
        if any(s is sl for sl, *_ in built):
            box(s, 10693400 // 2 - 300000, 7230000, 600000, 200000, [{"text": str(i + 1), "size": 8, "color": GREY, "align": PP_ALIGN.CENTER}], margin=0)
    for i, s_ in enumerate(prs.slides):
        if any(sh.has_text_frame and "##" in sh.text_frame.text for sh in s_.shapes) or \
                any(sh.has_table and any("##" in c.text for r in sh.table.rows for c in r.cells) for sh in s_.shapes):
            D.check.append(f"p.{i + 1}: '##' 표시 항목은 회사 확인 후 채움")
    # 첫 페이지 노트에 확인 목록 (페이지 실패 포함)
    first = built[0][0] if built else None
    if first is not None:
        first.notes_slide.notes_text_frame.text = "[확인 목록]\n" + "\n".join(f"- {c}" for c in D.check)
    prs.save(out_path)
    return D


def run(template, inputs, out, today=None):
    spec = load_spec()
    I = dict(inputs)
    if str(I.get("split_auto") or "").lower() in ("1", "true", "on", "yes"):
        P5.auto_split(I, spec, today)
    tenors = [int(x.split(":")[0]) for x in I["split"].split(",")]
    base = int(I.get("base") or sum(int(x.split(":")[1]) for x in I["split"].split(",")))
    args = SimpleNamespace(template=template, corp=I["corp"], issue=I["issue"], base=base, max=int(I.get("max") or base),
                           tenors=",".join(map(str, tenors)), grade=I.get("grade") or None, band=(I.get("band") or None),
                           today=today, split=I["split"], out=None)
    with tempfile.TemporaryDirectory() as td:
        args.out = os.path.join(td, "v3.pptx")
        F = P3.build(args)["facts"]
    if I.get("demand"):
        F["sch"]["수요예측"] = date.fromisoformat(I["demand"])
    X = context(F, I, spec)
    D = build(template, F, I, X, spec, out)
    return out, D.check


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--inputs", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--today")
    a = ap.parse_args()
    o, chk = run(a.template, json.load(open(a.inputs, encoding="utf-8")), a.out, a.today)
    print(o)
    print("\n".join(chk))
