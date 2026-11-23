"""제안서 문구 만들기 — 회사·등급·시장 데이터로 장표에 들어갈 문장과 숫자를 계산한다.

모든 문장은 데이터에서 나온 숫자만 쓴다. 데이터가 없으면 그 문장을 빼거나 ## 로 남긴다.
문구 기준(임계값·우선순위)은 docs/proposal_spec.md 의 '설정 > rules' 에서 바꿀 수 있다.
"""
import os
from datetime import date, timedelta
from types import SimpleNamespace

from tools import proposal_mkt as M

DEFAULT_RULES = {
    "trade_window_days": 30,      # 장외거래로 회사 고유 스프레드를 볼 기간
    "comps_days": 150,            # 비교 사례 기간
    "demand_months": 6,           # 같은 등급군 수요 통계 기간
    "high_debt_ratio": 200,       # 부채비율(%) 이상이면 예상 질문에 넣음
    "low_icr": 3.0,               # 이자보상배율(배) 미만이면 예상 질문에 넣음
    "high_nd_ebitda": 4.0,        # 순차입금/EBITDA(배) 이상이면 예상 질문에 넣음
    "spread_high_pct": 70,        # 1년 스프레드 분위 이상이면 '높은 편'
    "spread_low_pct": 30,         # 이하면 '낮은 편'
    "market_view_max_age_days": 10,  # 설정의 시장 코멘트(market_view)를 쓰는 최대 경과일
}


def series_at(pairs, d):
    best = None
    for dd, v in pairs:
        if dd <= d:
            best = v
        else:
            break
    return best


def won(v, unit="억"):
    """원 단위 → '2.3조' / '5,620억'"""
    if v is None:
        return "-"
    e = v / 1e8
    if unit == "조" or abs(e) >= 10000:
        return f"{e / 1e4:,.1f}조"
    return f"{e:,.0f}억"


def ro(word):
    """'로/으로' — 받침 있으면 '으로' (ㄹ받침 제외)."""
    ch = (word or " ")[-1]
    if "가" <= ch <= "힣":
        jong = (ord(ch) - 0xAC00) % 28
        return "으로" if jong and jong != 8 else "로"
    return "로"


def pct(v, d=1):
    return f"{v:.{d}f}%" if v is not None else "-"


def short(s, n):
    s = (s or "").strip().replace("\n", " ")
    return s if len(s) <= n else s[: n - 1].rstrip(" ,·") + "…"


def context(F, I, spec):
    """장표 공통 계산값."""
    rules = {**DEFAULT_RULES, **((spec.get("rules") or {}))}
    co = F.get("co") or {}
    grade = M.gkey(F["grade"]) or M.gkey(co.get("rating")) or F["grade"]
    cv = M.Curves() if M.gkey(grade) else M.Curves({})  # 회사채 등급을 못 읽으면 곡선을 쓰지 않음 (## 로 남김)
    tenors = sorted(F["split"])
    issue, today = F["issue"], F["today"]
    isins = [b.get("isin") for b in ((co.get("bonds") or {}).get("list") or []) if b.get("isin")]
    adj = M.issuer_spread(F["name"], grade, cv, isins, days=rules["trade_window_days"]) if cv else None
    indiv = M.parse_indiv(I.get("indiv"))
    base = M.base_rates(tenors, grade, cv, indiv, adj) if cv else {t: (None, "##") for t in tenors}
    band_bp = M.parse_band(F["band"]) or (-30, 30)
    dem, dem_rows = M.demand_stats(F["pubs"], grade, today, rules["demand_months"])
    dem_x, _ = M.demand_stats(F["pubs"], grade, today, rules["demand_months"], exact=True)
    cps = M.comps(F["pubs"], grade, today, rules["comps_days"], exclude=co.get("key") or F.get("corp"))
    prev = F["prev"]
    pdm = {int(round(r["_tenor"])): r for r in prev if r.get("_tenor")}
    # 만기별 목표 가산금리: 동사 직전 확정값과 같은 등급군 중앙값의 평균, 밴드 안으로
    tgt = {}
    for t in tenors:
        xs = []
        if t in pdm and pdm[t].get("_spread") is not None:
            xs.append(pdm[t]["_spread"])
        b = dem.get(M.tbucket(t)) or {}
        if b.get("sp_med") is not None:
            xs.append(b["sp_med"])
        v = round(sum(xs) / len(xs)) if xs else 0
        tgt[t] = max(band_bp[0], min(band_bp[1], v))
    exp = {t: (base[t][0] + tgt[t] / 100 if base[t][0] is not None else None) for t in tenors}
    # 일정 주변 수급
    sch = F["sch"]
    d0 = sch["수요예측"]
    wall, wall_n = M.maturity_wall(F["mats"], grade, d0 - timedelta(days=7), issue + timedelta(days=14)) if F.get("mats") else (0, 0)
    month_wall, _ = M.maturity_wall(F["mats"], grade, date(issue.year, issue.month, 1),
                                    date(issue.year + issue.month // 12, issue.month % 12 + 1, 1) - timedelta(days=1)) if F.get("mats") else (0, 0)
    pipe = M.pipeline(F["pubs"], grade, d0 - timedelta(days=7), d0 + timedelta(days=7), exclude=F.get("corp"))
    ly = M.last_year_same_month(F["pubs"], grade, issue.month, issue.year - 1)
    asof = cv.last if cv else None
    outlook = None
    try:
        from tools import proposal_outlook as OL
        outlook = OL.load_latest(os.environ.get("RESEARCH_DIR") or (spec.get("research") or {}).get("dir"), today)
    except Exception as e:  # noqa: BLE001 — 전망 파일 문제로 제안서 전체가 멈추지 않게
        print("outlook skipped:", type(e).__name__)
    return SimpleNamespace(rules=rules, co=co, grade=grade, band=M.band(grade), cv=cv, asof=asof, tenors=tenors, adj=adj,
                           indiv=indiv, base=base, band_bp=band_bp, dem=dem, dem_rows=dem_rows, dem_x=dem_x, comps=cps,
                           pdm=pdm, tgt=tgt, exp=exp, wall=wall, wall_n=wall_n, month_wall=month_wall, pipe=pipe, ly=ly,
                           outlook=outlook)


# ───────── 등급·스프레드 문장 ─────────
def spread_line(X, tenor=3):
    cv, g = X.cv, X.grade
    if not cv:
        return None
    sp = cv.sp(g, tenor)
    rng = cv.range(g, tenor)
    p = cv.pct_rank(g, tenor)
    c1 = cv.change(g, tenor, 30, spread=True)
    lv = ("1년 중 높은 편" if p is not None and p >= X.rules["spread_high_pct"] else
          "1년 중 낮은 편" if p is not None and p <= X.rules["spread_low_pct"] else "1년 범위의 중간")
    s = f"{g} {tenor}년 등급민평 {cv.y(g, tenor):.3f}%, 국고 대비 {sp:.0f}bp — {lv}"
    if rng:
        s += f"(최근 1년 {rng[0][1]:.0f}~{rng[1][1]:.0f}bp)"
    if c1 is not None:
        s += f", 1개월 {c1:+.0f}bp"
    return s


def adj_line(X, name):
    if X.indiv:
        return f"{name} 개별민평(입력값) 기준으로 산정"
    a = X.adj
    if not a:
        return None
    rel = "낮게" if a["bp"] < 0 else "높게" if a["bp"] > 0 else "비슷하게"
    return (f"최근 {a['days']}일 장외에서 동사 채권 {a['n']}건(거래대금 {a['amt']:,}억)이 {X.grade} 등급민평보다 "
            f"{abs(a['bp'])}bp {rel} 거래 — 같은 등급 내 {'우량' if a['bp'] < 0 else '열위' if a['bp'] > 0 else '평균'} 평가")


def term_premium(X, t0, t1):
    cv = X.cv
    if not cv:
        return None
    a, b = cv.y(X.grade, t0), cv.y(X.grade, t1)
    return (b - a) * 100 if a is not None and b is not None else None


def demand_line(X, t):
    b = X.dem.get(M.tbucket(t))
    if not b:
        return None
    s = (f"최근 {X.rules['demand_months']}개월 {X.band}급 {M.tbucket(t)}물 {b['n']}건 · 주문 {M.eok(b['orders'])}원"
         + (f" · 평균 경쟁률 {b['avg_ratio']:.1f}:1" if b.get("avg_ratio") else ""))
    if b.get("sp_med") is not None:
        s += f" · 확정 가산금리 중앙값 {b['sp_med']:+d}bp"
    if b.get("under"):
        s += f" (미매각 {b['under']}건)"
    return s


# ───────── 크레딧 스토리·예상 질문 (데이터 기반) ─────────
def credit_story(F, X, I):
    co = X.co
    m = co.get("metrics") or {}
    fi = co.get("fs_interim") or co.get("fs_annual") or {}
    ai = co.get("ai") or {}
    trig = co.get("triggers") or {}
    ag = trig.get("agencies") or []
    holders = (co.get("holders") or {}).get("top") or []
    bonds = co.get("bonds") or {}
    name = F["name"]
    notes_ = [x.strip() for x in ((I.get("credit_note") or "") + "\n" + (I.get("issues_note") or "")).splitlines() if x.strip()]

    pts = []
    # 1 사업
    biz = ai.get("one_liner") or (co.get("business") or {}).get("overview_short")
    if biz and not ai.get("one_liner"):  # 사업보고서 원문 요약은 첫 문장만
        biz = biz.split(". ")[0].split("다. ")[0]
    if not biz and ((co.get("listing") or {}).get("products") or co.get("industry")):
        biz = f"{co.get('industry', '')} — 주요 제품·사업: {(co.get('listing') or {}).get('products', '')}".strip(" —")
    mix = ai.get("mix_line")
    if biz:
        pts.append(("사업 기반", short(biz, 80) + (f" ({mix})" if mix else "")))
    # 2 실적
    if m.get("revenue"):
        s = f"{m.get('period', '')} 매출 {won(m['revenue'])}"
        if m.get("revenue_growth") is not None:
            s += f"(전년比 {m['revenue_growth']:+.1f}%)"
        if m.get("op") is not None:
            s += f", 영업이익 {won(m['op'])}"
        if m.get("op_margin") is not None:
            s += f"(이익률 {m['op_margin']:.1f}%)"
        if m.get("ebitda"):
            s += f", EBITDA {won(m['ebitda'])}"
        pts.append(("실적", s))
    # 3 재무·유동성
    liq = (m.get("cash") or 0) + (m.get("st_deposits") or 0)
    parts = []
    if liq and bonds.get("due12"):
        parts.append(f"현금성자산 {won(liq)}(12개월 내 사채 만기 {won(bonds['due12'])}의 {liq / bonds['due12']:.1f}배)")
    elif liq:
        parts.append(f"현금성자산 {won(liq)}")
    if m.get("ocf"):
        parts.append(f"영업현금흐름 {won(m['ocf'])}")
    if m.get("net_debt_ebitda") is not None:
        parts.append(f"순차입금/EBITDA {m['net_debt_ebitda']:.1f}배")
    if parts:
        pts.append(("재무·유동성", ", ".join(parts)))
    # 4 신용등급·지배구조
    if ag:
        s = " · ".join(f"{a.get('agency')} {a.get('rating', '')}" + (f"({a['outlook']})" if a.get("outlook") else "") for a in ag)
    else:
        s = f"신용등급 {X.grade}"
    if holders:
        h = holders[0]
        s += f", 최대주주 {h['name']} {h['pct']:.1f}%"
    if X.adj and X.adj["bp"] < 0:
        s += f" — 장외에서 등급 대비 {abs(X.adj['bp'])}bp 우량하게 거래"
    pts.append(("신용도", s))
    # 메모로 받은 강점은 앞쪽에
    for i, n_ in enumerate(notes_[:2]):
        pts.insert(i, ("핵심 포인트", short(n_, 90)))
    pts = pts[:4]
    while len(pts) < 4:
        pts.append(("##", "##"))

    # 예상 질문
    R = X.rules
    qa = []
    if m.get("debt_ratio") is not None and m["debt_ratio"] >= R["high_debt_ratio"]:
        a = f"차입금 {won(m.get('debt'))}"
        if "리스부채" in (fi.get("debt_items") or []):
            a += "(리스부채 포함)"
        if m.get("ebitda"):
            a += f" 대비 EBITDA {won(m['ebitda'])}"
        if m.get("ocf"):
            a += f", 영업현금흐름 {won(m['ocf'])}{ro(won(m['ocf']))} 상환 재원 확보"
        qa.append((f"부채비율 {m['debt_ratio']:.0f}%로 높은데?", a))
    if m.get("interest_coverage") is not None and m["interest_coverage"] < R["low_icr"]:
        qa.append((f"이자보상배율 {m['interest_coverage']:.1f}배 — 이자 부담은?",
                    f"EBITDA 기준으로는 {(m['ebitda'] / fi['interest_paid']):.1f}배" if m.get("ebitda") and fi.get("interest_paid") else
                    "고정금리 장기화로 금리 변동 노출 축소"))
    if m.get("net_debt_ebitda") is not None and m["net_debt_ebitda"] >= R["high_nd_ebitda"]:
        cx = m.get("capex") or fi.get("capex")
        qa.append((f"순차입금/EBITDA {m['net_debt_ebitda']:.1f}배 — 감축 계획은?",
                   (f"영업현금흐름 {won(m.get('ocf'))} vs 설비투자 {won(cx)}" + (f" — 연간 {won(m['ocf'] - cx)} 잉여현금으로 차입 축소 여력" if cx and m["ocf"] > cx else " — 투자 조정 계획 ##"))
                   if m.get("ocf") else "##"))
    capex = m.get("capex") or fi.get("capex")
    if capex and m.get("ocf") and capex > m["ocf"]:
        qa.append(("설비투자가 영업현금흐름보다 큰데?", f"{fi.get('period', '')} 투자 {won(capex)} vs 영업현금흐름 {won(m['ocf'])} — 부족분은 회사채·차입으로 조달, 이번 발행이 그 일부"))
    if ag:
        out = [a for a in ag if a.get("outlook")]
        items = trig.get("items") or []
        down = [x for x in items if x.get("dir") == "down"]
        a = (" · ".join(f"{x['agency']} {x['outlook']}" for x in out) or f"현 등급 {X.grade} 유지")
        if down:
            a += f" / 하향 검토 요인: {short(down[0].get('text'), 48).rstrip(',')}"
        qa.append(("신용등급 전망은?", a))
    if bonds.get("due12"):
        qa.append((f"12개월 내 만기 {won(bonds['due12'])} 차환은?",
                   f"이번 {F['base']:,}억(최대 {F['max']:,}억) 발행과 현금성자산 {won(liq)}{ro(won(liq))} 대응" if liq else f"이번 {F['base']:,}억 발행으로 일부 선조달"))
    if m.get("op_margin") is not None:
        ans = ai.get("caption") or (mix and f"매출 구성: {mix}")
        if not ans and m.get("ebitda") and m.get("revenue"):
            ans = f"EBITDA 마진 {m['ebitda'] / m['revenue'] * 100:.1f}%" + (f", 매출 {m['revenue_growth']:+.1f}% 성장" if m.get("revenue_growth") is not None else "")
        if ans:
            qa.append((f"영업이익률 {m['op_margin']:.1f}% — 수익성 지속성은?", short(ans, 80)))
    if m.get("revenue_growth") is not None and m["revenue_growth"] < 0:
        qa.insert(0, (f"매출이 {m['revenue_growth']:.1f}% 줄었는데?", "##(업황·회복 근거)"))
    # 부족하면 발행 자체에 대한 질문으로 채움
    extra = []
    if X.adj or X.pdm:
        sp_prev = [r.get("_spread") for r in X.pdm.values() if r.get("_spread") is not None]
        a = []
        if X.adj:
            a.append(f"장외에서 {X.grade} 등급민평 대비 {X.adj['bp']:+d}bp")
        if sp_prev:
            a.append(f"직전 회차 확정 {min(sp_prev):+d}bp")
        extra.append(("등급 대비 금리 수준은 적정한가?", " / ".join(a) + f" — 목표 가산금리 " + " / ".join(f"{t}년 {X.tgt[t]:+d}bp" for t in X.tenors)))
    use = (I.get("issues_note") or "").strip().splitlines()
    extra.append(("이번 조달 자금 용도는?", short(use[0], 70) if use else
                  (f"12개월 내 사채 만기 {won(bonds['due12'])} 차환" if bonds.get("due12") else "운영·시설자금 (증권신고서 기준 ##)")))
    if holders:
        extra.append(("지배구조·계열 지원 가능성은?", f"최대주주 {holders[0]['name']} {holders[0]['pct']:.1f}%" + (f" (특수관계인 포함 {co['holders'].get('total_pct', 0):.1f}%)" if (co.get('holders') or {}).get('total_pct') else "")))
    for q in extra:
        if len(qa) >= 5:
            break
        qa.append(q)
    qa = qa[:5]
    while len(qa) < 5:
        qa.append(("##", "##"))
    # 수요예측 메시지 한 줄
    tag = []
    if co.get("industry"):
        tag.append(co["industry"])
    if m.get("revenue"):
        tag.append(f"매출 {won(m['revenue'])}")
    if m.get("ocf"):
        tag.append(f"영업현금흐름 {won(m['ocf'])}")
    if ag:
        a0 = ag[0]
        tag.append(f"{a0.get('rating', X.grade)}" + (f"({a0['outlook']})" if a0.get("outlook") else ""))
    msg = " · ".join(tag) if tag else "##"
    return pts, qa, msg


# ───────── 시장 해석 (국내·미국) ─────────
def kr_lines(rt, cv=None):
    bs, k3, k10 = rt.get("base") or [], rt.get("ktb3") or [], rt.get("ktb10") or []
    if not bs or not k3:
        return [], None
    d = k3[-1][0]
    b = bs[-1][1]
    k3v, k10v = k3[-1][1], series_at(k10, d)
    gap = (k3v - b) * 100
    k3m = series_at(k3, (date.fromisoformat(d) - timedelta(days=30)).isoformat())
    view = ("시장은 추가 인상 또는 높은 금리의 장기화를 반영" if gap >= 40 else
            "시장은 기준금리 동결을 중심으로 반영" if gap > -15 else "시장은 기준금리 인하를 반영")
    L = [f"기준금리 {b:.2f}%, 국고 3년 {k3v:.3f}% — 3년물이 기준금리보다 {gap:+.0f}bp: {view}"]
    if k3m is not None:
        L.append(f"- 국고 3년 1개월 {(k3v - k3m) * 100:+.0f}bp, 10년 {k10v:.3f}%(10년-3년 {(k10v - k3v) * 100:+.0f}bp)")
    yr = (date.fromisoformat(d) - timedelta(days=365)).isoformat()
    ch = [(d1, v0, v1) for (d0, v0), (d1, v1) in zip(bs, bs[1:]) if v1 != v0 and d1 >= yr]
    if ch:
        L.append(f"최근 1년 기준금리 {len(ch)}회 변경: " + ", ".join(f"{x[0][2:7].replace('-', '.')} {x[1]:.2f}→{x[2]:.2f}%" for x in ch[-4:]))
    k3y = [v for dd, v in k3 if dd >= yr]
    if k3y:
        L.append(f"- 국고 3년 최근 1년 {min(k3y):.2f}~{max(k3y):.2f}% 범위, 현재 {'상단' if k3v >= max(k3y) - (max(k3y) - min(k3y)) / 3 else '하단' if k3v <= min(k3y) + (max(k3y) - min(k3y)) / 3 else '중간'}")
    return L, gap


def us_lines(rt):
    u2, u10, hi, lo = rt.get("ust2") or [], rt.get("ust10") or [], rt.get("ffr_hi") or [], rt.get("ffr_lo") or []
    if not u2 or not hi:
        return [], None
    d = u2[-1][0]
    mid = (hi[-1][1] + lo[-1][1]) / 2 if lo else hi[-1][1] - 0.125
    gap = (u2[-1][1] - mid) * 100
    m1 = series_at(u2, (date.fromisoformat(d) - timedelta(days=30)).isoformat())
    view = ("2년물이 정책금리를 웃돌아 추가 인상 가능성을 반영" if gap >= 25 else
            "2년물이 정책금리 부근 — 동결 중심" if gap > -25 else "2년물이 정책금리를 밑돌아 인하 기대를 반영")
    L = [f"美 정책금리 {lo[-1][1]:.2f}~{hi[-1][1]:.2f}%, 국채 2년 {u2[-1][1]:.2f}%({d[5:].replace('-', '/')}) — {view}"]
    if u10:
        L.append(f"- 국채 10년 {u10[-1][1]:.2f}%(10년-2년 {(u10[-1][1] - u2[-1][1]) * 100:+.0f}bp), 2년 1개월 {((u2[-1][1] - m1) * 100) if m1 is not None else 0:+.0f}bp")
    yr = (date.fromisoformat(d) - timedelta(days=365)).isoformat()
    ch = [(d1, v0, v1) for (d0, v0), (d1, v1) in zip(hi, hi[1:]) if v1 != v0 and d1 >= yr]
    if ch:
        L.append(f"최근 1년 정책금리(상단) {len(ch)}회 변경: " + ", ".join(f"{x[0][2:7].replace('-', '.')} {x[1]:.2f}→{x[2]:.2f}%" for x in ch[-4:]))
    u2y = [v for dd, v in u2 if dd >= yr]
    if u2y:
        L.append(f"- 국채 2년 최근 1년 {min(u2y):.2f}~{max(u2y):.2f}% 범위")
    return L, gap


def fresh_view(spec, key, today, rules):
    mv = spec.get("market_view") or {}
    asof = str(mv.get("asof") or "")
    try:
        age = (today - date.fromisoformat(asof)).days
    except ValueError:
        return []
    if age > rules["market_view_max_age_days"]:
        return []
    return [str(x) for x in (mv.get(key) or [])]
