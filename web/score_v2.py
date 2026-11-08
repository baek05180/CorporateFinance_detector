"""조달 방식 점수 v2 (시험): 관문 → 재무 40 · 시장 30 · 조건 20 · 이력 10."""
import math
from datetime import date, timedelta

GRADES = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-", "BB+", "BB", "BB-", "B+", "B", "B-"]


def _g(s):
    if not s:
        return None
    s = s.replace("0", "").strip()
    for g in sorted(GRADES, key=len, reverse=True):
        if s.startswith(g):
            return g
    return None


def _gi(g):
    return GRADES.index(g) if g in GRADES else None


def _sum(parts):
    return sum(p for _, p, _, _ in parts)


def _years_since(d, today):
    return (today - date.fromisoformat(d[:10])).days / 365 if d else 99


def market_ctx(deals, today):
    """같은 등급대 최근 6개월 결정 가산금리 평균(bp)"""
    since = (today - timedelta(days=183)).isoformat()
    by = {}
    for r in deals:
        if r.get("type") == "BOND" and r.get("offering") == "공모" and r.get("fix_spread") is not None and (r.get("rcept_dt") or "") >= since:
            g = _g(r.get("rating"))
            if g:
                b = "AA" if _gi(g) <= 3 else "A" if _gi(g) <= 6 else "BBB"
                by.setdefault(b, []).append(r["fix_spread"])
    return {k: sum(v) / len(v) for k, v in by.items()}


def score(co, mctx, today):
    m, fa, T = co.get("metrics") or {}, co.get("fs_annual") or {}, co.get("triggers") or {}
    deals, BP, H = co.get("deals") or [], co.get("bonds") or {}, co.get("holders") or {}
    mk = co.get("market")
    grade = _g(co.get("rating_src") or co.get("rating")) or next((_g(a.get("rating")) for a in T.get("agencies", [])), None)
    gi = _gi(grade)
    outlooks = [a.get("outlook") or "" for a in T.get("agencies", [])]
    debt, cash = fa.get("debt"), (fa.get("cash") or 0) + (fa.get("st_deposits") or 0)
    e = m.get("ebitda") if m.get("ebitda_ok") else None
    nde = (debt - cash) / e if debt is not None and e and e > 0 else None
    icr, dr, mcap = m.get("interest_coverage"), m.get("debt_ratio"), m.get("mcap")
    impaired = (m.get("equity") is not None and m.get("capital") and m["equity"] < m["capital"])
    full_imp = m.get("equity") is not None and m["equity"] <= 0
    loss = (m.get("op") or 0) < 0
    last = lambda t, off=None: max((d.get("board_dt") or d.get("rcept_dt") or d.get("issue_dt") or "" for d in deals
                                    if d.get("type") == t and (off is None or d.get("offering") == off)), default="")
    cnt3 = lambda t, off=None: sum(1 for d in deals if d.get("type") == t and (off is None or d.get("offering") == off)
                                   and _years_since(d.get("board_dt") or d.get("rcept_dt") or d.get("issue_dt"), today) <= 3)
    ph = co.get("price_hist") or []
    vol = None
    if len(ph) > 20:
        rets = [math.log(ph[i][1] / ph[i - 1][1]) for i in range(1, len(ph)) if ph[i - 1][1] and ph[i][1]]
        vol = (sum(r * r for r in rets) / len(rets)) ** .5 * math.sqrt(250) * 100 if rets else None
    halted = vol is not None and vol < 3 and mk in ("K", "Y", "N")
    out = {}

    # ── 공모채 ──
    P = []
    gate = "신용등급 미확인 — 평가사 등급 확인 필요" if gi is None else ("자본잠식" if impaired else None)
    down = [x for x in T.get("items", []) if x.get("dir") == "down" and "cur" in x]
    if down:
        if any(x["met"] for x in down):
            P.append(("재무", 0, "트리거", "하향 트리거 충족"))
        else:
            rel = min(abs(x["gap"]) / abs(x["thr"]) for x in down if x["thr"])
            P.append(("재무", 25 if rel >= .3 else 15 if rel >= .1 else 5, "트리거 여유", f"최소 {rel*100:.0f}%"))
    elif nde is not None:
        P.append(("재무", 25 if nde <= 2 else 15 if nde <= 4 else 5 if nde <= 6 else 0, "순차입금/EBITDA", f"{nde:.1f}배"))
    elif dr is not None:
        P.append(("재무", 20 if dr <= 100 else 12 if dr <= 200 else 5 if dr <= 300 else 0, "부채비율", f"{dr:.0f}%"))
    if icr is not None:
        P.append(("재무", 10 if icr >= 5 else 6 if icr >= 2 else 2 if icr >= 1 else 0, "이자보상배율", f"{icr:.1f}배"))
    if outlooks:
        neg = any(k in o for o in outlooks for k in ("부정", "하향"))
        pos = any(k in o for o in outlooks for k in ("긍정", "상향"))
        P.append(("재무", 0 if neg else 5 if pos else 3, "등급 전망", "/".join(sorted(set(o for o in outlooks if o)))))
    if gi is not None:
        P.append(("시장", 20 if gi <= 3 else 15 if gi <= 6 else 5 if gi <= 9 else 0, "신용등급", grade))
        b = "AA" if gi <= 3 else "A" if gi <= 6 else "BBB"
        if b in mctx:
            s = mctx[b]
            P.append(("시장", 10 if s <= 0 else 6 if s <= 10 else 2, f"{b}급 최근 결정금리", f"민평 {s:+.0f}bp"))
    rev, assets = m.get("revenue") or 0, m.get("assets") or 0
    P.append(("조건", 10 if rev >= 5e11 or assets >= 1e12 else 5 if rev >= 1e11 else 0, "규모", f"매출 {rev/1e8:,.0f}억"))
    if (BP.get("due12_by") or {}).get("공모"):
        P.append(("조건", 10, "12개월 내 공모채 만기", f"{BP['due12_by']['공모']/1e8:,.0f}억"))
    if cnt3("BOND", "공모"):
        P.append(("이력", 10, "최근 3년 공모채", f"{cnt3('BOND', '공모')}건"))
    out["공모채"] = (gate, _sum(P), P)

    # ── 사모채 ──
    P = []
    gate = "완전자본잠식" if full_imp else None
    if icr is not None:
        P.append(("재무", 15 if icr >= 3 else 10 if icr >= 1.5 else 5 if icr >= 1 else 0, "이자보상배율", f"{icr:.1f}배"))
    if m.get("current_ratio") is not None:
        cr = m["current_ratio"]
        P.append(("재무", 10 if cr >= 150 else 6 if cr >= 100 else 2, "유동비율", f"{cr:.0f}%"))
    if dr is not None:
        P.append(("재무", 15 if dr <= 150 else 8 if dr <= 300 else 3 if dr <= 500 else 0, "부채비율", f"{dr:.0f}%"))
    P.append(("시장", 20 if gi is None or gi >= 7 else 10 if gi >= 4 else 0, "등급 구간", grade or "등급 없음"))
    P.append(("시장", 0 if loss else 10, "영업이익", "적자" if loss else "흑자"))
    P.append(("조건", 10 if assets >= 1e11 else 5, "자산 규모", f"{assets/1e8:,.0f}억"))
    if (BP.get("due12_by") or {}).get("사모"):
        P.append(("조건", 10, "12개월 내 사모채 만기", f"{BP['due12_by']['사모']/1e8:,.0f}억"))
    if cnt3("BOND", "사모"):
        P.append(("이력", 10, "최근 3년 사모채", f"{cnt3('BOND', '사모')}건"))
    out["사모채"] = (gate, _sum(P), P)

    # ── CB ──
    P = []
    gate = None if mk in ("K", "Y", "N") else "비상장"
    if full_imp:
        gate = "완전자본잠식"
    if halted:
        gate = "6개월간 주가 변동 없음 — 거래정지 여부 확인"
    tp = H.get("total_pct")
    if tp is not None:
        after = tp / 1.1
        P.append(("재무", 15 if after >= 30 else 10 if after >= 20 else 5 if after >= 15 else 0, "희석 후 최대주주", f"{tp:.1f}%→{after:.1f}% (시총 10% 발행 가정)"))
    link_bal = sum(d.get("amount") or 0 for d in deals if d.get("type") in ("CB", "BW", "EB")
                   and (d.get("maturity") or "9") > today.isoformat() and _years_since(d.get("board_dt") or d.get("rcept_dt"), today) <= 3)
    if mcap:
        oh = link_bal / mcap * 100
        P.append(("재무", 15 if oh < 5 else 8 if oh < 15 else 0, "최근 3년 메자닌 발행(미만기)", f"시총의 {oh:.0f}%"))
        if (link_bal + mcap * 0.1) / mcap > 0.2:
            P.append(("재무", -10, "발행 후 메자닌 잔액", f"시총의 {(link_bal / mcap + 0.1) * 100:.0f}% (20% 초과 — 오버행)"))
    P.append(("재무", 10 if loss or (m.get("op_margin") or 0) < 3 else 5, "수익성", "저수익·적자 → 저쿠폰 CB 유리" if loss or (m.get("op_margin") or 0) < 3 else "흑자"))
    P.append(("시장", {"K": 20, "N": 10, "Y": 10}.get(mk, 0), "상장 시장", {"K": "코스닥", "N": "코넥스", "Y": "유가증권"}.get(mk, "비상장")))
    if vol is not None and not halted:
        P.append(("시장", 10 if 30 <= vol <= 80 else 5, "주가 변동성(연)", f"{vol:.0f}%"))
    if mcap:
        P.append(("조건", 10 if mcap >= 5e10 else 5 if mcap >= 2e10 else 0, "시가총액", f"{mcap/1e8:,.0f}억"))
    if dr is not None:
        P.append(("조건", 10 if dr > 200 else 5, "부채비율", f"{dr:.0f}% {'→ 전환 시 자본 확충 효과' if dr > 200 else ''}".strip()))
    if gi is not None and gi <= 6:
        P.append(("조건", -25, "신용등급", f"{grade} — 일반 사채로 희석 없이 조달 가능"))
    elif gi is None and (mcap or 0) >= 5e12:
        P.append(("조건", -25, "시가총액", f"{mcap/1e12:,.0f}조 — 대형 우량주는 희석 없는 사채가 일반적"))
    n = cnt3("CB")
    P.append(("이력", 10 if 1 <= n <= 2 else 3 if n >= 3 else 5, "최근 3년 CB", f"{n}건" + (" (누적 오버행)" if n >= 3 else "")))
    out["CB"] = (gate, _sum(P), P)

    # ── EB ──
    P = []
    ts, sh = co.get("treasury_shares"), m.get("shares")
    tv = ts * m["close"] if ts and m.get("close") else None  # 자사주 시가
    gate = None if ts else "자기주식 없음"
    if ts and tv is not None and tv < 30e8:
        gate = f"자사주 시가 {tv/1e8:,.0f}억 — 30억 미만이라 EB 실익 낮음"
    elif ts and sh and ts / sh < 0.01:
        gate = f"자사주 {ts / sh * 100:.2f}% — 발행주식의 1% 미만이라 EB 실익 낮음"
    if halted:
        gate = "6개월간 주가 변동 없음 — 거래정지 여부 확인"
    if tv is not None:
        P.append(("재무", 40 if tv >= 300e8 else 30 if tv >= 100e8 else 20 if tv >= 50e8 else 10, "자사주 시가", f"{tv/1e8:,.0f}억" + (f" (발행주식의 {ts / sh * 100:.1f}%)" if sh else "")))
    P.append(("시장", 20 if mk in ("K", "Y") else 5, "상장 시장", mk or "-"))
    if mcap:
        P.append(("시장", 10 if mcap >= 1e11 else 5, "시가총액", f"{mcap/1e8:,.0f}억"))
    P.append(("조건", 10, "규제", "자사주 EB 공시 강화 — 주주 설명 필요"))
    if cnt3("EB"):
        P.append(("이력", 10, "최근 3년 EB", f"{cnt3('EB')}건"))
    out["EB"] = (gate, _sum(P), P)

    # ── 유상증자 (상장: 주주배정·일반공모·제3자배정 / 비상장·적자 성장기업: 제3자배정 RCPS) ──
    P = []
    gate = "6개월간 주가 변동 없음 — 거래정지 여부 확인" if halted else None
    f = 0
    if impaired:
        f += 30
    if dr is not None and dr > 300:
        f += 15
    if (m.get("loss_years") or 0) >= 2:
        f += 10
    P.append(("재무", min(40, f), "자본 확충 필요", (("자본잠식 " if impaired else "") + (f"부채비율 {dr:.0f}% " if dr else "")
                                                   + (f"{m.get('loss_years')}년 연속 적자" if (m.get('loss_years') or 0) >= 2 else "")).strip() or "낮음"))
    if mk in ("K", "Y", "N"):
        if tp is not None:
            P.append(("시장", 15 if tp >= 30 else 8 if tp >= 20 else 3, "최대주주 지분(참여 여력)", f"{tp:.1f}%"))
        if len(ph) > 20 and not halted:
            dd = (ph[-1][1] / max(p[1] for p in ph) - 1) * 100
            P.append(("시장", 5 if dd < -30 else 15, "6개월 고점 대비", f"{dd:.0f}%"))
        if mcap:
            P.append(("조건", 10 if mcap >= 3e10 else 5, "시가총액", f"{mcap/1e8:,.0f}억"))
    else:
        g = m.get("revenue_growth")
        P.append(("시장", 15 if (g or 0) >= 20 else 8 if (g or 0) >= 0 else 3, "매출 성장률(RCPS 투자 매력)", f"{g:.0f}%" if g is not None else "정보 없음"))
        vc = sum(1 for d in deals if d.get("type") in ("CB", "BW", "EB") for i in (d.get("investors") or []) if i.get("kind") in ("신기사·VC", "PEF"))
        P.append(("시장", 10 if vc else 3, "VC·PEF 투자 이력", f"{vc}건" if vc else "확인 안 됨"))
        P.append(("조건", 5, "방식", "비상장 — 제3자배정 RCPS·CPS"))
    lr = last("RIGHTS")
    yrs = _years_since(lr, today) if lr else 99
    P.append(("이력", -20 if yrs < 1 else -10 if yrs < 2 else 10, "최근 유상증자", f"{lr[:10]} ({yrs:.1f}년 전)" if lr else "없음"))
    out["유상증자"] = (gate, _sum(P), P)

    # ── P-CBO (신용보증기금 유동화 보증) ──
    P = []
    big = rev >= 3e11 or assets >= 5e11 or (mcap or 0) >= 1e12
    gate = "재무정보 없음" if not rev else "자본잠식 — 보증 심사 통과 어려움" if impaired else ("대기업·우량 등급 — 지원 대상 아님" if big or (gi is not None and gi <= 6) else None)
    P.append(("재무", 15 if icr is not None and icr >= 1.5 else 8 if icr is not None and icr >= 1 else 0, "이자보상배율", f"{icr:.1f}배" if icr is not None else "정보 없음"))
    ly = m.get("loss_years") or 0
    P.append(("재무", 15 if ly == 0 else 5 if ly == 1 else 0, "영업손실 연수", f"{ly}년"))
    if dr is not None:
        P.append(("재무", 10 if dr <= 300 else 3, "부채비율", f"{dr:.0f}%"))
    P.append(("시장", 20 if gi is None or gi >= 7 else 5, "등급 구간", (grade or "등급 없음") + " — 단독 사모보다 보증부가 금리 유리" if gi is None or gi >= 7 else grade))
    P.append(("시장", 10, "규모", f"중소·중견 (매출 {rev/1e8:,.0f}억)"))
    P.append(("조건", 10 if (BP.get("due12") or 0) > 0 else 5, "12개월 내 사채 만기", f"{(BP.get('due12') or 0)/1e8:,.0f}억"))
    if BP.get("has_cbo") or any(d.get("pcbo") for d in deals):
        P.append(("이력", 10, "P-CBO 편입 이력", "있음"))
    out["P-CBO"] = (gate, _sum(P), P)

    # ── 영구채 (신종자본증권) ──
    P = []
    gate = None if gi is not None else "신용등급 평정 이력 없음 — 영구채는 등급 보유 회사 대상"
    if impaired:
        gate = "자본잠식"
    lev = [x for x in T.get("items", []) if x.get("dir") == "down" and "cur" in x and x["metric"] in ("부채비율", "차입금의존도", "순차입금/자본")]
    if lev:
        rel = min(abs(x["gap"]) / abs(x["thr"]) for x in lev if x["thr"])
        P.append(("재무", 30 if rel < .1 or any(x["met"] for x in lev) else 15 if rel < .3 else 5, "레버리지 트리거 여유", "충족" if any(x["met"] for x in lev) else f"{rel*100:.0f}%"))
    elif dr is not None:
        P.append(("재무", 30 if dr > 300 else 15 if dr > 200 else 0, "부채비율(자본 인정 효과)", f"{dr:.0f}%"))
    if icr is not None:
        P.append(("재무", 10 if icr >= 2 else 3, "이자보상배율", f"{icr:.1f}배"))
    if gi is not None:
        P.append(("시장", 20 if gi <= 6 else 10, "신용등급", grade))
    P.append(("시장", 10 if mk == "Y" else 5, "상장 시장", {"Y": "유가증권", "K": "코스닥", "N": "코넥스"}.get(mk, "비상장")))
    P.append(("조건", 5, "금리", "일반 사채보다 높고 콜(통상 5년) 이후 step-up 조건 확인"))
    out["영구채"] = (gate, _sum(P), P)
    return out


KEY = {"공모채": "public_bond", "사모채": "private_bond", "P-CBO": "pcbo", "CB": "cb", "EB": "eb", "유상증자": "rights", "영구채": "hybrid"}
TIE = ["public_bond", "private_bond", "pcbo", "hybrid", "cb", "eb", "rights"]


def _r(v, unit=10e8):
    return max(unit, round(v / unit) * unit) if v and v > 0 else None


def _eok(v):
    return f"{v/1e8:,.0f}억" if v < 1e12 else f"{v/1e12:,.1f}조".replace(".0조", "조")


def sizes(co, today):
    """상품별 예상 발행 규모 — 공모·사모채는 같은 '조달 필요액'에서 출발하고 시장 제약만 달리 적용."""
    m, B, deals = co.get("metrics") or {}, co.get("bonds") or {}, co.get("deals") or []
    assets, mcap, close = m.get("assets"), m.get("mcap"), m.get("close")
    out = {}
    due = sum((B.get("due12_by") or {}).get(k, 0) or 0 for k in ("공모", "사모"))
    past = [d.get("amount") for d in deals if d.get("type") == "BOND" and not d.get("pcbo") and d.get("amount")
            and _years_since(d.get("issue_dt") or d.get("rcept_dt"), today) <= 3]
    avg = sum(past) / len(past) if past else 0
    need, basis = (due, "12개월 내 회사채 만기") if due >= avg else (avg, "최근 3년 회사채 평균 발행액")
    if need and assets:
        if need > assets * 0.1:
            need, basis = assets * 0.1, "총자산의 10% 상한"
    elif not need and assets:
        need, basis = assets * 0.05, "총자산의 5% (발행 이력·만기 없음)"
    if need:
        pub = _r(max(need, 300e8), 50e8)
        out["public_bond"] = f"{_eok(pub)} ({basis}" + (", 공모 최소 300억" if need < 300e8 else "") + ")"
        out["_pub_short"] = need < 300e8
        cap = min(1000e8, assets * 0.05) if assets else 1000e8
        pv = _r(min(need, cap))
        out["private_bond"] = f"{_eok(pv)} ({basis}" + (", 사모 상한 적용" if need > cap else "") + ")"
    # 영구채: 자본 인정 효과가 의미 있는 규모 — 조달액으로 차입금을 갚는다고 보고 부채비율을 20%p 낮추는 금액
    L, E = m.get("liab"), m.get("equity")
    if L and E and E > 0:
        d = L / E
        tgt = d - 0.2
        lev = [x for x in ((co.get("triggers") or {}).get("items") or []) if x.get("metric") == "부채비율" and x.get("thr")]
        if lev and lev[0]["thr"] / 100 * 0.9 < tgt:
            tgt, why = lev[0]["thr"] / 100 * 0.9, f"부채비율 {d*100:.0f}%→{lev[0]['thr']*0.9:.0f}% (트리거 {lev[0]['thr']:.0f}%의 90%)"
        else:
            why = f"부채비율 {d*100:.0f}%→{tgt*100:.0f}% (20%p 인하)"
        x = (L - tgt * E) / (1 + tgt) if tgt > 0 else None
        if x:
            hi = E * 0.3
            if x > hi:
                x, why = hi, why + ", 자본의 30% 상한"
            x = max(x, 300e8)
            nd = (L - x) / (E + x) * 100
            out["hybrid"] = f"{_eok(_r(x, 50e8))} ({why}; 발행 후 약 {nd:.0f}%)"
            out["_hy_small"] = d < 1.0
    ts = co.get("treasury_shares")
    if ts and close:
        eb = ts * close * 0.9
        out["eb"] = f"{_eok(_r(eb))} (자사주 시가의 약 90%)"
    if mcap:
        rs = mcap * 0.1
        eq, cap_ = m.get("equity"), m.get("capital")
        note = "시총의 약 10%"
        if eq is not None and cap_ and eq < cap_:
            rs, note = max(rs, cap_ - eq), "자본잠식 보전액 이상"
        rs = min(max(rs, mcap * 0.05), mcap * 0.2) if note.startswith("시총") else min(rs, mcap * 0.3)
        out["rights"] = f"{_eok(_r(rs))} ({note})"
    return out


def apply(options, co, mctx, today):
    """기존 제안 목록(규모·비교 사례 유지)에 v2 점수를 덮어쓰고 순위를 다시 매긴다."""
    sc = score(co, mctx, today)
    by = {o["key"]: o for o in options}
    for name, (gate, total, parts) in sc.items():
        o = by.get(KEY[name])
        if not o:
            continue
        if KEY[name] == "hybrid":
            o["title"] = "영구채 (신종자본증권)"
        if KEY[name] == "rights" and co.get("market") not in ("K", "Y", "N"):
            o["title"] = "제3자배정 증자 (RCPS·CPS)"
        o["base"] = 0
        o["factors"] = [{"pts": p, "text": f"[{cat}] {lab}: {val}"} for cat, p, lab, val in parts]
        o["reasons"] = [f"{lab}: {val}" for _, p, lab, val in parts if p > 0][:5]
        o["cautions"] = [f"{lab}: {val}" for _, p, lab, val in parts if p <= 0][:4]
        if gate:
            o["score"], o["fit"], o["gate"] = 0, "부적합", gate
            o["cautions"] = [gate] + o["cautions"]
        else:
            o["score"] = max(0, min(100, round(total)))
            o["fit"] = "높음" if o["score"] >= 65 else "보통" if o["score"] >= 45 else "낮음"
    sz = sizes(co, today)
    for o in options:
        if o["key"] in sz:
            o["size"] = sz[o["key"]]
        if o["key"] == "public_bond" and sz.get("_pub_short") and not o.get("gate"):
            o["score"] = max(0, o["score"] - 15)
            o["factors"].append({"pts": -15, "text": "[조건] 공모 최소 규모: 조달 필요액이 300억 미만"})
            o["cautions"].append("조달 필요액이 공모 최소 규모(300억)에 못 미침")
            o["fit"] = "높음" if o["score"] >= 65 else "보통" if o["score"] >= 45 else "낮음"
        if o["key"] == "hybrid" and sz.get("_hy_small") and not o.get("gate"):
            o["score"] = max(0, o["score"] - 15)
            o["factors"].append({"pts": -15, "text": "[재무] 부채비율 100% 미만: 자본 인정 필요성 낮음"})
            o["cautions"].append("부채비율 100% 미만 — 자본 인정 효과가 크게 필요하지 않은 구간")
            o["fit"] = "높음" if o["score"] >= 65 else "보통" if o["score"] >= 45 else "낮음"
    profitable = (co.get("metrics") or {}).get("interest_coverage") or 0
    options.sort(key=lambda o: (-o["score"], TIE.index(o["key"]) if o["key"] in TIE and profitable >= 3 else 0))
    return options
