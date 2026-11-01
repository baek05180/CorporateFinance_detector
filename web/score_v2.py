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
    out = {}

    # ── 공모채 ──
    P = []
    gate = "등급 없음" if gi is None else ("자본잠식" if impaired else None)
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
    tp = H.get("total_pct")
    if tp is not None:
        after = tp / 1.1
        P.append(("재무", 15 if after >= 30 else 10 if after >= 20 else 5 if after >= 15 else 0, "희석 후 최대주주", f"{tp:.1f}%→{after:.1f}% (시총 10% 발행 가정)"))
    link_bal = sum(d.get("amount") or 0 for d in deals if d.get("type") in ("CB", "BW", "EB")
                   and (d.get("maturity") or "9") > today.isoformat() and _years_since(d.get("board_dt") or d.get("rcept_dt"), today) <= 3)
    if mcap:
        oh = link_bal / mcap * 100
        P.append(("재무", 15 if oh < 5 else 8 if oh < 15 else 0, "최근 3년 메자닌 발행(미만기)", f"시총의 {oh:.0f}%"))
    P.append(("재무", 10 if loss or (m.get("op_margin") or 0) < 3 else 5, "수익성", "저수익·적자 → 저쿠폰 CB 유리" if loss or (m.get("op_margin") or 0) < 3 else "흑자"))
    P.append(("시장", {"K": 20, "N": 10, "Y": 10}.get(mk, 0), "상장 시장", {"K": "코스닥", "N": "코넥스", "Y": "유가증권"}.get(mk, "비상장")))
    ph = co.get("price_hist") or []
    if len(ph) > 20:
        rets = [math.log(ph[i][1] / ph[i - 1][1]) for i in range(1, len(ph)) if ph[i - 1][1]]
        vol = (sum(r * r for r in rets) / len(rets)) ** .5 * math.sqrt(250) * 100
        if vol < 3:
            P.append(("시장", -20, "주가 변동성(연)", f"{vol:.0f}% — 거래정지·관리종목 여부 확인"))
        else:
            P.append(("시장", 10 if 30 <= vol <= 80 else 5, "주가 변동성(연)", f"{vol:.0f}%"))
    if mcap:
        P.append(("조건", 10 if mcap >= 5e10 else 5 if mcap >= 2e10 else 0, "시가총액", f"{mcap/1e8:,.0f}억"))
    if dr is not None:
        P.append(("조건", 10 if dr > 200 else 5, "부채비율", f"{dr:.0f}% {'→ 전환 시 자본 확충 효과' if dr > 200 else ''}".strip()))
    if gi is not None and gi <= 6:
        P.append(("조건", -25, "신용등급", f"{grade} — 일반 사채로 희석 없이 조달 가능"))
    n = cnt3("CB")
    P.append(("이력", 10 if 1 <= n <= 2 else 3 if n >= 3 else 5, "최근 3년 CB", f"{n}건" + (" (누적 오버행)" if n >= 3 else "")))
    out["CB"] = (gate, _sum(P), P)

    # ── EB ──
    P = []
    ts, sh = co.get("treasury_shares"), m.get("shares")
    gate = None if ts else "자기주식 없음"
    if ts and sh:
        r = ts / sh * 100
        P.append(("재무", 40 if r >= 5 else 30 if r >= 3 else 15, "자기주식 비율", f"{r:.1f}%"))
    P.append(("시장", 20 if mk in ("K", "Y") else 5, "상장 시장", mk or "-"))
    if mcap:
        P.append(("시장", 10 if mcap >= 1e11 else 5, "시가총액", f"{mcap/1e8:,.0f}억"))
    P.append(("조건", 10, "규제", "자사주 EB 공시 강화 — 주주 설명 필요"))
    if cnt3("EB"):
        P.append(("이력", 10, "최근 3년 EB", f"{cnt3('EB')}건"))
    out["EB"] = (gate, _sum(P), P)

    # ── 유상증자 ──
    P = []
    gate = None if mk in ("K", "Y", "N") else "비상장(제3자배정·RCPS 별도 검토)"
    f = 0
    if impaired:
        f += 30
    if dr is not None and dr > 300:
        f += 15
    if (m.get("loss_years") or 0) >= 2:
        f += 10
    P.append(("재무", min(40, f), "자본 확충 필요", ("자본잠식 " if impaired else "") + (f"부채비율 {dr:.0f}% " if dr else "") + (f"{m.get('loss_years')}년 연속 적자" if (m.get('loss_years') or 0) >= 2 else "")))
    if tp is not None:
        P.append(("시장", 15 if tp >= 30 else 8 if tp >= 20 else 3, "최대주주 지분(참여 여력)", f"{tp:.1f}%"))
    if len(ph) > 20:
        dd = (ph[-1][1] / max(p[1] for p in ph) - 1) * 100
        P.append(("시장", 5 if dd < -30 else 15, "6개월 고점 대비", f"{dd:.0f}%"))
    if mcap:
        P.append(("조건", 10 if mcap >= 3e10 else 5, "시가총액", f"{mcap/1e8:,.0f}억"))
    lr = last("RIGHTS")
    yrs = _years_since(lr, today) if lr else 99
    P.append(("이력", -20 if yrs < 1 else -10 if yrs < 2 else 10, "최근 유상증자", f"{lr[:10]} ({yrs:.1f}년 전)" if lr else "없음"))
    out["유상증자"] = (gate, _sum(P), P)
    return out
