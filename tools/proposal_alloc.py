"""만기별 발행금액 자동 배분 (팝업에서 '미정' 체크 시).

만기마다 네 점수를 0~1로 만든 뒤 가중합(균형)으로 나눈다.
  수요  : 예상 주문 (동사 직전 회차 해당 만기 주문, 같은 등급군 6개월 트랜치당 평균 주문)
  비용  : 예상 금리가 낮을수록 높음 (기준 민평 + 목표 가산금리)
  만기  : 그 해 동사 공모채 만기가 적을수록 높음
  전망  : 리서치가 금리 상승을 보면 긴 만기, 하락을 보면 짧은 만기에 가산
그 다음 만기별 상한(예상 주문 ÷ 목표 경쟁률)과 최소 금액을 맞추고 50억 단위로 반올림한다.
기준값은 docs/proposal_spec.md 의 rules.alloc 에서 바꾼다.
"""
from datetime import date, timedelta

from tools import proposal_mkt as M

DEFAULT = {"unit": 50, "min_tranche": 300, "target_ratio": 2.0,
           "weights": {"demand": 0.35, "cost": 0.2, "maturity": 0.3, "outlook": 0.15}}


def _norm(d):
    s = sum(v for v in d.values() if v is not None and v > 0)
    return {k: ((v or 0) / s if s else 1 / len(d)) for k, v in d.items()}


def allocate(tenors, base, grade, pubs, prev, lad, issue, today, curves=None, outlook=None, adj=None, rules=None):
    R = {**DEFAULT, **(rules or {})}
    W = {**DEFAULT["weights"], **(R.get("weights") or {})}
    unit, mn, tr = int(R["unit"]), int(R["min_tranche"]), float(R["target_ratio"])
    tenors = sorted(set(int(t) for t in tenors))
    dem, _ = M.demand_stats(pubs, grade, today, 6)
    pdm = {int(round(r["_tenor"])): r for r in prev if r.get("_tenor") and r.get("_dem")}
    # 수요: 트랜치당 예상 주문(억원)
    exp_orders = {}
    for t in tenors:
        xs = []
        if t in pdm:
            xs.append(pdm[t]["_dem"]["total"])
        b = dem.get(M.tbucket(t))
        if b and b["n"]:
            xs.append(b["orders"] / b["n"])
        exp_orders[t] = sum(xs) / len(xs) if xs else None
    known = [v for v in exp_orders.values() if v]
    fill = (min(known) * 0.7) if known else base / len(tenors) * tr  # 자료 없는 만기는 보수적으로
    exp_orders = {t: (v if v else fill) for t, v in exp_orders.items()}
    s_dem = _norm(exp_orders)
    # 비용: 예상 금리 (기준 민평, 없으면 국고+등급 스프레드)
    rate = {}
    for t in tenors:
        y = curves.y(grade, t) if curves else None
        rate[t] = (y + (adj["bp"] / 100 if adj else 0)) if y is not None else None
    if all(v is not None for v in rate.values()):
        lo = min(rate.values())
        s_cost = _norm({t: 1 / (1 + (rate[t] - lo) * 100 / 50) for t in tenors})
    else:
        s_cost = {t: 1 / len(tenors) for t in tenors}
    # 만기: 그 해 기존 만기가 적을수록
    yrs = {t: lad.get(issue.year + t, 0) for t in tenors}
    avg = (sum(yrs.values()) / len(yrs)) or 1
    s_mat = _norm({t: 1 / (1 + yrs[t] / avg) for t in tenors})
    # 전망: 1년 뒤 동사 등급 3년 금리 방향
    direction = 0
    if outlook and curves:
        from tools import proposal_outlook as OL
        g3, _ = OL.grade_path(outlook, grade, 3, curves)
        _, v1 = OL.at(outlook, g3, issue + timedelta(days=365))
        now = curves.y(grade, 3)
        if v1 is not None and now is not None:
            d = (v1 - now) * 100
            direction = 1 if d > 10 else -1 if d < -10 else 0
    if direction > 0:
        s_out = _norm({t: t for t in tenors})
    elif direction < 0:
        s_out = _norm({t: 1 / t for t in tenors})
    else:
        s_out = {t: 1 / len(tenors) for t in tenors}
    score = {t: W["demand"] * s_dem[t] + W["cost"] * s_cost[t] + W["maturity"] * s_mat[t] + W["outlook"] * s_out[t] for t in tenors}
    sw = sum(score.values())
    raw = {t: base * score[t] / sw for t in tenors}
    cap = {t: exp_orders[t] / tr for t in tenors}
    # 상한 초과분은 여유 있는 만기로 넘김 (전부 꽉 차면 그대로 둠)
    for _ in range(5):
        over = sum(max(0, raw[t] - cap[t]) for t in tenors)
        room = {t: cap[t] - raw[t] for t in tenors if raw[t] < cap[t]}
        if over <= 1 or not room:
            break
        rs = sum(room.values())
        for t in tenors:
            if raw[t] > cap[t]:
                raw[t] = cap[t]
            elif t in room:
                raw[t] += over * room[t] / rs
    # 최소 금액: 못 맞추는 만기는 빼지 않고 최소로 올린 뒤 나머지에서 줄임
    if base >= mn * len(tenors):
        short = {t: mn - raw[t] for t in tenors if raw[t] < mn}
        if short:
            need = sum(short.values())
            big = {t: raw[t] - mn for t in tenors if t not in short and raw[t] > mn}
            bs = sum(big.values()) or 1
            for t in tenors:
                raw[t] = mn if t in short else raw[t] - need * big.get(t, 0) / bs
    amt = {t: max(unit, int(round(raw[t] / unit)) * unit) for t in tenors}
    diff = base - sum(amt.values())
    k = max(tenors, key=lambda t: amt[t])
    amt[k] += diff
    prio = max(tenors, key=lambda t: (cap[t] - amt[t], t))
    detail = {t: {"orders": round(exp_orders[t]), "cap": round(cap[t]), "rate": rate[t], "year_mat": round(yrs[t]),
                  "s_demand": s_dem[t], "s_cost": s_cost[t], "s_mat": s_mat[t], "s_out": s_out[t], "score": score[t] / sw}
              for t in tenors}
    return {"split": amt, "priority": prio, "detail": detail, "direction": direction, "weights": W, "target_ratio": tr}


def split_str(split):
    return ",".join(f"{t}:{v}" for t, v in sorted(split.items()))
