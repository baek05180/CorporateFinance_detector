"""기업 화면·조달 제안 보고서용 계산.

모든 판단은 공시 숫자에 근거한 규칙(rule)으로 하며, 근거 문장을 함께 남긴다.
점수는 '이 회사에 먼저 이야기해 볼 만한 순서'를 정하려는 것이지 투자 판단이 아니다.
"""
import math
import re
from datetime import date, timedelta

EOK = 100_000_000

GRADES = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-",
          "BB+", "BB", "BB-", "B+", "B", "B-", "CCC", "CC", "C", "D"]


def grade_of(text):
    """'AA-(한국신용평가)', 'A0', 'BBB+/안정적' → 'AA-' / 'A' / 'BBB+'"""
    if not text:
        return None
    m = re.search(r"\b(AAA|AA|A|BBB|BB|B|CCC|CC|C|D)\s*([+\-0])?", str(text).upper())
    if not m:
        return None
    g = m.group(1) + ("" if m.group(2) in (None, "0") else m.group(2))
    return g if g in GRADES else None


def grade_rank(g):
    return GRADES.index(g) if g in GRADES else None


def eok(v):
    if v is None:
        return "-"
    e = v / EOK
    if abs(e) >= 10000:
        return f"{e / 10000:,.1f}조"
    return f"{e:,.0f}억" if abs(e) >= 10 else f"{e:,.1f}억"


def pct(a, b, digits=1):
    if a is None or not b:
        return None
    return round(a / b * 100, digits)


def safe_div(a, b, digits=2):
    if a is None or b in (None, 0):
        return None
    return round(a / b, digits)


# ── 재무 시계열·지표 ──
def financial_series(multi_row, annual_year, interim_row, interim_period, prof):
    """연간 3개년 + 최근 분기 누적 → [{"period","revenue","op","ni","assets","liab","equity",...}]"""
    out = []
    if multi_row:
        for i, slot in enumerate(("prev2", "prev", "cur")):
            vals = multi_row.get(slot) or {}
            if vals:
                out.append({"period": f"{annual_year - (2 - i)}", "kind": "annual", "fs": multi_row.get("fs"), **vals})
    elif prof and prof.get("fs_annual"):
        fa = prof["fs_annual"]
        out.append({"period": fa["period"].split()[0], "kind": "annual", "fs": fa.get("fs"),
                    **{k: fa.get(k) for k in ("revenue", "op", "ni", "assets", "liab", "equity", "capital")}})
    if interim_row and interim_row.get("cur"):
        out.append({"period": interim_period, "kind": "interim", "fs": interim_row.get("fs"), **interim_row["cur"]})
        if interim_row.get("prev"):
            # 직전 연도 같은 분기 (손익만 의미 있음)
            prev = {k: interim_row["prev"].get(k) for k in ("revenue", "op", "ni")}
            y, q = interim_period.split()
            out.append({"period": f"{int(y) - 1} {q}", "kind": "interim_prev", **prev})
    for r in out:
        r["op_margin"] = pct(r.get("op"), r.get("revenue"))
        r["net_margin"] = pct(r.get("ni"), r.get("revenue"))
        r["debt_ratio"] = pct(r.get("liab"), r.get("equity")) if (r.get("equity") or 0) > 0 else None
        r["current_ratio"] = pct(r.get("cur_assets"), r.get("cur_liab"))
    return out


def key_metrics(series, prof, price):
    annual = [r for r in series if r["kind"] == "annual"]
    last = annual[-1] if annual else {}
    prev = annual[-2] if len(annual) >= 2 else {}
    bs = next((r for r in reversed(series) if r["kind"] in ("annual", "interim") and r.get("equity") is not None), last)
    fd = (prof or {}).get("fs_interim") or (prof or {}).get("fs_annual") or {}
    fa = (prof or {}).get("fs_annual") or {}
    m = {
        "period": last.get("period"),
        "revenue": last.get("revenue"), "op": last.get("op"), "ni": last.get("ni"),
        "revenue_growth": pct((last.get("revenue") or 0) - (prev.get("revenue") or 0), prev.get("revenue")) if prev.get("revenue") else None,
        "op_margin": last.get("op_margin"),
        "loss_years": sum(1 for r in annual if (r.get("op") or 0) < 0),
        "bs_period": bs.get("period"),
        "assets": bs.get("assets"), "liab": bs.get("liab"), "equity": bs.get("equity"), "capital": bs.get("capital"),
        "debt_ratio": bs.get("debt_ratio"), "current_ratio": bs.get("current_ratio"),
        "cash": fd.get("cash"), "st_deposits": fd.get("st_deposits"), "debt": fd.get("debt"),
        "detail_period": fd.get("period"),
    }
    cash_like = (m["cash"] or 0) + (m["st_deposits"] or 0)
    m["net_debt"] = (m["debt"] - cash_like) if m["debt"] is not None else None
    ebitda = None
    if fa.get("op") is not None:
        ebitda = fa["op"] + (fa.get("dep") or 0) + (fa.get("amort") or 0)
    m["ebitda"] = ebitda
    m["ebitda_ok"] = bool(fa.get("dep") or fa.get("amort"))  # 감가상각비를 못 받았으면 EBITDA = 영업이익이라 과소
    m["net_debt_ebitda"] = safe_div(m["net_debt"], ebitda, 1) if ebitda and ebitda > 0 else None
    # 이자보상배율: 현금흐름표의 이자 지급액(없으면 손익계산서 이자비용, 그것도 없으면 금융원가)
    cost = fa.get("interest_paid") or fa.get("interest") or fa.get("fin_cost")
    m["interest_basis"] = "이자 지급액" if fa.get("interest_paid") else ("이자비용" if fa.get("interest") else ("금융원가" if fa.get("fin_cost") else None))
    m["interest_coverage"] = safe_div(fa.get("op"), cost, 1) if cost else None
    m["ocf"], m["capex"] = fa.get("ocf"), fa.get("capex")
    if m["equity"] is not None and m["capital"]:
        m["impairment"] = "완전잠식" if m["equity"] <= 0 else ("부분잠식" if m["equity"] < m["capital"] else None)
    if price:
        m["close"], m["mcap"], m["shares"] = price.get("close"), price.get("mcap"), price.get("shares")
        if m.get("ni") and m["ni"] > 0:
            m["per"] = safe_div(price["mcap"], m["ni"], 1)
        if m.get("equity") and m["equity"] > 0:
            m["pbr"] = safe_div(price["mcap"], m["equity"], 2)
    return {k: v for k, v in m.items() if v is not None}


# ── peer ──
def peers(code, listed, multi, prices, limit=7):
    me = listed.get(code)
    if not me:
        return []
    ind = me.get("industry")
    mine = (multi.get(code) or {}).get("cur", {}).get("revenue")
    rows = []
    for c, v in listed.items():
        if v.get("industry") != ind:
            continue
        cur = (multi.get(c) or {}).get("cur") or {}
        rev = cur.get("revenue")
        if c != code and not rev:
            continue
        dist = abs(math.log(max(rev, 1) / max(mine, 1))) if (rev and mine) else 99
        p = prices.get(c) or {}
        eq, ni = cur.get("equity"), cur.get("ni")
        rows.append({
            "code": c, "name": v["name"], "market": v.get("market"), "me": c == code, "dist": 0 if c == code else dist,
            "revenue": rev, "op_margin": pct(cur.get("op"), rev), "debt_ratio": pct(cur.get("liab"), eq) if (eq or 0) > 0 else None,
            "mcap": p.get("mcap"), "per": safe_div(p.get("mcap"), ni, 1) if (ni or 0) > 0 and p.get("mcap") else None,
            "pbr": safe_div(p.get("mcap"), eq, 2) if (eq or 0) > 0 and p.get("mcap") else None,
        })
    rows.sort(key=lambda r: r["dist"])
    return [{k: v for k, v in r.items() if v is not None and k != "dist"} for r in rows[:limit + 1]]


# ── 잔존 사채 ──
def bond_profile(bonds, today):
    if not bonds:
        return {}
    in12 = (today + timedelta(days=365)).isoformat()
    t = today.isoformat()
    live = [b for b in bonds if (b.get("mat") or "9999") >= t and (b.get("bal") or b.get("amt"))]

    def bucket(b):
        if b.get("link"):
            return "주식연계"
        if b.get("method") == "공모":
            return "공모"
        return "사모"   # 사모 + CBO기초사모
    ladder = {}
    for b in live:
        y = (b.get("mat") or "")[:4]
        ladder.setdefault(y, {"공모": 0, "사모": 0, "주식연계": 0})[bucket(b)] += (b.get("bal") or b.get("amt") or 0)
    due = [b for b in live if (b.get("mat") or "9999") <= in12]
    total = sum(b.get("bal") or b.get("amt") or 0 for b in live)
    due_sum = {k: sum(b.get("bal") or b.get("amt") or 0 for b in due if bucket(b) == k) for k in ("공모", "사모", "주식연계")}
    return {
        "count": len(live), "total": total,
        "due12": sum(due_sum.values()), "due12_by": due_sum,
        "ladder": [{"year": y, **v} for y, v in sorted(ladder.items())],
        "avg_coupon": round(sum((b.get("cpn") or 0) * (b.get("bal") or b.get("amt") or 0) for b in live if b.get("cpn")) /
                            max(1, sum((b.get("bal") or b.get("amt") or 0) for b in live if b.get("cpn"))), 2) if live else None,
        "has_cbo": any("CBO" in (b.get("method") or "") for b in live),
        "list": sorted(live, key=lambda b: b.get("mat") or ""),
    }


# ── 조달 이력 요약 ──
TYPE_LABEL = {"CB": "CB", "BW": "BW", "EB": "EB", "RIGHTS": "유상증자", "BOND": "회사채"}


def deal_label(r):
    if r["type"] == "BOND":
        if r.get("pcbo"):
            return "P-CBO"
        return f"회사채 {r.get('offering') or ''}".strip()
    return TYPE_LABEL.get(r["type"], r["type"])


def deal_date(r):
    return r.get("board_dt") or r.get("sub_dt") or r.get("issue_dt") or r.get("rcept_dt") or ""


def history_summary(deals, today):
    since = (today - timedelta(days=1095)).isoformat()
    recent = [d for d in deals if deal_date(d) >= since]
    by = {}
    for d in recent:
        k = deal_label(d)
        x = by.setdefault(k, {"label": k, "count": 0, "amount": 0, "last": ""})
        x["count"] += 1
        x["amount"] += d.get("amount") or 0
        x["last"] = max(x["last"], deal_date(d))
    return sorted(by.values(), key=lambda x: -x["amount"])


# ── 조달 방식 제안 ──
def _fit(score):
    return "높음" if score >= 65 else ("보통" if score >= 45 else "낮음")


def _round_size(v, unit=50 * EOK):
    if not v or v <= 0:
        return None
    return max(unit, round(v / unit) * unit)


def propose(ctx):
    """ctx: market, metrics, rating, history(list of deals), bonds(bond_profile), holders, treasury_shares, industry
    → [{"key","title","score","fit","reasons":[...],"cautions":[...],"size":str}]"""
    mk = ctx.get("market") or "E"            # Y 유가, K 코스닥, N 코넥스, E 비상장
    m = ctx.get("metrics") or {}
    deals = ctx.get("deals") or []
    bonds = ctx.get("bonds") or {}
    holders = ctx.get("holders") or {}
    g = ctx.get("rating")
    gr = grade_rank(g)
    rev = m.get("revenue")
    assets = m.get("assets")
    op = m.get("op")
    dr = m.get("debt_ratio")
    loss = op is not None and op < 0
    loss_years = m.get("loss_years", 0)
    impaired = m.get("impairment")
    mcap = m.get("mcap")
    has_fin = rev is not None or assets is not None
    cnt = lambda *types: sum(1 for d in deals if d["type"] in types)  # noqa: E731
    cnt_bond = lambda off: sum(1 for d in deals if d["type"] == "BOND" and d.get("offering") == off and not d.get("pcbo"))  # noqa: E731
    cbo_hist = sum(1 for d in deals if d.get("pcbo"))
    due_pub = (bonds.get("due12_by") or {}).get("공모", 0)
    due_pvt = (bonds.get("due12_by") or {}).get("사모", 0)
    due_link = (bonds.get("due12_by") or {}).get("주식연계", 0)
    big = (rev or 0) >= 5000 * EOK or (assets or 0) >= 10000 * EOK
    small = has_fin and (rev or 0) < 1500 * EOK and (assets or 0) < 3000 * EOK
    opts = []

    def add(key, title, base, reasons, cautions, size=None, note=None):
        """reasons·cautions: [(가감점, 문장)] → 점수 = 기본점수 + 가감점 합 (0~100으로 자름)"""
        factors = [{"pts": p, "text": t} for p, t in reasons + cautions if t]
        score = base + sum(f["pts"] for f in factors)
        opts.append({"key": key, "title": title, "base": base, "factors": factors,
                     "score": max(0, min(100, round(score))), "fit": _fit(score),
                     "reasons": [t for p, t in reasons if t], "cautions": [t for p, t in cautions if t],
                     "size": size, "note": note})

    # 1) 공모 회사채
    base, R, C = 20, [], []
    if gr is not None and gr <= grade_rank("A-"):
        R.append((40, f"최근 신용등급 {g} — 공모채 수요예측에 참여하는 기관 수요가 두터운 구간"))
    elif gr is not None and gr <= grade_rank("BBB-"):
        R.append((10, f"신용등급 {g}")); C.append((0, "BBB급은 수요예측 미매각 가능성이 있어 하이일드 펀드·리테일 수요 확인 필요"))
    elif gr is None:
        C.append((-10, "유효 신용등급이 확인되지 않음 — 공모 발행에는 신용평가 2곳 이상 필요"))
    if cnt_bond("공모"):
        R.append((15, f"최근 3년 공모 회사채 {cnt_bond('공모')}건 발행 이력 — 시장 인지도 확보"))
    if due_pub:
        R.append((15, f"12개월 내 공모채 만기 {eok(due_pub)} — 차환 발행 수요"))
    if big:
        R.append((15, f"매출 {eok(rev)} · 자산 {eok(assets)} 규모로 공모 최소 발행 규모(수백억) 소화 가능"))
    elif small:
        C.append((-20, "매출·자산 규모가 작아 공모 비용(평가·주관·공시) 대비 효율이 낮음"))
    if loss:
        C.append((-15, "최근 연도 영업손실 — 등급 하향 가능성 점검"))
    if dr and dr > 400:
        C.append((-10, f"부채비율 {dr:.0f}%"))
    size = _round_size(due_pub) or _round_size(_median([d.get("amount") for d in deals if d["type"] == "BOND" and d.get("offering") == "공모"]))
    add("public_bond", "공모 회사채", base, R, C, eok(size) if size else "500억~1,000억 (수요예측 결과에 따라 증액)" if big else None,
        "수요예측 → 증권신고서 효력 발생 → 청약·납입. 금리는 민평 대비 가산금리로 결정")

    # 2) 사모 회사채
    base, R, C = 25, [], []
    if gr is None or gr >= grade_rank("BBB+"):
        R.append((15, "등급이 없거나 BBB급 이하 — 공모보다 사모가 현실적인 구간" if gr is None else f"신용등급 {g} — 공모 수요가 얇은 구간이라 사모 선호"))
    if cnt_bond("사모"):
        R.append((10, f"최근 3년 사모채 {cnt_bond('사모')}건 발행 — 기존 투자자 네트워크 활용 가능"))
    if due_pvt:
        R.append((15, f"12개월 내 사모채 만기 {eok(due_pvt)} — 차환 수요"))
    if has_fin and not loss:
        R.append((10, "영업이익 흑자 — 이자 상환 능력 설명 가능"))
    if loss:
        C.append((-10, "영업손실 — 담보·보증 또는 조기상환(Put) 조건 요구 가능성"))
    if impaired:
        C.append((-20, f"자본{impaired} 상태 — 사채보다 자본 확충이 먼저"))
    if big and gr is not None and gr <= grade_rank("A-"):
        C.append((-15, "우량 등급이면 공모채가 금리 면에서 유리"))
    size = _round_size(due_pvt) or _round_size(_median([d.get("amount") for d in deals if d["type"] == "BOND" and d.get("offering") == "사모" and not d.get("pcbo")]))
    add("private_bond", "사모 회사채", base, R, C, eok(size) if size else "100억~300억",
        "49인 이하 투자자 대상, 증권신고서 불필요. 강제상환·신용등급 하락 시 조기상환 등 특약이 흔함")

    # 3) P-CBO
    base, R, C = 10, [], []
    if cbo_hist or bonds.get("has_cbo"):
        R.append((25, "P-CBO 편입 이력 보유 — 신용보증기금 심사 경험"))
    if small or (mk in ("K", "N", "E") and not big):
        R.append((25, "중소·중견 규모 — 신용보증기금 P-CBO 지원 대상에 해당할 가능성"))
    if gr is not None and grade_rank("BBB+") <= gr <= grade_rank("B-"):
        R.append((10, f"신용등급 {g} — 단독 사모 발행보다 보증부 P-CBO가 금리 면에서 유리할 수 있음"))
    if impaired:
        C.append((-30, "자본잠식 기업은 보증 심사 통과가 어려움"))
    if big:
        C.append((-30, "대기업·우량 기업은 P-CBO 지원 대상이 아님"))
    if loss_years >= 2:
        C.append((-10, "2년 이상 영업손실 — 보증 한도 축소 가능성"))
    add("pcbo", "P-CBO (신용보증기금 유동화 보증)", base, R, C, "수십억~100억 (보증 한도 내)",
        "기업이 사모사채를 발행하면 신보가 보증한 유동화증권(P-CBO)에 편입. 정기 차수(대개 월 단위)에 맞춰 신청")

    # 4) CB
    base, R, C = 10, [], []
    if mk in ("K", "N"):
        R.append((25, f"{'코스닥' if mk == 'K' else '코넥스'} 상장사 — 메자닌 투자자 풀이 가장 큰 시장"))
    elif mk == "Y":
        R.append((5, "유가증권 상장사 — CB 발행은 가능하나 코스닥보다 투자자 풀이 작음"))
    else:
        C.append((-20, "비상장사 CB는 투자자가 제한적 — 상장 예정이면 RCPS·CB 병행 검토"))
    if loss or (m.get("op_margin") is not None and m["op_margin"] < 3):
        R.append((10, "수익성이 낮아 일반 사채 금리 부담이 큼 — 표면금리 0~2% CB가 이자 부담을 줄임"))
    if dr and dr > 200:
        R.append((10, f"부채비율 {dr:.0f}% — 전환 시 자본으로 바뀌는 구조가 재무구조 개선에 유리"))
    if cnt("CB"):
        R.append((15, f"최근 3년 CB {cnt('CB')}건 발행 — 기존 메자닌 투자자 존재"))
    if cnt("CB") >= 3:
        C.append((-10, f"CB를 {cnt('CB')}차례 발행 — 누적 희석·오버행 우려로 조건(리픽싱 하한, 콜옵션)이 불리해질 수 있음"))
    if due_link:
        R.append((5, f"12개월 내 주식연계채권 만기·상환 {eok(due_link)} — 차환 CB 수요"))
    if gr is not None and gr <= grade_rank("A-"):
        C.append((-15, "우량 등급이면 일반 사채가 희석 없이 조달 가능"))
    if mcap and mcap < 300 * EOK:
        C.append((-10, f"시가총액 {eok(mcap)} — 발행 규모가 시총 대비 커지면 희석 부담"))
    cb_size = _round_size(mcap * 0.1, 10 * EOK) if mcap else _round_size(_median([d.get("amount") for d in deals if d["type"] == "CB"]), 10 * EOK)
    add("cb", "전환사채 (CB)", base, R, C, (eok(cb_size) + (" (시총의 약 10%)" if mcap else "")) if cb_size else None,
        "리픽싱 하한 70%, 발행 1년 후 Put, 최대주주 Call(통상 30% 이내)이 일반적인 조건")

    # 5) EB
    base, R, C = 0, [], []
    ts = ctx.get("treasury_shares") or 0
    shares = m.get("shares")
    if ts > 0:
        ratio = f" (발행주식의 {ts / shares * 100:.1f}%)" if shares else ""
        val = f", 시가 약 {eok(ts * m['close'])}" if m.get("close") else ""
        R.append((45, f"자기주식 {ts:,}주 보유{ratio}{val} — 교환 대상 주식으로 활용 가능"))
        if shares and ts / shares >= 0.03:
            R.append((10, "자기주식이 발행주식의 3% 이상 — 의미 있는 규모의 EB 가능"))
    else:
        C.append((0, "자기주식 보유가 확인되지 않음 — 계열사 지분 등 다른 교환 대상이 있는지 확인"))
    if cnt("EB"):
        R.append((10, f"EB 발행 이력 {cnt('EB')}건"))
    if ts > 0:
        C.append((0, "자기주식을 대상으로 한 EB는 공시·규제가 강화되는 추세 — 처분 목적·주주 설명 필요"))
    eb_size = None
    if ts and m.get("close"):
        eb_size = ts * m["close"] * 0.9
        if mcap:
            eb_size = min(eb_size, mcap * 0.3)
    add("eb", "교환사채 (EB)", base, R, C, eok(_round_size(eb_size, 10 * EOK)) if eb_size else None,
        "신주 발행 없이 보유 주식으로 교환하므로 추가 희석이 없음")

    # 6) 유상증자
    base, R, C = 10, [], []
    if impaired:
        R.append((40, f"자본{impaired} — 자본 확충이 우선 과제 (관리종목·상장적격성 이슈 예방)"))
    if dr and dr > 300:
        R.append((20, f"부채비율 {dr:.0f}% — 차입 여력보다 자본 조달이 필요"))
    if loss_years >= 2:
        R.append((10, "2년 이상 영업손실 — 부채성 조달의 상환 부담이 큼"))
    tp = holders.get("total_pct")
    if tp is not None:
        if tp >= 40:
            R.append((5, f"최대주주 등 지분 {tp:.1f}% — 주주배정 시 최대주주 참여로 경영권 유지 가능"))
        elif tp < 20:
            C.append((0, f"최대주주 등 지분 {tp:.1f}% — 대규모 증자 시 경영권 희석 주의"))
    if cnt("RIGHTS"):
        R.append((0, f"최근 3년 유상증자 {cnt('RIGHTS')}건"))
    if mk == "E":
        C.append((0, "비상장사는 제3자배정(전략적 투자자·VC) 또는 RCPS 형태가 현실적"))
    method = ("주주배정 후 실권주 일반공모" if mk in ("Y", "K") and (mcap or 0) >= 1000 * EOK
              else "제3자배정 (전략적·재무적 투자자)" if mk in ("K", "N", "E") else "주주배정")
    add("rights", f"유상증자 ({method})", base, R, C, eok(_round_size((mcap or 0) * 0.15, 10 * EOK)) + " (시총의 약 15%)" if mcap else None,
        "증권신고서 제출(공모) → 권리락 → 청약. 할인율·발행가 산정 방식 확인")

    # 7) 자본성 증권 (영구채·RCPS)
    base, R, C = 5, [], []
    if mk == "Y" and dr and dr > 200 and (gr is None or gr <= grade_rank("BBB-")):
        R.append((35, f"부채비율 {dr:.0f}% — 신종자본증권(영구채)은 회계상 자본으로 인정돼 부채비율을 낮춤"))
        if gr is not None:
            R.append((0, f"신용등급 {g} — 영구채 발행이 가능한 등급대"))
    if mk == "E" and has_fin:
        R.append((30, "비상장 성장 기업 — 상장 전 RCPS(상환전환우선주)로 지분 투자 유치"))
    if mk in ("K", "N") and loss and not impaired:
        R.append((10, "적자 상장사 — 제3자배정 RCPS·CPS로 FI 유치 가능"))
    C.append((0, "영구채는 금리가 일반 사채보다 높고 콜(통상 5년) 이후 금리 상향(step-up) 조건 확인"))
    add("hybrid", "자본성 증권 (영구채·RCPS)", base, R, C, None)

    opts.sort(key=lambda o: -o["score"])
    return opts


def _median(vals):
    v = sorted(x for x in vals if x)
    if not v:
        return None
    return v[len(v) // 2]


def summary_line(name, industry, metrics, hist, bonds, top):
    parts = [f"{name}은(는) {industry} 업종" if industry else name]
    if metrics.get("revenue") is not None:
        parts.append(f"{metrics.get('period')}년 매출 {eok(metrics['revenue'])}, 영업이익 {eok(metrics.get('op'))}")
    if metrics.get("debt_ratio") is not None:
        parts.append(f"부채비율 {metrics['debt_ratio']:.0f}%")
    s = ", ".join(parts) + "."
    if hist:
        s += " 최근 3년 조달은 " + ", ".join(f"{h['label']} {h['count']}건({eok(h['amount'])})" for h in hist[:4]) + "."
    if bonds.get("due12"):
        s += f" 12개월 안에 만기가 오는 사채가 {eok(bonds['due12'])} 있습니다."
    if top:
        s += " 우선 검토할 방식은 " + " · ".join(o["title"] for o in top[:2]) + "입니다."
    return s


# ── 비교 사례 ──
def pick_comps(opt_key, me, pool, limit=5):
    """me: {"key","industry","market","revenue","grade"} / pool: 최근 1년 같은 종류 조달 [(row, meta)]
    meta: {"industry","market","revenue","grade"}. 회사당 1건, 고른 근거를 함께 돌려준다."""
    out, seen = [], {me["key"]}

    def take(cands, basis):
        for row, meta in cands:
            if len(out) >= limit:
                return
            if row["key"] in seen:
                continue
            seen.add(row["key"])
            out.append((row, basis(meta) if callable(basis) else basis))

    if opt_key == "public_bond":
        gr = grade_rank(me.get("grade"))
        if gr is not None:
            near = [(r, m) for r, m in pool if m.get("grade") is not None and abs(grade_rank(m["grade"]) - gr) <= 1]
            take([(r, m) for r, m in near if me.get("industry") and m.get("industry") == me["industry"]],
                 lambda m: f"같은 업종 · 등급 {m['grade']}")
            take(near, lambda m: f"비슷한 등급 ({m['grade']})")
        elif me.get("industry"):
            take([(r, m) for r, m in pool if m.get("industry") == me["industry"]], "같은 업종")
        return out

    if me.get("industry"):
        take([(r, m) for r, m in pool if m.get("industry") == me["industry"]], "같은 업종")
    rev = me.get("revenue")
    if len(out) < limit and rev:
        sim = [(r, m) for r, m in pool if m.get("market") == me.get("market") and m.get("revenue")
               and rev / 3 <= m["revenue"] <= rev * 3]
        sim.sort(key=lambda x: abs(math.log(x[1]["revenue"] / rev)))
        mk = {"Y": "유가증권", "K": "코스닥", "N": "코넥스"}.get(me.get("market"), "")
        take(sim, lambda m: f"{mk} · 매출 {eok(m['revenue'])}(비슷한 규모)".strip(" ·"))
    return out


def clip_sentences(text, limit):
    """문장 단위로 자른다 (중간에서 '…'로 끊지 않음)."""
    if not text:
        return ""
    t = re.sub(r"\s+", " ", text).strip()
    if len(t) <= limit:
        return t
    parts = re.split(r"(?<=[다요음함됨]\.)\s+|(?<=\.)\s+(?=[가-힣A-Z\[(])", t)
    out = ""
    for p in parts:
        if len(out) + len(p) + 1 > limit:
            break
        out = (out + " " + p).strip()
    return out or t[:limit].rsplit(" ", 1)[0]


_EN = (("에이치디", "HD"), ("엘지", "LG"), ("에스케이", "SK"), ("케이티", "KT"), ("씨제이", "CJ"), ("지에스", "GS"),
       ("엘에스", "LS"), ("케이씨씨", "KCC"), ("디엘", "DL"), ("에이치엘", "HL"), ("비지에프", "BGF"), ("케이비", "KB"))


def norm_target(s):
    """교환대상 '에이치디현대마린솔루션 주식회사 기명식 보통주' → 'HD현대마린솔루션'"""
    import re
    t = re.sub(r"\(주\)|㈜|주식회사|기명식|무기명식|보통주식?|우선주식?|전자등록|주식", " ", s or "")
    t = re.split(r"[,(]", t)[0].strip().replace(" ", "")
    for a, b in _EN:
        if t.startswith(a):
            t = b + t[len(a):]
    return t



def load_triggers(path):
    """data/manual/rating_triggers.csv → {종목코드 또는 회사명: [행...]}"""
    import csv
    import os
    out = {}
    if not os.path.exists(path):
        return out
    with open(path, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            k = (r.get("stock_code") or "").strip() or (r.get("corp_name") or "").strip()
            if k:
                out.setdefault(k, []).append({k2: (v or "").strip() for k2, v in r.items()})
    return out


def _cur_metric(metric, unit, m):
    """트리거 지표의 현재값 (계산 못 하면 None). unit이 %면 백분율, 배면 배수."""
    e = m.get("ebitda") if m.get("ebitda_ok") else None
    rev, debt, nd, eq, assets = (m.get("revenue"), m.get("debt"), m.get("net_debt"), m.get("equity"), m.get("assets"))
    cov = m.get("interest_coverage")
    cost = (m.get("op") / cov) if cov and m.get("op") else None
    v = None
    if metric == "순차입금/EBITDA":
        v = nd / e if nd is not None and e and e > 0 else None
    elif metric == "차입금/EBITDA":
        v = debt / e if debt is not None and e and e > 0 else None
    elif metric == "부채비율":
        v = m.get("debt_ratio")
    elif metric == "차입금의존도":
        v = debt / assets * 100 if debt is not None and assets else None
    elif metric == "EBITDA/금융비용":
        v = e / cost if e is not None and cost else None
    elif metric == "영업이익/금융비용":
        v = cov
    elif metric == "EBITDA마진":
        v = e / rev * 100 if e is not None and rev else None
    elif metric == "영업이익률":
        v = m.get("op_margin")
    elif metric == "순차입금/자본":
        v = nd / eq * 100 if nd is not None and eq and eq > 0 else None
    elif metric == "매출액":
        div = {"조원": 1e12, "억원": 1e8}.get(unit)
        v = rev / div if rev is not None and div else None
    if v is None:
        return None
    if unit == "배" and metric in ("부채비율", "차입금의존도", "EBITDA마진", "영업이익률", "순차입금/자본"):
        v = v / 100
    elif unit == "%" and metric in ("순차입금/EBITDA", "차입금/EBITDA", "EBITDA/금융비용", "영업이익/금융비용"):
        v = v * 100
    return round(v, 2)


def trigger_status(rows, m):
    """평가사별 등급·전망 + 트리거별 현재값·충족 여부·여유"""
    if not rows:
        return None
    ags, items = {}, []
    for r in rows:
        ags.setdefault(r["agency"], {k: r.get(k) for k in ("agency", "rating", "outlook", "eval_date", "source_url")})
        if r.get("metric") in ("", "없음"):
            continue
        it = {"agency": r["agency"], "dir": r.get("direction"), "metric": r["metric"], "op": r.get("operator"),
              "unit": r.get("unit"), "text": r.get("condition_text"), "url": r.get("source_url")}
        try:
            thr = float(r["threshold"]) if r.get("threshold") else None
        except ValueError:
            thr = None
        it["thr"] = thr
        cur = _cur_metric(r["metric"], r.get("unit"), m) if thr is not None else None
        if cur is not None and it["op"] in (">", ">=", "<", "<="):
            met = {">": cur > thr, ">=": cur >= thr, "<": cur < thr, "<=": cur <= thr}[it["op"]]
            it.update(cur=cur, met=met, gap=round(thr - cur, 2))
        items.append(it)
    items.sort(key=lambda x: (x["metric"] == "기타", x.get("dir") != "down", x["agency"]))
    return {"agencies": list(ags.values()), "items": items}
