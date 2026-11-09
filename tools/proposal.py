"""공모사채 발행 제안서 생성 — 등록된 PPT 템플릿에 회사·시장 데이터를 채운다.

python -m tools.proposal --corp 00155276 --issue 2026-11-30 --base 1000 --max 2000 --tenors 2,3 --out out.pptx
수동 장표(투자기관 동향·배정 현황·하우스 금리전망 등)는 빼고, 자동 갱신 가능한 장표만 남긴다.
"""
import argparse
import copy
import json
import os
import re
from collections import defaultdict
from datetime import date, timedelta

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
D = lambda *p: os.path.join(ROOT, "data", *p)  # noqa: E731
WD = "월화수목금토일"
HOLI = {"2026-12-25": "성탄절", "2027-01-01": "신정", "2026-10-03": "개천절", "2026-10-09": "한글날"}


def J(*p, default=None):
    try:
        return json.load(open(D(*p)))
    except FileNotFoundError:
        return default


def eok(v):
    return f"{v / 1e8:,.0f}"


def md(d):
    d = date.fromisoformat(d) if isinstance(d, str) else d
    return f"{d.month}/{d.day}({WD[d.weekday()]})"


def bday(d, n):
    """영업일 n일 이동 (주말·공휴일 제외)"""
    step = 1 if n >= 0 else -1
    while n:
        d += timedelta(days=step)
        if d.weekday() < 5 and d.isoformat() not in HOLI:
            n -= step
    return d


def norm(s):
    return re.sub(r"\(주\)|㈜|주식회사|\s", "", s or "")


# ── 텍스트 채우기 (첫 run 서식 유지) ──
def _set_par(par, text, rpr=None):
    runs = par.runs
    if not runs:
        r = par.add_run()
        r.text = text
        if rpr is not None:
            old = r._r.find(qn("a:rPr"))
            if old is not None:
                r._r.remove(old)
            r._r.insert(0, copy.deepcopy(rpr))
        return
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def set_text(obj, text, rpr=None):
    tf = obj.text_frame
    lines = str(text).split("\n")
    ps = tf.paragraphs
    while len(ps) < len(lines):
        ps[-1]._p.addnext(copy.deepcopy(ps[-1]._p))
        ps = tf.paragraphs
    for p_, ln in zip(ps, lines):
        _set_par(p_, ln, rpr)
    for p_ in ps[len(lines):]:
        p_._p.getparent().remove(p_._p)


def set_bul(obj, lines):
    """글머리 수준을 지키며 채움: '- '로 시작하면 하위 수준(템플릿의 두 번째 문단 모양), 아니면 상위 수준."""
    tf = obj.text_frame
    ps = tf.paragraphs
    top = ps[0]._p
    sub = next((x._p for x in ps[1:] if (x.level or 0) > (ps[0].level or 0)), ps[1]._p if len(ps) > 1 else top)
    top, sub = copy.deepcopy(top), copy.deepcopy(sub)
    body = ps[0]._p.getparent()
    for x in ps:
        body.remove(x._p)
    for ln in lines:
        el = copy.deepcopy(sub if ln.startswith("- ") else top)
        body.append(el)
    for x, ln in zip(tf.paragraphs, lines):
        _set_par(x, ln[2:] if ln.startswith("- ") else ln)


def run_rpr(cell):
    for p_ in cell.text_frame.paragraphs:
        for r in p_.runs:
            x = r._r.find(qn("a:rPr"))
            if x is not None:
                return copy.deepcopy(x)
    return None


def table_rpr(table, row=1):
    for r in list(table.rows)[row:]:
        for c in r.cells:
            x = run_rpr(c)
            if x is not None and c.text.strip():
                return x
    return None


def fill_cell(c, text, rpr):
    """서식 통일: 기존 run 서식 대신 rpr로 모두 맞춘다."""
    for p_ in c.text_frame.paragraphs:
        for r in list(p_.runs):
            r._r.getparent().remove(r._r)
    set_text(c, text, rpr)


def shape(slide, name, nth=0):
    hits = [s for s in slide.shapes if s.name == name]
    return hits[nth] if len(hits) > nth else None


def drop(slide, *shapes):
    for s in shapes:
        if s is not None:
            s._element.getparent().remove(s._element)


def keep_only(slide, keep):
    for s in list(slide.shapes):
        if s.name not in keep and not (s.is_placeholder):
            drop(slide, s)


def del_rows(table, idxs):
    tbl = table._tbl
    trs = tbl.findall(qn("a:tr"))
    for i in sorted(idxs, reverse=True):
        tbl.remove(trs[i])


def del_cols(table, idxs):
    tbl = table._tbl
    grid = tbl.find(qn("a:tblGrid"))
    gcs = grid.findall(qn("a:gridCol"))
    for i in sorted(idxs, reverse=True):
        grid.remove(gcs[i])
        for tr in tbl.findall(qn("a:tr")):
            tcs = tr.findall(qn("a:tc"))
            tr.remove(tcs[i])


def unmerge(table):
    for tr in table._tbl.findall(qn("a:tr")):
        for tc in tr.findall(qn("a:tc")):
            for a in ("gridSpan", "rowSpan", "hMerge", "vMerge"):
                if a in tc.attrib:
                    del tc.attrib[a]


def fill_xml(cell):
    pr = cell._tc.tcPr
    f = pr.find(qn("a:solidFill")) if pr is not None else None
    return copy.deepcopy(f) if f is not None else None


def set_fill(cell, fx):
    pr = cell._tc.get_or_add_tcPr()
    for t in ("a:solidFill", "a:noFill"):
        old = pr.find(qn(t))
        if old is not None:
            pr.remove(old)
    if fx is not None:
        pr.insert(0, copy.deepcopy(fx))


def drop_slides(prs, nums):
    sl = prs.slides._sldIdLst
    ids = list(sl)
    for n in sorted(nums, reverse=True):
        rid = ids[n - 1].get(qn("r:id"))
        prs.part.drop_rel(rid)
        sl.remove(ids[n - 1])


# ── 데이터 ──
def load_ctx(corp, today):
    R = J("records.json")["records"]
    B = J("ref", "bond_docs.json", default={})
    dem = J("ref", "mkt", "demand.json", default={})
    out = J("bonds", "outstanding.json", default={"bonds": []})
    pubs = [r for r in R.values() if r.get("type") == "BOND" and r.get("offering") == "공모"]
    # 회사별 최근 등급 (공모채 신고서 기준)
    grade = {}
    for r in sorted(pubs, key=lambda r: r.get("rcept_dt") or ""):
        g = (B.get(r.get("rcept_no")) or B.get(r.get("original_rcept")) or {}).get("rating")
        if g:
            grade[norm(r["corp_name"])] = g
    # 공모채 트랜치별 수요예측 결과
    groups = defaultdict(list)
    for r in pubs:
        groups[(r["corp_code"], r.get("original_rcept") or r["rcept_no"])].append(r)
    for gk, rs in groups.items():
        rs.sort(key=lambda r: [int(x) if x.isdigit() else 0 for x in re.split(r"\D", r.get("series_raw") or "0")])
        tr = None
        for rno, nm in rs[0].get("filings") or []:
            t = (dem.get(rno) or {}).get("tranches")
            if t and len(t) == len(rs):
                tr = t
        for i, r in enumerate(rs):
            bd = B.get(r.get("rcept_no")) or B.get(r.get("original_rcept")) or {}
            s = (bd.get("series") or {}).get(r.get("series_raw")) or {}
            r["_band"], r["_spread"], r["_demand_dt"], r["_grade"] = s.get("band"), s.get("spread"), bd.get("demand_dt"), bd.get("rating")
            r["_dem"] = tr[i] if tr else None
            r["_tenor"] = round((date.fromisoformat(r["maturity"]) - date.fromisoformat(r["pay_dt"])).days / 365, 1) if r.get("maturity") and r.get("pay_dt") else None
    issued = {norm(b.get("name")): b for b in out["bonds"] if b.get("method") == "공모"}
    for r in pubs:
        b = issued.get(norm(f"{r['corp_name']}{r.get('series_raw')}"))
        r["_rate"] = (b or {}).get("cpn") or r.get("coupon")
    # 만기 목록: 최근 3년 공모 발행분(공시) + 채권기본정보의 공모채 (이름 중복 제거)
    mats, seen = [], set()
    for r in pubs:
        if r.get("maturity") and r.get("amount"):
            nm = f"{r['corp_name']} {r.get('series_raw')}"
            seen.add(norm(nm))
            mats.append({"name": nm, "mat": r["maturity"], "amt": r["amount"] / 1e8,
                         "grade": r.get("_grade") or grade.get(norm(r["corp_name"]), "")})
    keys = {(m["mat"], round(m["amt"])) for m in mats}
    for b in out["bonds"]:
        amt_ = round((b.get("bal") or b.get("amt") or 0) / 1e8)
        if re.search(r"CB|EB|BW|전환|교환|신주인수", b.get("name") or ""):
            continue
        if b.get("method") == "공모" and b.get("mat") and norm(b.get("name")) not in seen and (b["mat"], amt_) not in keys:
            mats.append({"name": b.get("name"), "mat": b["mat"], "amt": (b.get("bal") or b.get("amt") or 0) / 1e8,
                         "grade": grade.get(norm(b.get("issuer")), "")})
    mine = sorted([r for r in pubs if r["corp_code"] == corp], key=lambda r: (r.get("pay_dt") or "", r.get("series_raw") or ""))
    return {"pubs": pubs, "grade": grade, "mine": mine, "mats": mats, "out": out, "rates": J("ref", "mkt", "rates.json", default={}),
            "co": J("..", "public", "data", "co", f"{corp}.json") or {}}


def tenor_lab(t):
    return f"{int(round(t))}Y" if t else "-"


# ── 장표별 채우기 ──
def build(args):
    today = date.fromisoformat(args.today) if args.today else date.today()
    issue = date.fromisoformat(args.issue)
    C = load_ctx(args.corp, today)
    co, mine = C["co"], C["mine"]
    name = (co.get("name") or (mine[-1]["corp_name"] if mine else args.corp))
    grade = args.grade or (mine[-1]["_grade"] if mine else None) or co.get("rating") or "-"
    prs = Presentation(args.template)
    S = lambda n: prs.slides[n - 1]  # noqa: E731
    # 직전 발행(수요예측 결과가 있는 가장 최근 회차)
    last = [r for r in mine if r.get("_dem")]
    lk = last[-1]["original_rcept"] if last else None
    prev = [r for r in mine if r.get("original_rcept") == lk] if lk else []
    band = args.band or (prev[0]["_band"] if prev else "개별민평 -30bp ~ +30bp")
    band = band.replace("~", " ~ ").replace("  ", " ")
    tenors = [int(x) for x in args.tenors.split(",")]
    # 만기 물량 제안: 직전 회차 만기별 경쟁률 비율로 최초 물량을 나눔(50억 단위)
    split = {}
    if prev and all(r.get("_dem") for r in prev):
        w = {int(round(r["_tenor"])): r["_dem"]["total"] / (r["_dem"]["ratio"] or 1) for r in prev if r.get("_tenor")}
        tot = sum(w.get(t, 0) for t in tenors) or 1
        for t in tenors:
            split[t] = max(100, round(args.base * w.get(t, tot / len(tenors)) / tot / 50) * 50)
        diff = args.base - sum(split.values())
        split[tenors[-1]] += diff
    else:
        for t in tenors:
            split[t] = args.base // len(tenors)
    # 일정 (직전 회차와 같은 간격: 신고 D-8 · 수요예측 D-6 · 정정 D-4 영업일)
    sch = {"계약": bday(issue, -15), "신고": bday(issue, -8), "수요예측": bday(issue, -6), "정정": bday(issue, -4), "발행": issue}
    sch["협회"] = bday(sch["계약"], 5)
    ev = defaultdict(list)
    ev[sch["계약"]].append("대표주관계약 체결")
    ev[sch["협회"]].append("대표주관계약\n금융투자협회 신고기한")
    ev[sch["신고"]].append("증권신고서 제출")
    ev[sch["수요예측"]].append("수요예측")
    ev[sch["정정"]].append("정정 증권신고서 제출\n(발행조건 확정)")
    ev[issue].append("발행 및 납입\n투자설명서 공시")
    for d_, nm in ((date(2026, 11, 26), "韓 금통위"), (date(2026, 12, 9), "美 FOMC (12/8~9)"), (date(2027, 1, 15), "韓 금통위")):
        ev[d_].append(nm)
    for d_, nm in HOLI.items():
        ev[date.fromisoformat(d_)].append(nm)

    # ── 분석 재료: 시장 스프레드·만기별 수요·펀더멘털·차입구조 ──
    rt = C["rates"]
    aa3_, k3_ = dict(map(tuple, rt.get("aa3", []))), dict(map(tuple, rt.get("ktb3", [])))
    dl = max((d_ for d_ in aa3_ if d_ in k3_), default=None)
    sp_now = (aa3_[dl] - k3_[dl]) * 100 if dl else None
    base_r = (rt.get("base") or [[None, None]])[-1][1]
    since6 = (today - timedelta(days=183)).isoformat()
    aa_rows = [r for r in C["pubs"] if r.get("_dem") and (r.get("_grade") or "").startswith("AA") and (r.get("pay_dt") or "") >= since6]
    by_t, rat_t = defaultdict(lambda: defaultdict(float)), defaultdict(list)
    for r in aa_rows:
        tt = int(round(r["_tenor"] or 0))
        k = "2년" if tt <= 2 else "3년" if tt == 3 else "5년 이상"
        for g, v in r["_dem"]["groups"].items():
            by_t[k]["운용사" if g == "운용사" else "증권사" if g == "증권사" else "연기금·은행·보험 등"] += v
        if r["_dem"].get("ratio"):
            rat_t[k].append(r["_dem"]["ratio"])
    avg = {k: sum(v) / len(v) for k, v in rat_t.items() if v}
    m_ = co.get("metrics") or {}
    fi = co.get("fs_interim") or {}
    lad = defaultdict(float)
    for x in (co.get("bonds") or {}).get("ladder") or []:
        lad[int(x["year"])] += (x.get("공모") or 0) / 1e8
    basd = C["out"].get("basDt", "")
    basd = f"{basd[:4]}-{basd[4:6]}-{basd[6:]}" if len(basd) == 8 else ""
    for r in mine:
        if (r.get("pay_dt") or "") > basd and r.get("maturity"):
            lad[int(r["maturity"][:4])] += r["amount"] / 1e8
    yy = lambda t: issue.year + t  # noqa: E731
    nd_e, icr = m_.get("net_debt_ebitda"), m_.get("interest_coverage")
    capex_h, ocf_h = (fi.get("capex") or 0) / 1e8, (fi.get("ocf") or 0) / 1e8
    dm = {int(round(r["_tenor"])): r["_dem"] for r in prev if r.get("_dem") and r.get("_tenor")}
    sprd = {int(round(r["_tenor"])): r.get("_spread") for r in prev if r.get("_tenor")}

    # p1 표지
    set_text(shape(S(1), "TextBox 3"), f"㈜{name}를 위한\n공모사채 발행 제안서")
    t = shape(S(1), "TextBox 10")
    set_text(t, t.text_frame.text.split("\n")[0] + f"\n{today:%Y.%m}")

    # p4 주요 발행 조건
    s4 = S(4)
    set_text(shape(s4, "모서리가 둥근 직사각형 8"), f"㈜{name} 무보증 공모사채 발행 조건 Summary")
    tb = shape(s4, "표 7").table
    due12 = (co.get("bonds") or {}).get("due12") or 0
    prev_txt = f"직전 회차({prev[0]['series_raw'].split('-')[0]}회, {prev[0]['pay_dt'][2:].replace('-', '.')})" if prev else "직전 회차"
    set_text(tb.cell(1, 2), f"{args.base:,}억 (최대 증액 시 {args.max:,}억원까지 가능)")
    set_text(tb.cell(1, 3), (f"향후 12개월 내 공모채 만기 {due12/1e8:,.0f}억원 차환 수요 — 선제 조달로 차환 리스크 분산\n" if due12 else "")
             + (f"{fi.get('period', '')} 설비투자 {capex_h:,.0f}억원 > 영업현금흐름 {ocf_h:,.0f}억원 — 투자 집행기 외부조달 수요 지속\n" if capex_h else "")
             + f"최초 {args.base:,}억원 신고 후 수요예측 결과에 따라 최대 {args.max:,}억원까지 증액 — 11월 발행 공백 속 대기 수요 흡수")
    set_text(tb.cell(3, 2), " / ".join(f"{t}년({split[t]:,}억원)" for t in tenors))
    dm = {int(round(r["_tenor"])): r["_dem"] for r in prev if r.get("_dem") and r.get("_tenor")}
    set_text(tb.cell(3, 3), "\n".join(x for x in [
        (f"시장: AA- 스프레드 {sp_now:.0f}bp 박스권, AA급 경쟁률 " + " / ".join(f"{t}년 {avg[f'{t}년']:.1f}:1" for t in tenors if f"{t}년" in avg)) if sp_now else None,
        "차입구조: 만기 " + " / ".join(f"’{yy(t) % 100}년 {lad[yy(t)]:,.0f}억" for t in tenors) + " — 증액은 부담 적은 만기 우선",
        (f"펀더멘털: 순차입/EBITDA {nd_e:.1f}배 — 중기물로 금리·차환 부담 균형") if nd_e else None,
    ] if x))
    set_text(tb.cell(4, 2), band)
    set_text(tb.cell(4, 3), "이전 발행과 동일하게 책정")
    set_text(tb.cell(5, 2), grade)
    set_text(tb.cell(5, 3), f"{prev_txt} 증권신고서 기준 신용등급")
    set_text(tb.cell(6, 2), md(sch["수요예측"]))
    set_text(tb.cell(6, 3), f"{md(sch['신고'])} 증권신고서 제출, {md(sch['수요예측'])} 수요예측 → {md(sch['정정'])} 발행조건 확정\n"
             + "11/26(목) 금통위 전 수요예측을 마쳐 금리 결정 이벤트 리스크 최소화")
    set_text(tb.cell(7, 2), md(issue))
    set_text(tb.cell(8, 2), "")
    set_text(tb.cell(8, 3), "전회차와 동일")
    del_rows(tb, [9])

    # p5 발행 일정 달력
    s5 = S(5)
    tb = shape(s5, "표 6").table
    fx = {"holi": fill_xml(tb.cell(2, 1)), "file": fill_xml(tb.cell(6, 5)), "dem": fill_xml(tb.cell(8, 2)), "iss": fill_xml(tb.cell(10, 2))}
    r_date, r_ev = run_rpr(tb.cell(1, 2)), run_rpr(tb.cell(6, 5))
    unmerge(tb)
    mon = sch["계약"] - timedelta(days=sch["계약"].weekday())
    for w in range(5):
        for k in range(5):
            d_ = mon + timedelta(days=7 * w + k)
            fill_cell(tb.cell(1 + 2 * w, 1 + k), f"{d_.month}/{d_.day}", r_date)
            c = tb.cell(2 + 2 * w, 1 + k)
            e = ev.get(d_, [])
            fill_cell(c, "\n".join(e), r_ev)
            kind = ("holi" if d_.isoformat() in HOLI else "iss" if d_ == issue else "dem" if d_ == sch["수요예측"]
                    else "file" if d_ in (sch["신고"], sch["정정"]) else None)
            set_fill(c, fx[kind] if kind else None)
    set_text(shape(s5, "TextBox 9"), f"국내외 통화정책 이벤트를 고려하여 {md(issue)} 발행일정 제안")
    set_bul(shape(s5, "TextBox 5"), (
             f"주요 일정(11/26 韓 금통위, 12/8~9 美 FOMC)을 고려하여 {md(sch['신고'])} 신고, {md(sch['수요예측'])} 수요예측, {md(issue)} 발행/납입 일정 제안\n"
             "- 韓 기준금리 3.00%(10/22 동결) — 7·8월 연속 인상 후 숨고르기, 결정문에서 '추가 인상의 시기와 속도'를 언급해 11월 인상 가능성 상존\n"
             "- 美 연방기금금리 3.75~4.00%(10/28 동결, 9월 25bp 인상) — 12월 인상 확률이 높게 반영된 상황\n"
             f"→ 금통위(11/26) 이전 수요예측({md(sch['수요예측'])})으로 가산금리를 먼저 확정해 이벤트 리스크를 최소화하는 전략 제안").split("\n"))

    # p6 트랜치 구성 제안
    s6 = S(6)
    tb = shape(s6, "표 13").table
    pubs = C["pubs"]
    since = (today - timedelta(days=183)).isoformat()
    aa = [r for r in pubs if r.get("_dem") and (r.get("_grade") or "").startswith("AA") and (r.get("pay_dt") or "") >= since]
    by_t = defaultdict(lambda: defaultdict(float))
    rat_t = defaultdict(list)
    for r in aa:
        tt = int(round(r["_tenor"] or 0))
        k = "2년" if tt <= 2 else "3년" if tt == 3 else "5년 이상"
        for g, v in r["_dem"]["groups"].items():
            gg = "운용사" if g == "운용사" else "증권사" if g == "증권사" else "연기금·은행·보험 등"
            by_t[k][gg] += v
        if r["_dem"].get("ratio"):
            rat_t[k].append(r["_dem"]["ratio"])
    avg = {k: sum(v) / len(v) for k, v in rat_t.items() if v}
    lt = min(tenors, key=lambda t: lad[yy(t)])
    rp6 = run_rpr(tb.cell(1, 2))
    for i, t in enumerate(tenors[:2]):
        set_text(tb.cell(1 + i, 1), f"{t}년 : {split[t]:,}억원\n(최대 {split[t] * args.max // args.base:,}억원)")
        k = f"{t}년"
        L = []
        if k in avg:
            L.append(f"[수요] 최근 6개월 AA급 {t}년물 평균 경쟁률 {avg[k]:.1f}:1(참여 {sum(by_t[k].values()) / 1e4:,.2f}조원)"
                     + (" — 북클로징·인상 경계 속 단기물 선호" if t == 2 else " — 벤치마크 만기, ALM 기관 수요"))
        if t == 2 and base_r:
            L.append(f"[금리] 11월 인상 가능성은 국고 3년-기준금리 {(k3_[dl] - base_r) * 100:.0f}bp로 선반영 — 단기물 금리 상단 제한")
        L.append(f"[차입구조] 만기 ’{yy(t) % 100}년 공모채 만기 {lad[yy(t)]:,.0f}억원"
                 + (" — 상대적 여유, 증액분 우선 배정" if t == lt else " — 만기 집중, 증액은 수요 확인 후"))
        if t != 2 and nd_e:
            L.append(f"[펀더멘털] 순차입금/EBITDA {nd_e:.1f}배, 이자보상배율 {icr:.1f}배 — 만기 장기화로 차환 주기 분산")
        if t in dm:
            L.append(f"[직전 회차] {t}년물 경쟁률 {dm[t]['ratio']:.1f}:1" + (f", 확정 개평 {sprd[t]:+d}bp" if sprd.get(t) is not None else ""))
        fill_cell(tb.cell(1 + i, 2), "\n".join(L), rp6)
    ch = shape(s6, "차트 11").chart
    cats = [k for k in ("2년", "3년", "5년 이상") if k in by_t]
    cd = CategoryChartData()
    cd.categories = cats
    for g in ("운용사", "증권사", "연기금·은행·보험 등"):
        cd.add_series(g, [round(by_t[k][g]) for k in cats])
    ch.replace_data(cd)
    for dl in ch._chartSpace.xpath(".//c:dLbl"):
        dl.getparent().remove(dl)
    for x in [x for x in s6.shapes if x.shape_type == 9 or x.name in ("그룹 12", "그룹 17")]:
        drop(s6, x)
    for nm in ("타원 21", "타원 22", "타원 26", "타원 27", "타원 28"):
        drop(s6, shape(s6, nm))
    set_text(shape(s6, "모서리가 둥근 직사각형 8"), "AA급 공모사채 만기별 수요예측 참여금액 (최근 6개월)")
    set_text(shape(s6, "TextBox 29"), f"자료: 기업공시시스템 ({today.month}/{today.day} 수요예측 완료 건 기준, 단위: 억원)")
    drop(s6, shape(s6, "차트 1"), shape(s6, "표 2"), shape(s6, "TextBox 3"), shape(s6, "모서리가 둥근 직사각형 8", 1))
    set_text(shape(s6, "TextBox 10"), "연말 우량물 희소성과 만기별 수요·만기 분산을 고려한 " + "/".join(str(t) for t in tenors) + "년 만기 구성")

    # p7·8 금통위·연준 위원 (11월 업데이트)
    s7 = S(7)
    tb = [x for x in s7.shapes if x.has_table][0].table
    upd7 = {
        "신현송": ("매파", "7·8월 연속 인상(2.50→3.00%) 주도. 10월 동결 후에도 추가 인상의 시기·속도를 결정하겠다는 입장"),
        "황건일": ("중도", "8/27 인상 때 '동결' 소수의견 — 인상 속도에 신중(총재는 '전술적 차이'라고 설명)"),
        "김진일": ("중도~매파", "5월 취임. 금융위기 예방을 중시하며 스스로 '중간값보다 반 클릭 위'라고 평가"),
    }
    for r in list(tb.rows)[1:]:
        nm = r.cells[0].text.strip()
        if nm in upd7:
            set_text(r.cells[4], upd7[nm][0])
            set_text(r.cells[5], upd7[nm][1])
        if nm == "김진일":
            set_text(r.cells[1], "금통위원")
    set_text(shape(s7, "TextBox 4"), "7·8월 연속 인상으로 기준금리 3.00%, 10월 만장일치 동결\n물가 3% 내외 지속 속 '추가 인상 시기·속도' 언급 — 매파적 동결")
    set_text(shape(s7, "Text 17"), "10/22 결정문: 물가가 3% 내외의 높은 오름세를 이어가는 가운데 두 차례 인상의 파급효과를 점검하며 추가 인상의 시기와 속도를 결정. 8월 점도표(6개월 후)는 3.25%가 최다(21개 중 10개)")
    s8 = S(8)
    tb = [x for x in s8.shapes if x.has_table][0].table
    for r in list(tb.rows)[1:]:
        nm = r.cells[0].text.strip()
        c5 = r.cells[5].text
        c5 = re.sub(r"\s*4월 FOMC[^.]*$", "", c5.strip())
        if nm.startswith("Lisa"):
            set_text(r.cells[4], "매파")
            set_text(r.cells[5], "인플레이션 리스크를 우려해 10월 FOMC에서 25bp 인상 주장(소수의견)")
        elif nm.startswith("Kevin"):
            set_text(r.cells[5], "9월 25bp 인상(2023년 7월 이후 첫 인상) 주도. 포워드 가이던스는 지양")
        else:
            c5 = re.split(r"(?<=[.다])\s", c5)[0]
            set_text(r.cells[5], (c5 + " " if c5 else "") + "9월 인상·10월 동결 찬성")
    set_text(shape(s8, "TextBox 5"), "9월 25bp 인상 후 10월 3.75~4.00% 동결(10:2, 인상 소수의견 2명)\n9월 CPI 3.7%로 재가속 — 12월 추가 인상 확률이 높게 반영")
    set_text(shape(s8, "Text 17"), "10/28 FOMC: Cook 이사·Hammack 총재가 25bp 인상을 주장하며 반대. 성명서는 경제활동을 'solid'→'strong'으로 상향, 9월 인상이 물가 목표 복귀를 돕는다는 문구 삭제")

    # p10 직전 발행 수요예측 결과
    s10 = S(10)
    if prev:
        p0 = prev[0]
        sn = p0["series_raw"].split("-")[0]
        set_text(shape(s10, "텍스트 개체 틀 3"), f"\n前발행 수요예측 상세결과 ({name}{sn}, {(p0['_demand_dt'] or '')[2:].replace('-', '.')})")
        tot_d = sum(r["_dem"]["total"] for r in prev)
        init = sum(r["_dem"]["total"] / r["_dem"]["ratio"] for r in prev if r["_dem"].get("ratio"))
        fin = sum(r["amount"] for r in prev) / 1e8
        tb = shape(s10, "표 20").table
        set_text(tb.cell(1, 1), f"{grade}")
        set_text(tb.cell(2, 1), f"모집 {init:,.0f}억원 → 주문 {tot_d:,.0f}억원 (경쟁률 {tot_d / init:.2f}:1)")
        set_text(tb.cell(3, 1), " 신고 " + " / ".join(f"{tenor_lab(r['_tenor']).replace('Y', '년')} {r['_dem']['total'] / r['_dem']['ratio']:,.0f}억원" for r in prev) + f" → 총 {init:,.0f}억원")
        set_text(tb.cell(4, 1), " 발행 " + " / ".join(f"{tenor_lab(r['_tenor']).replace('Y', '년')} {r['amount'] / 1e8:,.0f}억원" for r in prev) + f" → 총 {fin:,.0f}억원")
        uses = defaultdict(float)
        for r in prev:
            for k, v in (r.get("uses") or {}).items():
                if "비용" not in k:
                    uses[k] += v
        set_text(tb.cell(5, 1), " · ".join(uses) or "-")
        set_text(tb.cell(6, 1), " " + " / ".join(f"{tenor_lab(r['_tenor']).replace('Y', '년')} 개평 {r['_spread']:+d}bp" + (f" ({r['_rate']:.3f}%)" if r.get("_rate") else "") for r in prev if r.get("_spread") is not None))
        del_rows(tb, [5, 7])
        set_text(shape(s10, "모서리가 둥근 직사각형 8", 0), f"{name}{sn} 발행개요")
        set_text(shape(s10, "모서리가 둥근 직사각형 8", 1), f"{name}{sn} 수요예측 기관군별 참여")
        tb = shape(s10, "표 2").table
        set_text(tb.cell(1, 2), "참여건수")
        set_text(tb.cell(1, 4), "참여건수")
        for i, r in enumerate(prev[:2]):
            set_text(tb.cell(0, 1 + 2 * i), tenor_lab(r["_tenor"]).replace("Y", "년물"))
        rows = [("증권사", ["증권사"]), ("운용사", ["운용사"]), ("연기금·은행·보험", ["연기금·은행·보험"]), ("기타·외국", ["기타", "외국(실적유)", "외국(실적무)"])]
        for ri, (lab, ks) in enumerate(rows):
            set_text(tb.cell(2 + ri, 0), lab)
            for i, r in enumerate(prev[:2]):
                g, c = r["_dem"]["groups"], r["_dem"]["cnt"]
                set_text(tb.cell(2 + ri, 1 + 2 * i), f"{sum(g.get(k, 0) for k in ks):,.0f}" if sum(g.get(k, 0) for k in ks) else "-")
                set_text(tb.cell(2 + ri, 2 + 2 * i), f"{sum(c.get(k, 0) for k in ks):,.0f}" if sum(c.get(k, 0) for k in ks) else "-")
        for i, r in enumerate(prev[:2]):
            set_text(tb.cell(6, 1 + 2 * i), f"{r['_dem']['total']:,.0f}")
            set_text(tb.cell(6, 2 + 2 * i), f"{r['_dem']['n']:,.0f}")
        best = max(prev, key=lambda r: r["_dem"]["ratio"] or 0)
        big = max(((g, sum(r["_dem"]["groups"].get(g, 0) for r in prev)) for g in ("운용사", "증권사", "연기금·은행·보험")), key=lambda x: x[1])
        set_text(shape(s10, "TextBox 4"), f"주문 {tot_d:,.0f}억원(경쟁률 {tot_d / init:.1f}:1)으로 {fin:,.0f}억원까지 증액 발행. {tenor_lab(best['_tenor']).replace('Y', '년물')} 수요 우위")
        set_text(shape(s10, "직사각형 24"),
                 f"{big[0]} 중심의 참여\n전체 주문 {tot_d:,.0f}억원 중 {big[0]} {big[1]:,.0f}억원({big[1] / tot_d * 100:.0f}%)으로 최대 투자자군\n\n"
                 + "\n".join(f"{tenor_lab(r['_tenor']).replace('Y', '년물')}: 신고 {r['_dem']['total'] / r['_dem']['ratio']:,.0f}억 → 주문 {r['_dem']['total']:,.0f}억 (경쟁률 {r['_dem']['ratio']:.1f}:1), 확정 개평 {r['_spread']:+d}bp, 발행 {r['amount'] / 1e8:,.0f}억" for r in prev)
                 )
        set_text(shape(s10, "모서리가 둥근 직사각형 8", 2), f"{name}{sn} 수요예측 결과 분석")
        set_text(shape(s10, "모서리가 둥근 직사각형 8", 3), "이번 회차 시사점")
        set_text(shape(s10, "직사각형 27"),
                 f"{tenor_lab(best['_tenor']).replace('Y', '년물')} 수요 우위 활용\n직전 회차에서 경쟁률이 가장 높았던 만기 중심으로 최초 물량을 배분하고, 증액 여력을 확보\n\n"
                 f"동일 밴드({band}) 유지\n직전 회차 확정 가산금리(" + " / ".join(f"{r['_spread']:+d}bp" for r in prev if r.get("_spread") is not None) + ")가 밴드 하단~중간에서 결정돼 같은 밴드로 금리 경쟁력 확보 가능\n\n"
                 f"기존 참여 투자자군 재공략\n직전 주문의 {big[0]}·증권사 수요를 중심으로 사전 마케팅(NDR) 진행")

    # p13 향후 공모사채 발행 정보
    s13 = S(13)
    tb = shape(s13, "표 2").table
    up = sorted([r for r in pubs if (r.get("pay_dt") or "") > today.isoformat() and r.get("_tenor") and r.get("_demand_dt") and r["corp_code"] != args.corp], key=lambda r: (r.get("_demand_dt") or "", r["corp_name"], r.get("series_raw") or ""))
    rp13 = table_rpr(tb)
    unmerge(tb)
    del_cols(tb, [4, 6, 7])
    nrow = len(tb.rows) - 1
    up = up[:nrow]
    for i in range(nrow):
        r = up[i] if i < len(up) else None
        cells = [tb.cell(1 + i, j) for j in range(8)]
        vals = [""] * 8
        if r:
            first = i == 0 or up[i - 1]["corp_code"] != r["corp_code"] or up[i - 1].get("original_rcept") != r.get("original_rcept")
            vals = [r["corp_name"] if first else "", (r.get("_grade") or "") if first else "", tenor_lab(r.get("_tenor")), eok(r["amount"]),
                    (r.get("_band") or "") if first else "", md(r["_demand_dt"]) if first and r.get("_demand_dt") else "",
                    md(r["pay_dt"]) if first else "", (r.get("lead") or "").replace("증권", "").replace("투자", "") if first else ""]
        for c, v in zip(cells, vals):
            fill_cell(c, v, rp13)
    if len(up) < nrow:
        del_rows(tb, list(range(1 + len(up), 1 + nrow)))
    set_text(shape(s13, "TextBox 3"), "자료: 기업공시시스템 (증권신고서 기준, 금액은 신고 금액·억원)")
    set_text(shape(s13, "TextBox 4"), f"{today.month}/{today.day} 기준 수요예측·발행 예정 공모사채 {len({(r['corp_code'], r.get('original_rcept')) for r in up})}건")

    # p17 크레딧 (AA- 3년 vs 국고 3년)
    s17 = S(17)
    rt = C["rates"]
    one = (today - timedelta(days=365)).isoformat()
    aa3, k3 = dict(map(tuple, rt.get("aa3", []))), dict(map(tuple, rt.get("ktb3", [])))
    ds = [d_ for d_ in sorted(aa3) if d_ >= one and d_ in k3]
    cd = CategoryChartData()
    cd.categories = [date.fromisoformat(d_) for d_ in ds]
    cd.add_series("SP", [round((aa3[d_] - k3[d_]) * 100, 1) for d_ in ds])
    cd.add_series("국고 3년", [k3[d_] for d_ in ds])
    cd.add_series("AA- 3년", [aa3[d_] for d_ in ds])
    shape(s17, "차트 4").chart.replace_data(cd)
    set_text(shape(s17, "모서리가 둥근 직사각형 8"), "최근 1년 금리 및 스프레드 추이 (AA- 3년)")
    set_text(shape(s17, "TextBox 63"), "자료: 한국은행 경제통계시스템(ECOS)")
    if ds:
        a, b = ds[-1], ds[0]
        sp1, sp0 = (aa3[a] - k3[a]) * 100, (aa3[b] - k3[b]) * 100
        lo = min(ds, key=lambda d_: aa3[d_] - k3[d_])
        MC = J("manual", "market_comment.json", default={})
        cr = list(MC.get("credit", []))
        cr.insert(1, f"- {a[5:].replace('-', '/')} 기준 AA- 3년 {aa3[a]:.3f}%, 국고 3년 대비 {sp1:.0f}bp — 1년 전({b[:7]}) {sp0:.0f}bp 대비 {sp1 - sp0:+.0f}bp, 기간 중 최저 {(aa3[lo] - k3[lo]) * 100:.0f}bp({lo[:7]})")
        set_bul(shape(s17, "TextBox 17"), cr)
    drop(s17, shape(s17, "그림 14"), shape(s17, "TextBox 15"), shape(s17, "모서리가 둥근 직사각형 8", 1))

    # p18 국고채·기준금리 — 그림을 자동 차트로 교체
    s18 = S(18)
    pic = shape(s18, "그림 49")
    old = [x for x in s18.shapes if x.name == "그룹 23"]
    two = (today - timedelta(days=730)).isoformat()
    k10, k5, base = dict(map(tuple, rt.get("ktb10", []))), dict(map(tuple, rt.get("ktb5", []))), dict(map(tuple, rt.get("base", [])))
    ds = [d_ for d_ in sorted(k3) if d_ >= two and d_ in k10 and d_ in k5]
    bl, cur = [], None
    for d_ in sorted(set(ds) | set(base)):
        cur = base.get(d_, cur)
        if d_ in ds:
            bl.append(cur)
    cd = CategoryChartData()
    cd.categories = [date.fromisoformat(d_) for d_ in ds]
    cd.add_series("국고채 3년", [k3[d_] for d_ in ds])
    cd.add_series("국고채 5년", [k5[d_] for d_ in ds])
    cd.add_series("국고채 10년", [k10[d_] for d_ in ds])
    cd.add_series("기준금리", bl)
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    if old:
        o = old[0]
        gf = s18.shapes.add_chart(XL_CHART_TYPE.LINE, o.left, o.top, o.width, o.height, cd)
        c = gf.chart
        c.has_legend, c.legend.position, c.legend.include_in_layout = True, XL_LEGEND_POSITION.TOP, False
        c.legend.font.size = c.font.size = __import__("pptx").util.Pt(8)
        for sr in c.plots[0].series:
            sr.smooth = False
            sr.format.line.width = __import__("pptx").util.Pt(1.25)
        c.category_axis.tick_labels.number_format, c.category_axis.tick_labels.number_format_is_linked = "yy.mm", False
        c.value_axis.has_major_gridlines = False
        vals = [v for v in list(k3.values()) + list(k10.values()) + bl if v]
        c.value_axis.minimum_scale = max(0, int(min(vals) * 2) / 2 - 0.5) if vals else None
        c.value_axis.tick_labels.number_format, c.value_axis.tick_labels.number_format_is_linked = "0.0", False
        drop(s18, o)
    drop(s18, pic, shape(s18, "직사각형 50"), shape(s18, "TextBox 51"), shape(s18, "TextBox 52"), shape(s18, "모서리가 둥근 직사각형 8", 1))
    set_text(shape(s18, "TextBox 53"), "자료: 한국은행 경제통계시스템(ECOS)")
    if ds:
        a = ds[-1]
        kr = list(J("manual", "market_comment.json", default={}).get("kr", []))
        kr.append(f"- {a[5:].replace('-', '/')} 기준 국고 3년 {k3[a]:.3f}% / 5년 {k5[a]:.3f}% / 10년 {k10[a]:.3f}%, 기준금리 대비 3년물 {(k3[a] - (bl[-1] or 0)) * 100:+.0f}bp")
        set_bul(shape(s18, "TextBox 25"), kr)

    # p19 美 채권시장
    s19 = S(19)
    u2, u10, fh = dict(map(tuple, rt.get("ust2", []))), dict(map(tuple, rt.get("ust10", []))), dict(map(tuple, rt.get("ffr_hi", [])))
    ds = [d_ for d_ in sorted(u10) if d_ >= two and d_ in u2]
    fl, cur = [], None
    for d_ in sorted(set(ds) | set(fh)):
        cur = fh.get(d_, cur)
        if d_ in ds:
            fl.append(cur)
    cd = CategoryChartData()
    cd.categories = ds
    cd.add_series("美 국채 2년물", [u2[d_] for d_ in ds])
    cd.add_series("美 국채 10년물", [u10[d_] for d_ in ds])
    cd.add_series("美 기준금리(상단)", fl)
    shape(s19, "차트 22").chart.replace_data(cd)
    keep_only(s19, {"텍스트 개체 틀 3", "TextBox 17", "모서리가 둥근 직사각형 8", "차트 22", "TextBox 24", "TextBox 31"})
    drop(s19, shape(s19, "모서리가 둥근 직사각형 8", 1))
    set_text(shape(s19, "TextBox 24"), "자료: FRED(St. Louis Fed)")
    if ds:
        a = ds[-1]
        us = list(J("manual", "market_comment.json", default={}).get("us", []))
        set_bul(shape(s19, "TextBox 17"), us)

    # p20 만기도래 현황 (공모 회사채, 내년)
    s20 = S(20)
    yr = today.year + 1
    gmap = C["grade"]
    mon_g = defaultdict(lambda: defaultdict(float))
    for b in C["mats"]:
        if not b["mat"].startswith(str(yr)):
            continue
        g = b["grade"]
        gg = "AAA급" if g.startswith("AAA") else "AA급" if g.startswith("AA") else "A급" if g.startswith("A") else "BBB급 이하"
        mon_g[int(b["mat"][5:7])][gg] += b["amt"]
    cd = CategoryChartData()
    cd.categories = [f"{m}월" for m in range(1, 13)]
    for g in ("AAA급", "AA급", "A급", "BBB급 이하"):
        cd.add_series(g if g != "BBB급 이하" else "BBB급 이하·미확인", [round(mon_g[m][g]) for m in range(1, 13)])
    cd.add_series("열2", [None] * 12)
    cd.add_series("분기", [None] * 12)
    shape(s20, "차트 35").chart.replace_data(cd)
    tb = shape(s20, "표 49").table
    for ri, g in enumerate(("AAA급", "AA급", "A급", "BBB급 이하")):
        set_text(tb.cell(ri, 0), g.replace("급", "").replace(" 이하", "") if g != "BBB급 이하" else "그 외")
        for m in range(1, 13):
            set_text(tb.cell(ri, m), f"{mon_g[m][g]:,.0f}")
    tots = [sum(mon_g[m].values()) for m in range(1, 13)]
    lbl = sorted([s for s in s20.shapes if s.has_text_frame and re.fullmatch(r"[\d,]+", s.text_frame.text.strip() or "x")], key=lambda s: (s.top < 2500000, s.left))
    qs = sorted([s for s in lbl if s.name in ("TextBox 21", "TextBox 23", "TextBox 24", "TextBox 25")], key=lambda s: s.left)
    ms = sorted([s for s in lbl if s not in qs], key=lambda s: s.left)
    for s, v in zip(ms, tots):
        set_text(s, f"{v:,.0f}")
    for i, s in enumerate(qs):
        set_text(s, f"{sum(tots[3 * i:3 * i + 3]):,.0f}")
    set_text(shape(s20, "텍스트 개체 틀 3"), f"\n‘{yr % 100}년 회사채 만기도래 현황")
    set_text(shape(s20, "모서리가 둥근 직사각형 8"), f"{yr}년 공모 회사채 만기도래 현황 (월별·등급별)")
    set_text(shape(s20, "TextBox 34"), "자료: 기업공시시스템(최근 3년 공모 발행분), 금융위원회 채권기본정보 / BBB급 이하에 등급 미확인분 포함")
    q = [sum(tots[3 * i:3 * i + 3]) for i in range(4)]
    pk = max(range(12), key=lambda i: tots[i])
    set_bul(shape(s20, "TextBox 17"), (
             f"‘{yr % 100}년 공모 회사채 만기도래 규모는 약 {sum(tots) / 1e4:,.1f}조원 수준\n"
             f"- 1분기 {q[0] / 1e4:,.1f}조원 / 2분기 {q[1] / 1e4:,.1f}조원 / 3분기 {q[2] / 1e4:,.1f}조원 / 4분기 {q[3] / 1e4:,.1f}조원, 상반기 비중 {(q[0] + q[1]) / max(1, sum(q)) * 100:.0f}%\n"
             f"- 월별로는 {pk + 1}월에 약 {tots[pk] / 1e4:,.1f}조원으로 가장 많은 금액이 만기 예정\n"
             "연초 만기 집중 전 선제적 차환 발행 수요 — 연말 발행으로 내년 1분기 차환 부담 분산 가능").split("\n"))

    # p22·23 월별 만기 달력 (AA급) — 발행월과 다음 달
    for sl, (y, m) in ((S(22), (issue.year, issue.month)), (S(23), ((issue.year + issue.month // 12), issue.month % 12 + 1))):
        tb = [x for x in sl.shapes if x.has_table][0].table
        unmerge(tb)
        first = date(y, m, 1)
        mon = first - timedelta(days=first.weekday())
        if first.weekday() >= 5:
            mon += timedelta(days=7)
        r_d, r_e = run_rpr(tb.cell(3, 1)), table_rpr(tb, 4)
        evd = defaultdict(list)
        for b in C["mats"]:
            g = b["grade"] or "-"
            if b["amt"] < 100:
                continue
            d_ = date.fromisoformat(b["mat"])
            while d_.weekday() >= 5 or d_.isoformat() in HOLI:
                d_ += timedelta(days=1)
            if mon <= d_ < mon + timedelta(days=35):
                evd[d_].append((b["name"], b["amt"], g))
        for w in range(5):
            wk_tot = 0
            for k in range(5):
                d_ = mon + timedelta(days=7 * w + k)
                fill_cell(tb.cell(1 + 2 * w, 1 + k), f"{d_.month}월 {d_.day}일" if d_.month == m else "", r_d)
                items = evd.get(d_, []) if d_.month == m or w == 0 else []
                wk_tot += sum(x[1] for x in items)
                txt = HOLI.get(d_.isoformat()) or "\n".join(f"{nm}({amt:,.0f}/{g})" for nm, amt, g in items[:5]) + (f"\n외 {len(items) - 5}건" if len(items) > 5 else "")
                fill_cell(tb.cell(2 + 2 * w, 1 + k), txt, r_e)
                set_fill(tb.cell(2 + 2 * w, 1 + k), None)
            set_text(tb.cell(2 + 2 * w, 0), f"{m}월 {w + 1}주 만기도래\n{wk_tot:,.0f}억원")
        set_text(shape(sl, "텍스트 개체 틀 3"), f"\n‘{y % 100}년 {m}월 만기도래 내역 (공모 회사채)")
        tot = sum(x[1] for d_, v in evd.items() if d_.month == m for x in v)
        set_text(shape(sl, "TextBox 8"), f"{m}월 공모 회사채 만기도래 약 {tot:,.0f}억원 (100억 이상)")
        set_text(shape(sl, "TextBox 5"), "주) 비영업일 만기 도래분은 직후 영업일로 반영, 등급은 최근 증권신고서 기준\n자료: 기업공시시스템, 금융위원회 채권기본정보")

    # p24 최근 분기 AA- 공모사채 발행내역
    s24 = S(24)
    tb = shape(s24, "표 2").table
    qs_ = date(today.year, 3 * ((today.month - 1) // 3) + 1, 1)
    rows = sorted([r for r in pubs if (r.get("_grade") or "") == grade and qs_.isoformat() <= (r.get("pay_dt") or "") <= today.isoformat()],
                  key=lambda r: (r["pay_dt"], r["corp_name"], r.get("series_raw") or ""), reverse=True)
    nrow = len(tb.rows) - 1
    rows = rows[:nrow]
    for i in range(nrow):
        r = rows[i] if i < len(rows) else None
        if not r:
            continue
        dm_ = r.get("_dem") or {}
        init = dm_["total"] / dm_["ratio"] if dm_.get("ratio") else None
        vals = [f"{r['corp_name']} {r['series_raw']}", r["pay_dt"], r.get("maturity") or "", tenor_lab(r.get("_tenor")).replace("Y", ""),
                f"{init:,.0f}" if init else "-", eok(r["amount"]), f"{dm_['total']:,.0f}" if dm_.get("total") else "-",
                f"{dm_['ratio']:.1f}" if dm_.get("ratio") else "-", (r.get("_band") or "-").replace("개별민평", "개별민평 "),
                f"{r['_spread']}bp" if r.get("_spread") is not None else "-", f"{r['_rate']:.3f}" if r.get("_rate") else "-"]
        for j, v in enumerate(vals):
            set_text(tb.cell(1 + i, j), v)
    if len(rows) < nrow:
        del_rows(tb, list(range(1 + len(rows), 1 + nrow)))
    qn_ = (today.month - 1) // 3 + 1
    set_text(shape(s24, "텍스트 개체 틀 3"), f"\n‘{today.year % 100}년 {qn_}분기 {grade}급 공모사채 발행내역")
    set_text(shape(s24, "TextBox 12"), "자료: 기업공시시스템, 금융위원회 채권기본정보 (확정금리 미반영 건은 '-')")

    # 수동 장표 제거: 11·12 투자기관 동향, 14·15 배정기관, 21 하우스 금리전망 (+ 예정 발행이 없으면 13)
    drop_slides(prs, [11, 12, 14, 15, 21] + ([] if up else [13]))
    prs.save(args.out)
    return {"split": split, "sch": {k: v.isoformat() for k, v in sch.items()}, "prev": [r["series_raw"] for r in prev], "up": len(up), "p24": len(rows)}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--corp", required=True)
    ap.add_argument("--issue", required=True)
    ap.add_argument("--base", type=int, default=1000)
    ap.add_argument("--max", type=int, default=2000)
    ap.add_argument("--tenors", default="2,3")
    ap.add_argument("--grade")
    ap.add_argument("--band")
    ap.add_argument("--today")
    ap.add_argument("--out", required=True)
    print(build(ap.parse_args()))
