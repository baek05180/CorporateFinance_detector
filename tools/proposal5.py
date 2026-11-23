"""공모사채 발행 제안서 v5 — 기존 생성기(tools.proposal) 결과물을 v5 포맷으로 바꾼다.

  python -m tools.proposal5 --template 원본템플릿.pptx --inputs inputs.json --out out.pptx

inputs.json (사이트 팝업이 보내는 값)
  {"corp": "00126566", "issue": "2027-01-15", "demand": "2027-01-06", "base": 3000, "max": 4000,
   "split": "2:600,3:1500,5:900", "priority": 3, "band": "", "fee_bp": "", "issues_note": "", "credit_note": ""}

구성·문구 기본값·이벤트·휴일은 docs/proposal_spec.md 의 '설정' 블록에서 읽는다.
확정할 수 없는 값은 빨간 ## 로 남긴다.
"""
import argparse
import copy
import json
import os
import re
import sys
import tempfile
from datetime import date, timedelta
from types import SimpleNamespace

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_MARKER_STYLE, XL_TICK_LABEL_POSITION
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from tools import proposal as P3  # noqa: E402
from tools.pptx_kit import (  # noqa: E402
    BODY_FONT, DGREY, LGREY, MSO_ANCHOR, NAVY, PALE, PALEBLUE, PP_ALIGN, RED, SKY, SUB, TXT, W, X0, Proto, RGBColor,
    bar, etree, fill_tf, headline, headline_font, notes, rect, replace_in_slide, set_shape_text, source, table, textbox,
    title)
from tools.proposal_spec import load_spec  # noqa: E402
from tools import proposal_mkt as M  # noqa: E402
from tools.proposal_text import (  # noqa: E402
    adj_line, context, credit_story, demand_line, fresh_view, kr_lines, spread_line, term_premium, us_lines)
from tools.proposal_mkt import ymd  # noqa: E402

CW = 4770000
X2 = X0 + CW + 250000
CW3 = (W - 2 * 200000) // 3
RX, RW = 5489407, 4752279
WD = "월화수목금토일"


# ───────────────────────── 공통 ─────────────────────────
def md(d):
    d = date.fromisoformat(d) if isinstance(d, str) else d
    return f"{d.month}/{d.day}({WD[d.weekday()]})"


def ymd2(d):
    d = date.fromisoformat(d) if isinstance(d, str) else d
    return f"'{d.year % 100}.{d.month}/{d.day}"


def josa(name, a, b):
    """받침 유무로 조사 선택 (을/를, 이/가 ...)."""
    ch = name.strip()[-1]
    if "가" <= ch <= "힣":
        return a if (ord(ch) - 0xAC00) % 28 else b
    return b


def H(slide, text):
    tb = headline(slide, text)
    headline_font(tb)
    return tb


def tname(t):
    return f"{int(round(t))}년"


def num(v, d=0):
    return f"{v:,.{d}f}"


def series_at(pairs, d):
    """[[date, v], ...] 에서 d 이전 가장 가까운 값."""
    best = None
    for dd, v in pairs:
        if dd <= d:
            best = v
        else:
            break
    return best


def find_slide(prs, *keys):
    for s in prs.slides:
        for sh in s.shapes:
            if sh.has_text_frame and sh.name.startswith("텍스트 개체 틀") and all(k in sh.text_frame.text for k in keys):
                return s
    return None


def new_slide(prs, layout):
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    return s


def dup_slide(prs, src):
    """같은 레이아웃으로 새 장표를 만들고 도형·그림 관계를 그대로 복사."""
    s = new_slide(prs, src.slide_layout)
    rmap = {}
    for rid, rel in src.part.rels.items():
        if rel.reltype.endswith("/slideLayout") or rel.reltype.endswith("/notesSlide"):
            continue
        rmap[rid] = s.part.relate_to(rel._target, rel.reltype) if not rel.is_external else s.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
    for el in src.shapes._spTree.iterchildren():
        if el.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        new = copy.deepcopy(el)
        for node in new.iter():
            for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
                if node.get(attr) in rmap:
                    node.set(attr, rmap[node.get(attr)])
        s.shapes._spTree.append(new)
    return s


def relayout(slide, layout):
    for rel in slide.part.rels.values():
        if rel.reltype.endswith("/slideLayout"):
            rel._target = layout.part


def remove_slide(prs, slide):
    for el in list(prs.slides._sldIdLst):
        rid = el.get(qn("r:id"))
        if prs.part.related_part(rid).slide is slide:
            prs.slides._sldIdLst.remove(el)
            prs.part.drop_rel(rid)


def reorder(prs, order):
    lst = prs.slides._sldIdLst
    by = {}
    for el in lst:
        by[id(prs.part.related_part(el.get(qn("r:id"))).slide)] = el
    els = [by[id(s)] for s in order if id(s) in by]
    rest = [el for el in lst if el not in els]
    for el in list(lst):
        lst.remove(el)
    for el in els + rest:
        lst.append(el)


def R(slide, old, new):
    try:
        replace_in_slide(slide, old, new)
    except ValueError:
        pass


# ───────────────────────── 계산 ─────────────────────────
def compute(F, I, spec):
    """장표 문구에 쓰는 값 계산."""
    tenors = sorted(F["split"])
    issue = F["issue"]
    lad = {int(k): v for k, v in F["lad"].items()}
    add = {issue.year + t: F["split"][t] for t in tenors}
    prev = F["prev"]
    pdm = {int(round(r["_tenor"])): r for r in prev if r.get("_tenor")}
    pri = int(I.get("priority") or tenors[-1])
    if pri not in tenors:
        pri = tenors[-1]
    # 옵션 B: 가장 짧은 만기로 증액분을 몰아주는 대안
    inc = max(F["max"] - F["base"], 0)
    a_max = dict(F["split"])
    a_max[pri] += inc
    b_max = dict(F["split"])
    b_max[tenors[0]] += inc
    # 시장 수요
    by_t, avg = F["by_t"], F["avg"]
    mkt = {}
    for t in tenors:
        k = "2년" if t <= 2 else "3년" if t == 3 else "5년 이상"
        mkt[t] = (avg.get(k), sum(by_t.get(k, {}).values()) if k in by_t else None, k)
    # 이벤트 (오늘 ~ 납입 후 30일)
    lo, hi = F["today"].isoformat(), (issue + timedelta(days=31)).isoformat()
    evs = [(d, nm) for d, nm in sorted(spec["events"].items()) if lo <= d <= hi]
    return SimpleNamespace(tenors=tenors, lad=lad, add=add, pdm=pdm, pri=pri, inc=inc, a_max=a_max, b_max=b_max, mkt=mkt, evs=evs)


def rates_now(rt, key, d=None):
    s = rt.get(key) or []
    if not s:
        return None, None
    if d is None:
        return s[-1][0], s[-1][1]
    v = series_at(s, d)
    return d, v


# ───────────────────────── 이야기 초안 (선택: Claude API) ─────────────────────────
def draft_story(F, I):
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        return None
    try:
        import anthropic
        co = F.get("co") or {}
        prev = [{"series": r.get("series_raw"), "tenor": r.get("_tenor"), "ratio": (r.get("_dem") or {}).get("ratio"),
                 "spread_bp": r.get("_spread"), "amount_eok": r["amount"] / 1e8} for r in F["prev"]]
        facts = {"회사": F["name"], "신용등급": F["grade"], "업종": co.get("industry"), "사업": (co.get("business") or "")[:1200],
                 "재무지표": F.get("metrics"), "직전 회차": prev, "향후 12개월 공모채 만기(억원)": round((F["due12"] or 0) / 1e8),
                 "최근 이슈 메모": I.get("issues_note"), "크레딧 포인트 메모": I.get("credit_note")}
        prompt = ("아래 회사 자료만 근거로 공모사채 발행 제안서의 '크레딧 스토리 4개'와 '투자자 예상 질문·답변 5개'를 한국어로 써라. "
                  "자료에 없는 숫자·사실은 지어내지 말고 꼭 필요하면 ## 로 남겨라. 각 항목 70자 이내. "
                  'JSON만 출력: {"points": [["제목","내용"],...4], "qa": [["질문","답변 방향"],...5]}\n\n'
                  + json.dumps(facts, ensure_ascii=False, default=str))
        msg = anthropic.Anthropic(api_key=key).messages.create(
            model=os.environ.get("PROPOSAL_MODEL", "claude-sonnet-5-5"), max_tokens=1500,
            messages=[{"role": "user", "content": prompt}])
        txt = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
        out = json.loads(txt[txt.index("{"): txt.rindex("}") + 1])
        if len(out.get("points", [])) >= 4 and len(out.get("qa", [])) >= 5:
            return out
    except Exception as e:  # noqa: BLE001
        print("story draft skipped:", type(e).__name__)  # 내용은 로그에 남기지 않음
    return None


# ───────────────────────── 변환 ─────────────────────────
def transform(v3_path, F, I, spec, out_path):
    prs = Presentation(v3_path)
    PR = Proto(prs)
    L = {l.name: l for l in prs.slide_layouts}
    LAY_I, LAY_II, LAY_III, LAY_SUM = L["2. 발행 조건 제안"], L["1_2. 발행 조건 제안"], L["3. 세일즈 전략"], L["1. 서머리"]
    S = list(prs.slides)
    cover, disc, divI = S[0], S[1], S[2]
    end = S[-1]
    cond, sched, tranche = find_slide(prs, "주요 발행 조건"), find_slide(prs, "주요 발행 일정"), find_slide(prs, "트랜치")
    mpc, fed, prev_s = find_slide(prs, "금융통화위원회"), find_slide(prs, "연준"), find_slide(prs, "前발행")
    credit, kr, us = find_slide(prs, "크레딧 시장"), find_slide(prs, "한국 채권시장"), find_slide(prs, "미국 채권시장")
    mat_next = find_slide(prs, "회사채 만기도래 현황")
    months = [s for s in S if any(sh.has_text_frame and sh.name.startswith("텍스트 개체 틀") and "만기도래 내역" in sh.text_frame.text for sh in s.shapes)]
    aa = find_slide(prs, "공모사채 발행내역")
    up = find_slide(prs, "향후 공모사채")
    blanks = [s for s in S[3:-1] if s.slide_layout.name == "빈 화면"]
    divII = next((s for s in blanks if S.index(s) > S.index(fed)), None)
    divIII = next((s for s in blanks if S.index(s) > S.index(prev_s)), None)

    C = compute(F, I, spec)
    name, grade, issue, sch = F["name"], F["grade"], F["issue"], F["sch"]
    tenors, split = C.tenors, F["split"]
    rt = F["rates"]
    on = spec.get("slides") or {}
    dflt = spec.get("defaults") or {}
    band = (I.get("band") or "").strip() or F["band"]
    band_note = "" if (I.get("band") or "").strip() else dflt.get("band_note", "직전 회차와 동일")
    fee = f"{I['fee_bp']}bp" if str(I.get("fee_bp") or "").strip() else "##bp"
    tenor_txt = " / ".join(f"{t}년 {num(split[t])}억" for t in tenors)
    month_word = f"{issue.month}월"

    # ── 표지·Disclaimer·목차 ──
    R(cover, f"㈜{name}를 위한", f"㈜{name}{josa(name, '을', '를')} 위한")
    for sh in cover.shapes:
        if sh.has_text_frame and "위한" in sh.text_frame.text:
            # 가장 긴 줄이 박스 폭에 들어가도록 글자 크기를 줄인다 (한글 1자 ≈ 글자 크기 1em)
            n = max(len(p.text) for p in sh.text_frame.paragraphs)
            fit = int(sh.width / 12700 / (n * 1.02) * 100)
            for r_ in sh._element.iter(qn("a:rPr")):
                if r_.get("sz") and int(r_.get("sz")) > fit:
                    r_.set("sz", str(max(fit, 2000)))
    for sh in disc.shapes:
        if sh.has_text_frame and "COPYRIGHT" in sh.text_frame.text:
            R(disc, re.search(r"COPYRIGHT©\d{4}", sh.text_frame.text).group(0) if re.search(r"COPYRIGHT©\d{4}", sh.text_frame.text) else "COPYRIGHT©", f"COPYRIGHT©{F['today'].year}")
    divA = dup_slide(prs, divIII)
    for d_, hl in ((divI, None), (divII, None), (divIII, None), (divA, 3)):
        add_app_item(d_, hl)

    # ── Executive Summary ──
    X = context(F, I, spec)
    summ = build_summary(prs, LAY_SUM, PR, F, I, X, C, spec, band, band_note, fee) if on.get("exec_summary", True) else None

    # ── 발행조건 ──
    build_cond(cond, F, I, C, spec, band, band_note, fee, X)

    # ── 예상 발행금리·프라이싱 ──
    rate = build_pricing(prs, LAY_I, PR, F, C, X) if on.get("pricing", True) else None
    # ── 만기구조 ──
    funds = build_maturity(prs, LAY_I, PR, F, C) if on.get("maturity", True) else None
    # ── 트랜치 ──
    build_tranche(tranche, PR, F, C, X)
    # ── 옵션 ──
    opt = build_options(prs, LAY_I, PR, F, C, X) if on.get("options", True) else None
    # ── 일정 ──
    build_schedule(sched, F, C, X)
    # ── 직전 회차 ──
    build_prev(prev_s, PR, F)
    # ── 크레딧 스토리 ──
    story = build_story(prs, LAY_II, PR, F, I, X) if on.get("story", True) else None
    # ── 시장: 대시보드·전망·우측 패널 ──
    dash = build_dashboard(prs, LAY_III, PR, F, X) if on.get("dashboard", True) else None
    outlook = build_outlook(prs, LAY_III, PR, F, X) if on.get("outlook", True) else None
    relayout(mpc, LAY_III)
    relayout(fed, LAY_III)
    R(fed, "Wiliams", "Williams")
    R(fed, "Washburm", "Washburn")
    R(mpc, "대응와", "대응과")
    rewrite_credit(credit, PR, F, X, spec)
    rewrite_kr(kr, F, X, spec)
    rewrite_us(us, F, X, spec)
    panel_credit(credit, PR, F, X)
    panel_kr(kr, PR, F, X)
    panel_us(us, PR, F, spec, X)
    strip_invisible_outlines(prs)

    order = [cover, disc, summ, divI, cond, rate, funds, tranche, opt, sched, divII, prev_s, story, divIII, dash, outlook,
             credit, kr, mpc, us, fed, mat_next, divA] + months + [aa, up, end]
    order = [s for s in order if s is not None]
    keep = {id(s) for s in order}
    for s in list(prs.slides):
        if id(s) not in keep:
            remove_slide(prs, s)
    reorder(prs, order)
    prs.save(out_path)
    return out_path


def style_line(c, cols):
    c.has_legend = True
    c.legend.position = XL_LEGEND_POSITION.TOP
    c.legend.include_in_layout = False
    c.legend.font.size = Pt(8)
    for s_, col in zip(c.plots[0].series, cols):
        s_.format.line.color.rgb = RGBColor.from_string(col)
        s_.format.line.width = Pt(1.75)
        s_.smooth = False
        s_.marker.style = XL_MARKER_STYLE.NONE
    c.value_axis.tick_labels.font.size = Pt(8)
    c.category_axis.tick_labels.font.size = Pt(8)
    vals = [v for sr in c.plots[0].series for v in sr.values if v is not None]
    if vals:
        c.value_axis.minimum_scale = round(int(min(vals) * 4) / 4 - 0.25, 2)
        c.value_axis.maximum_scale = round(int(max(vals) * 4) / 4 + 0.5, 2)
    c.value_axis.tick_labels.number_format = "0.0"
    c.value_axis.tick_labels.number_format_is_linked = False
    c.value_axis.major_gridlines.format.line.color.rgb = RGBColor.from_string("D9D9D9")
    c.value_axis.format.line.fill.background()
    c.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    c.font.name = BODY_FONT


def add_app_item(slide, highlight=None):
    sp = slide.shapes._spTree
    ell = [sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip() == "III" and "타원" in sh.name]
    lab = [sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip() == "시장 동향"]
    if not ell or not lab or any(sh.has_text_frame and sh.text_frame.text.strip() == "Appendix" for sh in slide.shapes):
        return
    dy = 868938
    for src, txt in ((ell[0], "※"), (lab[0], "Appendix")):
        el = copy.deepcopy(src._element)
        for c in el.iter(qn("p:cNvPr")):
            c.set("id", str(max(int(e.get("id")) for e in sp.iter(qn("p:cNvPr"))) + 1))
            for ext in c.findall(qn("a:extLst")):
                c.remove(ext)
        off = el.find(".//" + qn("a:off"))
        off.set("y", str(int(off.get("y")) + dy))
        sp.append(el)
        set_shape_text(el, txt)
        if txt == "※":
            sf = el.find(".//" + qn("a:solidFill"))
            for ch in list(sf):
                sf.remove(ch)
            etree.SubElement(sf, qn("a:schemeClr")).set("val", "tx1")
    if highlight is not None:
        ells = sorted([sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip() in ("I", "II", "III", "※")
                       and sh.width < 700000 and sh.top > 2000000], key=lambda s: s.top)
        for i, e in enumerate(ells):
            sf = e._element.spPr.find(qn("a:solidFill"))
            if sf is None:
                continue
            for ch in list(sf):
                sf.remove(ch)
            if i == highlight:
                etree.SubElement(sf, qn("a:srgbClr")).set("val", NAVY)
            else:
                etree.SubElement(sf, qn("a:schemeClr")).set("val", "tx1")


def strip_invisible_outlines(prs):
    for sl in prs.slides:
        for tag in ("a:rPr", "a:endParaRPr", "a:defRPr"):
            for r_ in sl._element.iter(qn(tag)):
                ln = r_.find(qn("a:ln"))
                if ln is not None:
                    al = ln.find(".//" + qn("a:alpha"))
                    if al is not None and al.get("val") == "0":
                        r_.remove(ln)


B = lambda t: {"text": t, "lvl": 0}  # noqa: E731


GR = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-"]


def same_band(a, b):
    if not a or not b or a not in GR or b not in GR:
        return False
    return abs(GR.index(a) - GR.index(b)) <= 1


def build_maturity(prs, LAY, PR, F, C):
    s = new_slide(prs, LAY)
    title(s, PR, "사채 만기구조 및 발행 후 변화")
    y0 = F["issue"].year
    years = list(range(y0, y0 + max(C.tenors) + 1))
    base = [C.lad.get(y, 0) for y in years]
    add = [C.add.get(y, 0) for y in years]
    after = [a + b for a, b in zip(base, add)]
    mx = max(range(len(years)), key=lambda i: after[i])
    H(s, f"'{years[mx] % 100}년 만기가 가장 많음({num(after[mx])}억) — 신규 만기는 여유 있는 연도로 분산")
    bar(s, PR, "공모사채 연도별 만기 (억원)", X0, 1800000, CW)
    cd = CategoryChartData()
    cd.categories = [f"'{y % 100}" for y in years]
    cd.add_series("기존", base)
    cd.add_series("이번 발행", add)
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Emu(X0), Emu(2110000), Emu(CW), Emu(4300000), cd).chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(9)
    for s_, col in zip(ch.plots[0].series, [NAVY, SKY]):
        s_.format.fill.solid()
        s_.format.fill.fore_color.rgb = RGBColor.from_string(col)
    pl = ch.plots[0]
    pl.gap_width = 60
    pl.has_data_labels = True
    pl.data_labels.font.size = Pt(8.5)
    pl.data_labels.number_format = "#,##0;;"
    pl.data_labels.number_format_is_linked = False
    pl.data_labels.position = XL_LABEL_POSITION.CENTER
    pl.data_labels.font.color.rgb = RGBColor.from_string("FFFFFF")
    ch.value_axis.major_gridlines.format.line.color.rgb = RGBColor.from_string("D9D9D9")
    ch.value_axis.tick_labels.font.size = Pt(9)
    ch.category_axis.tick_labels.font.size = Pt(10)
    ch.value_axis.format.line.fill.background()
    ch.font.name = BODY_FONT
    bar(s, PR, "발행 전후 연도별 만기 비교 (억원)", X2, 1800000, CW)
    rows = [["연도", "기존", "이번 발행", "발행 후"]]
    inv = {y0 + t: t for t in C.tenors}
    for y, b_, a_, f_ in zip(years, base, add, after):
        rows.append([f"'{y % 100}", num(b_) if b_ else "-", f"{num(a_)} ({inv[y]}년)" if a_ else "-", num(f_) if f_ else "-"])
    table(s, rows, X2, 2110000, CW, [0.8, 1.1, 1.1, 1.0], rowh=[300000] + [330000] * len(years), size=9.5,
          align_cols={i: PP_ALIGN.CENTER for i in range(1, 4)}, fills={(i + 1, 2): PALEBLUE for i, a_ in enumerate(add) if a_},
          bold_cells={(i, 3) for i in range(1, len(years) + 1)})
    lo = min((y for y in years if y > y0), key=lambda y: C.lad.get(y, 0))
    textbox(s, X2, 2110000 + 300000 + 330000 * len(years) + 200000, CW, 1500000, [
        {"text": "시사점", "bold": True, "color": NAVY, "size": 11},
        {"text": f"'{years[mx] % 100}년 발행 후 만기 {num(after[mx])}억원으로 최대 — 차환 계획을 가장 먼저 세울 해", "lvl": 0},
        {"text": f"'{lo % 100}년 기존 만기 {num(C.lad.get(lo, 0))}억원으로 여유 — 증액분 배정 후보", "lvl": 0},
        {"text": f"향후 12개월 공모사채 만기 {num((F['due12'] or 0) / 1e8)}억원 차환 계획과 연계", "lvl": 0}],
        size=10, fill=PALE, margin=120000, space_after=4)
    source(s, "주) 공모사채 잔액 기준(채권기본정보·공시), 이번 발행은 기본 금액 기준 / 자료: 기업공시시스템, 금융위원회", w=9000000)
    notes(s, "[작성 가이드] 동사 공모사채 만기 사다리 + 이번 발행 반영 후 변화.")
    return s


def build_prev(sl, PR, F):
    prev = F["prev"]
    if not prev or not all(r.get("_dem") for r in prev):
        return build_prev_blank(sl, PR, F)
    keep = None
    for sh in list(sl.shapes):
        if sh.has_text_frame and "前발행" in sh.text_frame.text:
            keep = sh
            continue
        sh._element.getparent().remove(sh._element)
    name = F["name"]
    sn = prev[0]["series_raw"].split("-")[0]
    dd = (prev[0].get("_demand_dt") or prev[0].get("pay_dt") or "")
    set_shape_text(keep._element, f"前발행 수요예측 상세결과 ({name} {sn}회, {dd[2:].replace('-', '.')})")
    tot = sum(r["_dem"]["total"] for r in prev)
    init = sum(r["_dem"]["total"] / r["_dem"]["ratio"] for r in prev if r["_dem"].get("ratio"))
    fin = sum(r["amount"] for r in prev) / 1e8
    tl = [tname(r["_tenor"]) for r in prev]
    H(sl, f"{num(init)}억원 모집에 {num(tot)}억원 주문({tot / init:.1f}:1) — {num(fin)}억원 발행")
    bar(sl, PR, f"{name} {sn}회 발행개요", X0, 1800000, CW)
    sp = [f"{t} {r['_spread']:+d}bp" for t, r in zip(tl, prev) if r.get("_spread") is not None]
    rt_ = [f"{t} {r['_rate']:.3f}%" for t, r in zip(tl, prev) if r.get("_rate")]
    table(sl, [
        ["구분", "내용"],
        ["신용등급", F["grade"]],
        ["수요예측", f"{dd}, 모집 {num(init)}억원 → 주문 {num(tot)}억원 ({tot / init:.1f}:1)"],
        ["신고 → 확정", " / ".join(f"{t} {num(r['_dem']['total'] / r['_dem']['ratio'])} → {num(r['amount'] / 1e8)}" for t, r in zip(tl, prev)) + "억원"],
        ["가산금리", " / ".join(sp) or "-"],
        ["확정금리", " / ".join(rt_) or "##"],
        ["대표주관", ", ".join(prev[0].get("underwriters") or []) or prev[0].get("lead") or "-"],
    ], X0, 2110000, CW, [1.1, 3.6], rowh=[280000] + [360000] * 6, size=9)
    bar(sl, PR, "만기별 모집금액 vs 주문금액 (억원)", X2, 1800000, CW)
    cd = CategoryChartData()
    cd.categories = tl
    cd.add_series("모집", [round(r["_dem"]["total"] / r["_dem"]["ratio"]) for r in prev])
    cd.add_series("주문", [round(r["_dem"]["total"]) for r in prev])
    cp = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(X2), Emu(2110000), Emu(CW), Emu(2280000), cd).chart
    cp.has_legend = True
    cp.legend.position = XL_LEGEND_POSITION.TOP
    cp.legend.include_in_layout = False
    cp.legend.font.size = Pt(9)
    for s_, col in zip(cp.plots[0].series, [LGREY, NAVY]):
        s_.format.fill.solid()
        s_.format.fill.fore_color.rgb = RGBColor.from_string(col)
    cp.plots[0].gap_width = 80
    cp.plots[0].has_data_labels = True
    cp.plots[0].data_labels.font.size = Pt(9)
    cp.plots[0].data_labels.number_format = "#,##0"
    cp.plots[0].data_labels.number_format_is_linked = False
    cp.plots[0].data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    cp.value_axis.visible = False
    cp.value_axis.has_major_gridlines = False
    cp.category_axis.tick_labels.font.size = Pt(10)
    cp.font.name = BODY_FONT
    bar(sl, PR, f"{sn}회 수요예측 기관군별 참여 (억원)", X0, 4720000, CW)
    groups = [("운용사", ["운용사"]), ("증권사", ["증권사"]), ("연기금·은행·보험", ["연기금·은행·보험"]), ("기타·외국", ["기타", "외국(실적유)", "외국(실적무)"])]
    rows = [["기관군"] + tl + ["합계"]]
    for lab, ks in groups:
        vals = [sum(r["_dem"]["groups"].get(k, 0) for k in ks) for r in prev]
        rows.append([lab] + [num(v) if v else "-" for v in vals] + [num(sum(vals)) if sum(vals) else "-"])
    rows.append(["합계"] + [num(r["_dem"]["total"]) for r in prev] + [num(tot)])
    table(sl, rows, X0, 5030000, CW, [1.7] + [0.75] * len(prev) + [0.85], rowh=290000, size=8.5,
          align_cols={i: PP_ALIGN.CENTER for i in range(1, len(prev) + 2)}, fills={(5, c): PALEBLUE for c in range(len(prev) + 2)},
          bold_cells={(5, c) for c in range(1, len(prev) + 2)})
    best = max(prev, key=lambda r: r["_dem"]["ratio"] or 0)
    bigg = max(((g, sum(sum(r["_dem"]["groups"].get(k, 0) for k in ks) for r in prev)) for g, ks in groups), key=lambda x: x[1])
    bar(sl, PR, "이번 회차 시사점", X2, 4500000, CW)
    textbox(sl, X2, 4810000, CW, 1950000, [
        {"text": f"**{tname(best['_tenor'])}물 수요 우위**", "color": NAVY},
        B(f"{tname(best['_tenor'])}물 경쟁률 {best['_dem']['ratio']:.1f}:1로 가장 높음 — 물량 배분에 반영"),
        {"text": f"**{bigg[0]} 중심 참여**", "color": NAVY},
        B(f"전체 주문의 {bigg[1] / tot * 100:.0f}%({num(bigg[1])}억원) — 같은 투자자군 재공략"),
        {"text": "**동일 밴드 유지**", "color": NAVY},
        B("직전 확정 가산금리가 밴드 하단~중간 — 같은 밴드로 금리 경쟁력 확보")],
        size=9, fill=PALE, margin=100000, space_after=2)
    source(sl, "자료: 투자설명서(수요예측 결과), 기업공시시스템", y=7090000, w=4500000)
    notes(sl, "[작성 가이드] 직전 회차: 모집→주문→확정, 만기별 경쟁률·가산금리, 기관군별 참여, 시사점.")


def build_prev_blank(sl, PR, F):
    """수요예측 결과가 수집되지 않았을 때: 템플릿(다른 회사) 내용을 지우고 아는 값 + ## 로 다시 그린다."""
    prev, name = F["prev"], F["name"]
    keep = None
    for sh in list(sl.shapes):
        if sh.has_text_frame and "前발행" in sh.text_frame.text and keep is None:
            keep = sh
            continue
        sh._element.getparent().remove(sh._element)
    if prev:
        sn = (prev[0].get("series_raw") or "").split("-")[0]
        dd = prev[0].get("_demand_dt") or prev[0].get("pay_dt") or ""
        set_shape_text(keep._element, f"前발행 수요예측 상세결과 ({name} {sn}회, {dd[2:].replace('-', '.')})")
        tl = [tname(r["_tenor"]) if r.get("_tenor") else "-" for r in prev]
        fin = sum(r["amount"] for r in prev) / 1e8
        H(sl, f"## (직전 {sn}회 수요예측 결과 한 줄) — {num(fin)}억원 발행")
        rows = [["구분", "내용"], ["신용등급", F["grade"]], ["수요예측", f"{dd or '##'}, 모집 ##억원 → 주문 ##억원 (##:1)"],
                ["발행금액", " / ".join(f"{t} {num(r['amount'] / 1e8)}" for t, r in zip(tl, prev)) + "억원"],
                ["가산금리", " / ".join(f"{t} {r['_spread']:+d}bp" if r.get("_spread") is not None else f"{t} ##bp" for t, r in zip(tl, prev))],
                ["확정금리", " / ".join(f"{t} {r['_rate']:.3f}%" if r.get("_rate") else f"{t} ##%" for t, r in zip(tl, prev))],
                ["대표주관", ", ".join(prev[0].get("underwriters") or []) or prev[0].get("lead") or "##"]]
        sub = f"{sn}회"
    else:
        set_shape_text(keep._element, "前발행 수요예측 상세결과")
        H(sl, "최근 공모사채 발행 이력 없음 — 동일 등급 사례로 대체 (##)")
        tl = ["##"]
        rows = [["구분", "내용"], ["신용등급", F["grade"]], ["수요예측", "##"], ["발행금액", "##"], ["가산금리", "##"], ["확정금리", "##"], ["대표주관", "##"]]
        sub = "직전 회차"
    bar(sl, PR, f"{name} {sub} 발행개요", X0, 1800000, CW)
    table(sl, rows, X0, 2110000, CW, [1.1, 3.6], rowh=[280000] + [360000] * 6, size=9)
    bar(sl, PR, "만기별 모집금액 vs 주문금액 (억원)", X2, 1800000, CW)
    table(sl, [["구분"] + tl, ["모집"] + ["##"] * len(tl), ["주문"] + ["##"] * len(tl), ["경쟁률"] + ["##:1"] * len(tl)],
          X2, 2110000, CW, [1.2] + [1] * len(tl), rowh=360000, size=9, align_cols={i: PP_ALIGN.CENTER for i in range(1, len(tl) + 1)})
    bar(sl, PR, f"{sub} 수요예측 기관군별 참여 (억원)", X0, 4720000, CW)
    grp = ["운용사", "증권사", "연기금·은행·보험", "기타·외국", "합계"]
    table(sl, [["기관군"] + tl + ["합계"]] + [[g] + ["##"] * (len(tl) + 1) for g in grp], X0, 5030000, CW,
          [1.7] + [0.75] * len(tl) + [0.85], rowh=290000, size=8.5, align_cols={i: PP_ALIGN.CENTER for i in range(1, len(tl) + 2)})
    bar(sl, PR, "이번 회차 시사점", X2, 4500000, CW)
    textbox(sl, X2, 4810000, CW, 1950000, [{"text": "**## (수요 우위 만기)**", "color": NAVY}, B("##"),
                                           {"text": "**## (주요 참여 기관군)**", "color": NAVY}, B("##")], size=9, fill=PALE, margin=100000, space_after=2)
    source(sl, "자료: 투자설명서(수요예측 결과), 기업공시시스템", y=7090000, w=4500000)
    notes(sl, "[작성 가이드] 수요예측 결과가 아직 수집되지 않은 회차입니다. KIND/DART 투자설명서 '수요예측 참여 내역'에서 ## 를 채우세요.")


def build_outlook(prs, LAY, PR, F, X=None):
    O = getattr(X, "outlook", None) if X else None
    if not O:
        return build_outlook_blank(prs, LAY, PR, F, X)
    from tools import proposal_outlook as OL
    s = new_slide(prs, LAY)
    title(s, PR, "향후 금리 전망 (리서치센터)")
    today, issue, cv, g = F["today"], F["issue"], X.cv, X.grade
    cols = [(k, c) for k, c in enumerate(O["cols"]) if c[1] >= today][:7]
    vd = O["vintage"]
    g3, gsrc = OL.grade_path(O, g, 3, cv)
    k3 = OL.ktb_path(O, 3)
    lab_i, v_i = OL.at(O, g3, issue)
    lab_e, v_e = OL.at(O, g3, date(issue.year, 12, 31))
    now = cv.y(g, 3) if cv else None
    if v_i is not None and now is not None:
        d1 = (v_i - now) * 100
        trend = "상승 전망: 금리 확정을 앞당기는 편이 유리" if d1 > 5 else "하락 전망: 발행 시점 분산 여지" if d1 < -5 else "보합 전망"
        H(s, f"리서치({vd.month}/{vd.day}): {g} 3년 현재 {now:.2f}% → {lab_i} {v_i:.2f}%" + (f" → {lab_e} {v_e:.2f}%" if v_e is not None and lab_e != lab_i else "") + f" — {trend}")
    else:
        H(s, f"리서치({vd.month}/{vd.day}) 금리 전망")

    def cur(key):
        if key == "base":
            v = (F["rates"].get("base") or [[None, None]])[-1][1]
            return f"{v:.2f}%" if v is not None else "-"
        if key == "ffr":
            v = (F["rates"].get("ffr_hi") or [[None, None]])[-1][1]
            return f"{v:.2f}%" if v is not None else "-"
        if key == "ust10":
            v = (F["rates"].get("ust10") or [[None, None]])[-1][1]
            return f"{v:.2f}%" if v is not None else "-"
        if key == "cp":
            v = (F["rates"].get("cp91") or [[None, None]])[-1][1]
            return f"{v:.2f}%" if v is not None else "-"
        return "-"
    f2 = lambda v: f"{v:.2f}%" if v is not None else "-"  # noqa: E731
    rows = [["지표", "현재"] + [c[1][0] for c in cols]]
    pick = lambda path: [f2(path[k]) if path else "-" for k, _ in cols]  # noqa: E731
    rows.append(["한은 기준금리", cur("base")] + pick(OL.series(O, "KR", "base")))
    rows.append(["국고 3년", f2(cv.y("KTB", 3)) if cv else "-"] + pick(k3))
    rows.append(["국고 10년", f2(cv.y("KTB", 10)) if cv else "-"] + pick(OL.ktb_path(O, 10)))
    hl = []
    for t in X.tenors[:3]:
        gp, _ = OL.grade_path(O, g, t, cv)
        hl.append(len(rows))
        rows.append([f"{g} {t}년", f2(cv.y(g, t)) if cv else "-"] + pick(gp))
    if g3 and k3:
        hl.append(len(rows))
        rows.append([f"{g} 3년 스프레드", f"{cv.sp(g, 3):.0f}bp" if cv else "-"] + [f"{(g3[k] - k3[k]) * 100:.0f}bp" if g3[k] is not None and k3[k] is not None else "-" for k, _ in cols])
    rows.append(["美 기준금리(상단)", cur("ffr")] + pick(OL.series(O, "US", "base")))
    rows.append(["美 국채 10년", cur("ust10")] + pick(OL.series(O, "US", "UST", 10)))
    if OL.series(O, "CP", 91):
        rows.append(["CP 91일", cur("cp")] + pick(OL.series(O, "CP", 91)))
    ic = next((n for n, (k, c) in enumerate(cols) if c[1] >= issue), None)
    fills = {(r, c): PALEBLUE for r in hl for c in range(len(rows[0]))}
    if ic is not None:
        fills.update({(r, 2 + ic): "FCE4D6" for r in range(1, len(rows))})
    table(s, rows, X0, 1800000, W, [1.8, 1.0] + [1.0] * len(cols), rowh=270000, size=9,
          align_cols={c: PP_ALIGN.CENTER for c in range(1, len(rows[0]))}, fills=fills)
    y2 = 1800000 + 270000 * len(rows) + 150000
    bar(s, PR, f"{g} 3년·국고 3년 전망 경로", X0, y2, CW)
    if g3 and k3 and cv:
        cd = CategoryChartData()
        cd.categories = ["현재"] + [c[1][0] for c in cols]
        cd.add_series(f"{g} 3년", [round(cv.y(g, 3), 3)] + [round(g3[k], 3) if g3[k] is not None else None for k, _ in cols])
        cd.add_series("국고 3년", [round(cv.y("KTB", 3), 3)] + [round(k3[k], 3) if k3[k] is not None else None for k, _ in cols])
        ch = s.shapes.add_chart(XL_CHART_TYPE.LINE, Emu(X0), Emu(y2 + 300000), Emu(CW), Emu(6700000 - y2 - 300000), cd).chart
        style_line(ch, [NAVY, SKY])
    bar(s, PR, "리서치 코멘트 (요약)", X2, y2, CW)
    cm = O.get("comments") or {}
    paras = []
    for k, lab in (("credit", "크레딧"), ("kr", "국내 금리"), ("us", "미국")):
        if cm.get(k):
            paras.append({"text": lab, "bold": True, "color": NAVY, "size": 9})
            paras.append({"text": OL.first_sentences(cm[k], 120), "size": 8.5})
    if lab_i:
        paras.append({"text": f"발행 시사점: 납입({md(issue)})이 속한 {lab_i} {g} 3년 전망 {v_i:.2f}% ({(v_i - now) * 100:+.0f}bp vs 현재)" if now is not None else "", "bold": True, "color": RED, "size": 9})
    textbox(s, X2, y2 + 300000, CW, 6700000 - y2 - 300000, paras, size=8.5, fill=PALE, margin=90000, space_after=2)
    age = (today - vd).days
    source(s, f"자료: 신한투자증권 리서치센터 금리 전망({vd:%Y.%m.%d}, 시트 '{O['sheet']}'){' — 전망일 ' + str(age) + '일 경과, 갱신 필요' if age > 45 else ''} · "
              f"{g} 경로: {gsrc} · 현재값: 금융투자협회·ECOS·FRED", w=9600000)
    notes(s, "[작성 가이드] proposal-vault 의 research/ 폴더에 최신 전망 엑셀을 올리면 자동 반영(가장 앞 시트 = 최신). 납입 시점이 속한 열을 강조.")
    return s


def build_outlook_blank(prs, LAY, PR, F, X=None):
    s = new_slide(prs, LAY)
    title(s, PR, "향후 전망 시나리오 (3·6개월)")
    H(s, "향후 3·6개월 금리 전망: ##")
    rt = F["rates"]
    cur = lambda k, d=2: (f"{rt[k][-1][1]:.{d}f}%" if rt.get(k) else "##")  # noqa: E731
    table(s, [["지표", "현재", "3개월 후 기본", "3개월 후 낙관", "3개월 후 비관", "6개월 후 기본"],
              ["한은 기준금리", cur("base"), "##", "##", "##", "##"], ["국고 3년", cur("ktb3"), "##", "##", "##", "##"],
              [f"{X.grade} 3년", f"{X.cv.y(X.grade, 3):.2f}%" if X and X.cv else "##", "##", "##", "##", "##"], [f"{X.grade} 스프레드", f"{X.cv.sp(X.grade, 3):.0f}bp" if X and X.cv else "##", "##", "##", "##", "##"],
              ["美 FF 상단", cur("ffr_hi"), "##", "##", "##", "##"]], X0, 1800000, W, [1.4, 1.2, 1.4, 1.4, 1.4, 1.4], rowh=430000, size=10,
          align_cols={i: PP_ALIGN.CENTER for i in range(1, 6)}, fills={**{(r, 2): PALEBLUE for r in range(1, 6)}, **{(0, 3): SKY, (0, 4): RED}})
    for i, (k, col) in enumerate((("기본", NAVY), ("낙관", SKY), ("비관", RED))):
        x = X0 + i * (CW3 + 200000)
        rect(s, x, 4580000, CW3, 330000, fill=col)
        textbox(s, x, 4580000, CW3, 330000, [{"text": k, "align": PP_ALIGN.CENTER, "bold": True}], size=10.5, color="FFFFFF", anchor=MSO_ANCHOR.MIDDLE, margin=0)
        textbox(s, x, 4910000, CW3, 1100000, [B("확률: ##%"), B("전망 근거: ##"), B("전망 근거: ##"), {"text": "발행 시사점: ##", "lvl": 0, "bold": True, "color": col}],
                size=9.5, fill=PALE, margin=110000, space_after=6)
    source(s, "자료: 신한투자증권 리서치센터 (전망 기준일 ##)", w=9000000)
    notes(s, "[작성 가이드] 전망 엑셀(시트 '전망') 연동 예정. 연결 전까지 ## 수기 입력.")
    return s



def build_summary(prs, LAY, PR, F, I, X, C, spec, band, band_note, fee):
    name, issue, sch, tenors, split = F["name"], F["issue"], F["sch"], C.tenors, F["split"]
    dflt = spec.get("defaults") or {}
    s = new_slide(prs, LAY)
    title(s, PR, "권고안 요약")
    exp_txt = " / ".join(f"{t}년 {X.exp[t]:.2f}%" if X.exp[t] is not None else f"{t}년 ##%" for t in tenors)
    H(s, f"{issue.month}월 {'·'.join(str(t) for t in tenors)}년 {num(F['base'])}억 발행 권고 — 예상 금리 {exp_txt}")
    ag = (X.co.get("triggers") or {}).get("agencies") or []
    rating = " · ".join(f"{a.get('agency')} {a.get('rating', '')}" + (f"({a['outlook']})" if a.get("outlook") else "") for a in ag) or X.grade
    bar(s, PR, "권고안", X0, 1800000, CW)
    table(s, [
        ["항목", "제안"],
        ["발행금액", f"{num(F['base'])}억원 (증액 시 최대 {num(F['max'])}억원)"],
        ["만기 구성", " / ".join(f"{t}년 {num(split[t])}억" for t in tenors) + f" — 증액 시 {C.pri}년 우선" + (" (자동 배분)" if X.alloc else "")],
        ["금리밴드", f"{band}" + (f" ({band_note})" if band_note else "")],
        ["예상 발행금리", exp_txt + f" ({X.base[tenors[0]][1]} 기준)"],
        ["일정", f"{md(sch['신고'])} 신고 → {md(sch['수요예측'])} 수요예측 → {md(issue)} 납입"],
        ["신용등급", rating],
        ["인수·수수료", f"{num(F['base'])}억 {dflt.get('underwrite', '총액인수 확약')} / {fee}"],
    ], X0, 2110000, CW, [1.1, 3.6], rowh=300000, size=9)
    bar(s, PR, "권고 근거", X2, 1800000, CW)
    reasons = []
    if F["prev"] and all(r.get("_dem") for r in F["prev"]):
        tot = sum(r["_dem"]["total"] for r in F["prev"])
        init = sum(r["_dem"]["total"] / r["_dem"]["ratio"] for r in F["prev"] if r["_dem"].get("ratio"))
        sp = [r["_spread"] for r in F["prev"] if r.get("_spread") is not None]
        reasons.append(("수요", f"직전 회차 {num(init)}억 모집에 {num(tot)}억 주문({tot / init:.1f}:1)" + (f", 확정 가산금리 {min(sp):+d}~{max(sp):+d}bp" if sp and min(sp) != max(sp) else f", 확정 가산금리 {sp[0]:+d}bp" if sp else "")))
    else:
        b = X.dem.get("3년") or next(iter(X.dem.values()), None)
        if b:
            reasons.append(("수요", f"최근 6개월 {X.band}급 공모채 {sum(v['n'] for v in X.dem.values())}건, 주문 {M.eok(sum(v['orders'] for v in X.dem.values()))}원"
                            + (f", 3년물 평균 경쟁률 {b['avg_ratio']:.1f}:1" if b.get("avg_ratio") else "")))
    sl = spread_line_short(X)
    if sl:
        reasons.append(("금리", sl))
    aj = X.adj
    if aj:
        reasons.append(("동사 신용", f"장외 거래가 {X.grade} 등급민평 대비 {aj['bp']:+d}bp({aj['n']}건) — 등급민평보다 {'낮은' if aj['bp'] < 0 else '높은'} 금리 근거"))
    ev_txt = "·".join(f"{nm.split(' (')[0]}({date.fromisoformat(d).month}/{date.fromisoformat(d).day})" for d, nm in C.evs if d > sch["정정"].isoformat())[:60]
    reasons.append(("일정", (f"{ev_txt} 전 {md(sch['정정'])} 금리 확정" if ev_txt else f"{md(sch['정정'])} 정정신고로 금리 확정")
                    + (f", 발행 전후 2주 {X.band}급 만기 {num(X.wall)}억 재투자 수요" if X.wall else "")))
    for i, (k, v) in enumerate(reasons[:3]):
        y = 2110000 + i * 680000
        rect(s, X2, y, CW, 600000, fill=PALE)
        rect(s, X2 + 90000, y + 120000, 360000, 360000, fill=NAVY, shape=2)
        textbox(s, X2 + 90000, y + 120000, 360000, 360000, [{"text": str(i + 1), "align": PP_ALIGN.CENTER, "bold": True}],
                size=12, color="FFFFFF", anchor=MSO_ANCHOR.MIDDLE, margin=0)
        textbox(s, X2 + 560000, y + 30000, CW - 620000, 540000,
                [{"text": k, "bold": True, "color": NAVY, "size": 10.5}, {"text": v}], size=9, anchor=MSO_ANCHOR.MIDDLE, margin=0)
    bar(s, PR, "리스크 및 대응", X0, 4720000, CW)
    risks = []
    p = X.cv.pct_rank(X.grade, 3) if X.cv else None
    if p is not None and p >= X.rules["spread_high_pct"]:
        risks.append([f"{X.grade} 스프레드 1년 상단({p}%) — 추가 확대 시 금리 상승", "밴드 상단 이내 확정, 증액은 하단 주문 범위만"])
    else:
        risks.append([f"{X.grade} 스프레드 재확대 가능성", "수요예측 전 기관 사전 수요(IOI) 확인 후 밴드 확정"])
    under = sum(v.get("under", 0) for v in X.dem.values())
    if under:
        risks.append([f"최근 6개월 {X.band}급 미매각 {under}건", "만기별 최초 물량 보수적 배분, 인수단 확약"])
    elif X.pipe:
        risks.append([f"같은 주 {X.band}급 경합 {len(X.pipe)}건", "수요예측일 조정 또는 주문 선점 마케팅"])
    else:
        risks.append(["직전 회차 투자자 한도 소진", "한도 잔여 점검 + 신규 기관 발굴"])
    risks.append(["한·미 정책 이벤트 전후 금리 변동", "이벤트 전 금리 확정(정정신고)"])
    table(s, [["리스크", "대응"]] + risks, X0, 5030000, CW, [1.9, 2.4], rowh=400000, size=9)
    lab = X.base[tenors[0]][1]
    bar(s, PR, f"{X.grade} 3년 등급민평·국고 3년 추이 (최근 6개월)", X2, 4140000, CW)
    if X.cv:
        st = (date.fromisoformat(X.cv.last) - timedelta(days=183)).isoformat()
        ds = [d for d in X.cv.dates if d >= st]
        cd0 = CategoryChartData()
        cd0.categories = [d[5:].replace("-", "/") for d in ds]
        cd0.add_series(f"{X.grade} 3년 등급민평", [round(X.cv.y(X.grade, 3, d), 3) for d in ds])
        cols = [NAVY]
        if X.adj and not X.indiv:
            cd0.add_series(f"동사 추정 개별민평({X.adj['bp']:+d}bp)", [round(X.cv.y(X.grade, 3, d) + X.adj["bp"] / 100, 3) for d in ds])
            cols.append(RED)
        cd0.add_series("국고 3년", [round(X.cv.y("KTB", 3, d), 3) for d in ds])
        cols.append(SKY)
        c0 = s.shapes.add_chart(XL_CHART_TYPE.LINE, Emu(X2), Emu(4440000), Emu(CW), Emu(2050000), cd0).chart
        style_line(c0, cols)
    textbox(s, X2, 6500000, CW, 230000, [f"자료: 금융투자협회 채권시가평가수익률(평가사 평균), {X.cv.last if X.cv else ''} 기준 · 예상 금리 = {lab} + 목표 가산금리"],
            size=8, color="7F7F7F", margin=0)
    notes(s, "[작성 가이드] 권고안(숫자) / 근거 3개 / 리스크·대응 / 등급민평·국고채 추이. 예상 금리는 기준 민평 + 목표 가산금리(동사 직전 확정값과 같은 등급군 중앙값의 평균).")
    return s


def spread_line_short(X):
    cv = X.cv
    if not cv:
        return None
    g = X.grade
    sp, p, c1 = cv.sp(g, 3), cv.pct_rank(g, 3), cv.change(g, 3, 30, spread=True)
    lv = "1년 중 높은 편" if p is not None and p >= X.rules["spread_high_pct"] else "1년 중 낮은 편" if p is not None and p <= X.rules["spread_low_pct"] else "1년 중간 수준"
    return f"{g} 3년 {cv.y(g, 3):.3f}%, 국고 대비 {sp:.0f}bp({lv}, 1개월 {c1:+.0f}bp)" if c1 is not None else f"{g} 3년 국고 대비 {sp:.0f}bp({lv})"


def build_cond(cond, F, I, C, spec, band, band_note, fee, X):
    old = [sh for sh in cond.shapes if sh.has_table][0]
    tx, ty, tw = old.left, old.top, old.width
    old._element.getparent().remove(old._element)
    issue, sch = F["issue"], F["sch"]
    tenors, split = C.tenors, F["split"]
    dflt = spec.get("defaults") or {}
    amt = []
    if F["due12"]:
        amt.append(B(f"**[만기도래 사채]** 향후 12개월 공모사채 만기 {num(F['due12'] / 1e8)}억원 — 차환 재원 선제 확보"))
    if (I.get("issues_note") or "").strip():
        amt.append(B("**[자금 수요]** " + I["issues_note"].strip().splitlines()[0][:110]))
    elif F["capex_h"]:
        amt.append(B(f"**[자금 수요]** {F['period']} 설비투자 {num(F['capex_h'])}억원 vs 영업현금흐름 {num(F['ocf_h'])}억원"
                     + (" — 외부조달 필요" if F["capex_h"] > F["ocf_h"] else " — 투자 재원 보완")))
    tot6 = sum(v["orders"] for v in X.dem.values())
    if F["prev"] and all(r.get("_dem") for r in F["prev"]):
        tot = sum(r["_dem"]["total"] for r in F["prev"])
        amt.append(B(f"**[시장 수급]** 직전 회차 주문 {num(tot)}억원, 최근 6개월 {X.band}급 주문 {M.eok(tot6)}원 — {num(F['base'])}억 신고·최대 {num(F['max'])}억 증액 소화 가능"))
    elif tot6:
        amt.append(B(f"**[시장 수급]** 최근 6개월 {X.band}급 공모채 주문 {M.eok(tot6)}원 — {num(F['base'])}억 신고·최대 {num(F['max'])}억 증액 소화 가능"))
    mat = []
    sps = [f"{t}년 {X.cv.sp(X.grade, t):.0f}bp" for t in tenors] if X.cv else []
    if sps:
        tp = term_premium(X, tenors[0], tenors[-1])
        mat.append(B(f"**[금리]** {X.grade} 국고 대비 " + " / ".join(sps) + (f" — {tenors[0]}→{tenors[-1]}년 기간 프리미엄 {tp:.0f}bp" if tp is not None else "")))
    dem_bits = [f"{t}년 {C.pdm[t]['_dem']['ratio']:.1f}:1" for t in tenors if t in C.pdm and C.pdm[t].get("_dem")]
    mk_bits = [f"{M.tbucket(t)} {X.dem[M.tbucket(t)]['avg_ratio']:.1f}:1" for t in tenors if (X.dem.get(M.tbucket(t)) or {}).get("avg_ratio")]
    if dem_bits or mk_bits:
        mat.append(B("**[수요]** " + ("직전 회차 " + " / ".join(dem_bits) + "; " if dem_bits else "") + (f"{X.band}급 6개월 평균 " + " / ".join(dict.fromkeys(mk_bits)) if mk_bits else "")))
    mat.append(B("**[만기 분산]** " + " / ".join(f"'{(issue.year + t) % 100}년 기존 {num(C.lad.get(issue.year + t, 0))}억" for t in tenors) + f" — 증액 시 {C.pri}년 우선"))
    ag = (X.co.get("triggers") or {}).get("agencies") or []
    rating_detail = (" · ".join(f"{a.get('agency')} {a.get('rating', '')}" + (f"({a['outlook']}, {a['eval_date'][2:].replace('-', '.')})" if a.get("outlook") and a.get("eval_date") else f"({a['outlook']})" if a.get("outlook") else "") for a in ag)
                     or f"유효등급 {X.grade}")
    lo, hi = X.band_bp
    band_y = " / ".join(f"{t}년 {X.base[t][0] + lo / 100:.2f}~{X.base[t][0] + hi / 100:.2f}%" for t in tenors if X.base[t][0] is not None)
    exp_y = " / ".join(f"{t}년 {X.exp[t]:.2f}%" if X.exp[t] is not None else f"{t}년 ##%" for t in tenors)
    exp_d = " / ".join(f"{t}년 {X.base[t][0]:.3f}% {X.tgt[t]:+d}bp" for t in tenors if X.base[t][0] is not None)
    ev_bits = [f"{nm}({date.fromisoformat(d).month}/{date.fromisoformat(d).day})" for d, nm in C.evs if d >= sch["신고"].isoformat()][:4]
    rows = [
        ["구분", "조건", "상세 내용"],
        ["발행금액", f"{num(F['base'])}억원\n(최대 {num(F['max'])}억원)", amt],
        ["인수", f"{num(F['base'])}억원\n{dflt.get('underwrite', '총액인수 확약')}", [B("수요예측 결과와 무관하게 신고금액 전액 인수 — 조달 확실성 확보")]],
        ["발행 만기", " / ".join(f"{t}년 {num(split[t])}억" for t in tenors), mat],
        ["금리밴드", band, [B((band_note + " — " if band_note else "") + (f"금리 환산 {band_y}" if band_y else ""))]],
        ["예상 발행금리", exp_y, [B(f"{X.base[tenors[0]][1]} + 목표 가산금리: {exp_d}" if exp_d else "개별민평 확인 후 입력 ##")]],
        ["신용등급", X.grade, [B(rating_detail)]],
        ["주요 일정", f"수요예측 {md(sch['수요예측'])}\n발행 {md(issue)}", [
            B(f"{md(sch['신고'])} 증권신고서 → {md(sch['수요예측'])} 수요예측 → {md(sch['정정'])} 정정신고서(금리 확정) → {md(issue)} 납입"),
            B(("·".join(ev_bits) + " 일정 고려") if ev_bits else "주요 이벤트 없음")]],
        ["인수수수료", fee, [B(dflt.get("fee_note", "직전 회차와 동일"))]],
    ]
    table(cond, rows, tx, ty, tw, [1.0, 1.75, 7.0], rowh=[300000, 1020000, 340000, 1020000, 400000, 400000, 340000, 760000, 340000],
          size=9, align_cols={1: PP_ALIGN.CENTER}, name="ConditionTable")
    for sh in cond.shapes:
        if sh.has_text_frame and sh.text_frame.text.startswith("주) 상기 조건은"):
            sh.top = Emu(7060000)
    notes(cond, "[작성 가이드] 발행금액=만기도래 사채·자금 수요·시장 수급 / 만기=금리(등급민평 스프레드·기간 프리미엄)·수요·만기분산 / 예상 금리=기준 민평+목표 가산금리.")


def build_pricing(prs, LAY, PR, F, C, X):
    s = new_slide(prs, LAY)
    title(s, PR, "예상 발행금리 및 프라이싱 전략")
    tenors = C.tenors
    lo, hi = X.band_bp
    lab = X.base[tenors[0]][1]
    H(s, f"{lab} 기준 예상 확정금리 " + " / ".join(f"{t}년 {X.exp[t]:.2f}%" for t in tenors if X.exp[t] is not None))
    bar(s, PR, f"만기별 예상 확정금리 ({lab} 기준)", X0, 1800000, CW)
    rows = [["구분"] + [f"{t}년" for t in tenors],
            [f"{X.grade} 등급민평"] + [f"{X.cv.y(X.grade, t):.3f}%" if X.cv else "##" for t in tenors]]
    if lab != "등급민평":
        rows.append([lab] + [f"{X.base[t][0]:.3f}%" if X.base[t][0] is not None else "##" for t in tenors])
    rows += [[f"밴드 하단({lo:+d}bp)"] + [f"{X.base[t][0] + lo / 100:.3f}%" if X.base[t][0] is not None else "##" for t in tenors],
             ["목표(예상 확정)"] + [f"{X.tgt[t]:+d}bp → {X.exp[t]:.3f}%" if X.exp[t] is not None else "##" for t in tenors],
             [f"밴드 상단({hi:+d}bp)"] + [f"{X.base[t][0] + hi / 100:.3f}%" if X.base[t][0] is not None else "##" for t in tenors],
             ["국고 대비"] + [f"{(X.exp[t] - X.cv.y('KTB', t)) * 100:.0f}bp" if X.exp[t] is not None and X.cv else "-" for t in tenors]]
    ti = [r[0] for r in rows].index("목표(예상 확정)")
    table(s, rows, X0, 2110000, CW, [1.4] + [1.1] * len(tenors), rowh=320000, size=9.5,
          align_cols={i: PP_ALIGN.CENTER for i in range(1, len(tenors) + 1)}, fills={(ti, c): PALEBLUE for c in range(len(tenors) + 1)},
          bold_cells={(ti, c) for c in range(1, len(tenors) + 1)})
    mine = list(F["prev"])
    pick = mine + [r for r in X.comps if r not in mine][: max(0, 8 - len(mine))]
    bar(s, PR, f"최근 {X.grade} 내외 공모채 프라이싱 사례", X2, 1800000, CW)
    prow = [["발행사 (납입일)", "등급", "만기", "경쟁률", "가산금리", "확정금리"]]
    for r in pick[:8]:
        d = r.get("_dem") or {}
        prow.append([f"{r['corp_name']}{r.get('series_raw', '')} ({(r.get('pay_dt') or '')[5:].replace('-', '/')})", M.gkey(r.get("_grade")) or "-",
                     tname(r["_tenor"]) if r.get("_tenor") else "-", f"{d['ratio']:.1f}:1" if d.get("ratio") else "-",
                     f"{r['_spread']:+d}bp" if r.get("_spread") is not None else "-", f"{r['_rate']:.3f}%" if r.get("_rate") else "-"])
    table(s, prow, X2, 2110000, CW, [2.1, 0.55, 0.55, 0.75, 0.8, 0.8], rowh=290000, size=8,
          align_cols={i: PP_ALIGN.CENTER for i in range(1, 6)}, fills={(i, c): PALEBLUE for i in range(1, len(mine) + 1) for c in range(6)})
    bar(s, PR, "프라이싱 근거", X0, 4720000, W)
    meds = [f"{k} {v['sp_med']:+d}bp" for k, v in sorted(X.dem.items()) if v.get("sp_med") is not None]
    pc = [("같은 등급군 시장 수준", (f"최근 6개월 {X.band}급 확정 가산금리 중앙값 " + " / ".join(meds)) if meds else f"최근 {X.band}급 확정 사례 부족"),
          ("동사 평가", adj_line(X, F["name"]) or ("직전 회차 확정 " + " / ".join(f"{t}년 {C.pdm[t]['_spread']:+d}bp" for t in tenors if t in C.pdm and C.pdm[t].get("_spread") is not None)) or "직전 회차·장외거래 자료 없음 — 등급민평 기준"),
          ("목표 가산금리", " / ".join(f"{t}년 {X.tgt[t]:+d}bp" for t in tenors) + " — 동사 직전 확정값과 같은 등급군 중앙값의 평균, 밴드 안에서")]
    for i, (k, v) in enumerate(pc):
        x = X0 + i * (CW3 + 200000)
        textbox(s, x, 5030000, CW3, 1350000, [{"text": k, "bold": True, "color": NAVY, "size": 10.5}, {"text": v}],
                size=9, fill=PALEBLUE if i == 2 else PALE, margin=100000, anchor=MSO_ANCHOR.MIDDLE, space_after=4)
    source(s, f"자료: 금융투자협회(채권시가평가수익률 {X.cv.last if X.cv else ''}·장외 체결), 기업공시시스템", w=9000000)
    notes(s, "[작성 가이드] 기준 민평 우선순위: 팝업 입력 개별민평 > 장외거래 추정 개별민평 > 등급민평. 확정 전 최신 개별민평으로 갱신.")
    return s


def build_tranche(sl, PR, F, C, X):
    for sh in list(sl.shapes):
        if sh.has_table or (sh.has_text_frame and sh.name.startswith("TextBox") and sh.top > 4500000 and sh.left > 5000000):
            sh._element.getparent().remove(sh._element)
    for sh in sl.shapes:
        if sh.has_text_frame and sh.text_frame.text.startswith(("등급·만기별 수요와", "만기별 제안 근거")):
            set_shape_text(sh._element, "만기별 금리·수급 근거")
        if sh.name == "모서리가 둥근 직사각형 8":
            for r_ in sh._element.iter(qn("a:rPr")):
                r_.set("sz", "1200")
    for sh in sl.shapes:
        if sh.name == "TextBox 10" or (sh.has_text_frame and 1200000 < sh.top < 1700000 and sh.width > 8000000):
            set_shape_text(sh._element, (f"만기별 금액 자동 배분(수요·비용·만기 분산·전망 균형) — {C.pri}년물 증액 우선" if X.alloc else
                                         f"{C.pri}년물 중심 증액 권고 — {X.grade} 등급민평 기간 프리미엄·같은 등급군 수요·만기 분산 반영"))
            break
    issue = F["issue"]
    rows = [["구분", "발행 금액", "제안 근거"]]
    for i, t in enumerate(C.tenors):
        y = issue.year + t
        bits = []
        if X.cv:
            g_y, sp = X.cv.y(X.grade, t), X.cv.sp(X.grade, t)
            s1 = f"[금리] {X.grade} {t}년 등급민평 {g_y:.3f}%(국고 대비 {sp:.0f}bp)"
            if i > 0:
                tp = term_premium(X, C.tenors[i - 1], t)
                if tp is not None:
                    s1 += f", {C.tenors[i - 1]}년 대비 +{tp:.0f}bp"
            c1 = X.cv.change(X.grade, t, 30, spread=True)
            if c1 is not None:
                s1 += f", 스프레드 1개월 {c1:+.0f}bp"
            bits.append(s1 + ".")
        dl = demand_line(X, t)
        if dl:
            bits.append(f"[수급] {dl}.")
        r = C.pdm.get(t)
        if r and r.get("_dem"):
            bits.append(f"[동사] 직전 {t}년물 {num(r['_dem']['total'])}억 주문({r['_dem']['ratio']:.1f}:1)" + (f", 확정 {r['_spread']:+d}bp." if r.get("_spread") is not None else "."))
        bits.append(f"[만기] '{y % 100}년 기존 {num(C.lad.get(y, 0))}억 + {num(F['split'][t])}억 = {num(C.lad.get(y, 0) + F['split'][t])}억.")
        al = (X.alloc or {}).get("detail", {}).get(t) if X.alloc else None
        if al:
            tr = f"{X.alloc['target_ratio']:.0f}:1"
            if F['split'][t] > al['cap']:
                bits.append(f"[배분] 예상 주문 {num(al['orders'])}억 — 경쟁률 {tr} 기준 적정 {num(al['cap'])}억을 상회하므로 사전 수요 확보 필요, 배분 비중 {al['score'] * 100:.0f}%.")
            else:
                bits.append(f"[배분] 예상 주문 {num(al['orders'])}억 → 경쟁률 {tr} 기준 적정 {num(al['cap'])}억 이내, 배분 비중 {al['score'] * 100:.0f}%.")
        if t == C.pri:
            bits.append("→ 증액 우선 배정 권고.")
        rows.append([f"{t}년", f"{num(F['split'][t])}억원", " ".join(bits)])
    bottom = min((sh.top for sh in sl.shapes if sh.top > 3000000 and (getattr(sh, "has_chart", False) and sh.has_chart or sh.name == "모서리가 둥근 직사각형 8")), default=4000000)
    rh = max(420000, (bottom - 1838003 - 260000 - 60000) // len(C.tenors))
    table(sl, rows, 451083, 1838003, 9790601, [0.6, 0.9, 8.3], rowh=[260000] + [rh] * len(C.tenors), size=8,
          align_cols={1: PP_ALIGN.CENTER}, fills={(1 + C.tenors.index(C.pri), c): PALEBLUE for c in (1, 2)})
    rr = [["항목"] + [f"{t}년 ({num(F['split'][t])}억)" for t in C.tenors],
          [f"{X.grade} 등급민평"] + [f"{X.cv.y(X.grade, t):.3f}%" if X.cv else "-" for t in C.tenors],
          ["국고 대비"] + [f"{X.cv.sp(X.grade, t):.0f}bp" if X.cv else "-" for t in C.tenors],
          [f"{X.band}급 평균 경쟁률"] + [(f"{X.dem[M.tbucket(t)]['avg_ratio']:.1f}:1" if (X.dem.get(M.tbucket(t)) or {}).get("avg_ratio") else "-") for t in C.tenors],
          [f"{X.band}급 확정 가산(중앙)"] + [(f"{X.dem[M.tbucket(t)]['sp_med']:+d}bp" if (X.dem.get(M.tbucket(t)) or {}).get("sp_med") is not None else "-") for t in C.tenors],
          ["직전 회차 경쟁률"] + [(f"{C.pdm[t]['_dem']['ratio']:.1f}:1" if t in C.pdm and C.pdm[t].get("_dem") else "-") for t in C.tenors],
          ["만기 연도 기존"] + [f"'{(issue.year + t) % 100} {num(C.lad.get(issue.year + t, 0))}억" for t in C.tenors]]
    table(sl, rr, 5489407, 4716722, 4752279, [1.4] + [1.15] * len(C.tenors), rowh=[260000] + [300000] * 6, size=8,
          align_cols={i: PP_ALIGN.CENTER for i in range(1, len(C.tenors) + 1)})
    w = (X.alloc or {}).get("weights") or {}
    notes(sl, f"[작성 가이드] 만기별: 금리({X.grade} 등급민평·국고 대비 스프레드·기간 프리미엄) / 수급({X.band}급 6개월 수요) / 동사 직전 회차 / 만기 분산."
          + (f" 자동 배분 비중 — 수요 {w.get('demand')}, 비용 {w.get('cost')}, 만기 분산 {w.get('maturity')}, 전망 {w.get('outlook')} (명세 rules.alloc)." if X.alloc else ""))


def build_options(prs, LAY, PR, F, C, X):
    s = new_slide(prs, LAY)
    title(s, PR, "만기 전략 옵션 비교")
    t0, tp = C.tenors[0], C.pri
    y = F["issue"].year
    prem = term_premium(X, t0, tp) if tp != t0 else None
    exp_a = sum(X.exp[t] * C.a_max[t] for t in C.tenors if X.exp[t] is not None) / max(1, sum(C.a_max.values()))
    exp_b = sum(X.exp[t] * C.b_max[t] for t in C.tenors if X.exp[t] is not None) / max(1, sum(C.b_max.values()))

    def mat_line(sp):
        return " / ".join(f"'{(y + t) % 100} {num(C.lad.get(y + t, 0) + sp[t])}" for t in C.tenors)
    H(s, f"옵션 A({tp}년 우선 증액) 권고 — 평균 예상금리 {exp_a:.2f}%(옵션 B 대비 {(exp_a - exp_b) * 100:+.0f}bp)로 {tp}년 고정금리 확보")
    da, db = C.pdm.get(tp), C.pdm.get(t0)
    cards = [
        (f"옵션 A  ·  {tp}년 우선 증액 (권고)", NAVY, PALEBLUE,
         [["최대 구성", " / ".join(f"{t}년 {num(C.a_max[t])}" for t in C.tenors) + "억"], ["평균 예상금리", f"{exp_a:.2f}%"], ["만기 (억원)", mat_line(C.a_max)]],
         [f"{tp}년 고정금리 확보 — 차환 주기 분산",
          (f"직전 회차 {tp}년물 {da['_dem']['ratio']:.1f}:1 수요 확인" if da and da.get("_dem") else demand_line(X, tp) or "같은 등급군 수요 활용"),
          "신규 만기를 뒤로 분산"],
         [f"{t0}년 대비 기간 프리미엄 {prem:.0f}bp" if prem is not None else "기간 프리미엄만큼 금리 상승",
          f"'{(y + tp) % 100}년 만기 {num(C.lad.get(y + tp, 0) + C.a_max[tp])}억원으로 증가"],
         "만기 장기화와 금리 고정이 조달비용보다 우선일 때"),
        (f"옵션 B  ·  {t0}년 우선 증액", DGREY, PALE,
         [["최대 구성", " / ".join(f"{t}년 {num(C.b_max[t])}" for t in C.tenors) + "억"], ["평균 예상금리", f"{exp_b:.2f}%"], ["만기 (억원)", mat_line(C.b_max)]],
         ["단기물 비중이 높아 평균 조달금리 최저",
          (f"직전 회차 {t0}년물 {db['_dem']['ratio']:.1f}:1 수요" if db and db.get("_dem") else demand_line(X, t0) or "단기물 수요 흡수"),
          f"기간 프리미엄 {prem:.0f}bp 절감" if prem is not None else "기간 프리미엄 절감"],
         [f"'{(y + t0) % 100}년 만기 {num(C.lad.get(y + t0, 0) + C.b_max[t0])}억원으로 증가", "차환 시점 금리 재노출"],
         "단기 조달비용 최소화가 우선일 때"),
    ]
    for i, (hd, col, bg, kv, pros, cons, fit) in enumerate(cards):
        x = X0 if i == 0 else X2
        rect(s, x, 1800000, CW, 420000, fill=col)
        textbox(s, x, 1800000, CW, 420000, [{"text": hd, "align": PP_ALIGN.CENTER, "bold": True}], size=12, color="FFFFFF",
                anchor=MSO_ANCHOR.MIDDLE, margin=0)
        table(s, kv, x, 2290000, CW, [1.0, 2.6], rowh=300000, size=9.5, header_rows=0, align_cols={1: PP_ALIGN.CENTER},
              bold_cells={(0, 1), (1, 1), (2, 1)})
        textbox(s, x, 3270000, CW, 1250000, [{"text": "장점", "bold": True, "color": NAVY, "size": 10.5}] + [B(t) for t in pros],
                size=10, fill=bg, margin=110000, space_after=4)
        textbox(s, x, 4600000, CW, 1050000, [{"text": "단점", "bold": True, "color": RED, "size": 10.5}] + [B(t) for t in cons],
                size=10, fill="FBEAEA", margin=110000, space_after=4)
        textbox(s, x, 5720000, CW, 500000, [{"text": "**적합한 상황**  " + fit}], size=10, margin=60000, anchor=MSO_ANCHOR.MIDDLE)
    cut = 3.0
    textbox(s, X0, 6330000, W, 420000, [{"text": f"**신한 의견: 옵션 A** — {tp}년물 경쟁률 {cut:.0f}:1 미만이면 옵션 B로 전환", "align": PP_ALIGN.CENTER}],
            size=10.5, color=NAVY, fill=PALEBLUE, margin=60000, anchor=MSO_ANCHOR.MIDDLE)
    notes(s, "[작성 가이드] 평균 예상금리 = 만기별 예상 확정금리를 최대 구성 금액으로 가중평균.")
    return s


def build_schedule(sl, F, C, X):
    sch, issue = F["sch"], F["issue"]
    for sh in sl.shapes:
        if sh.has_text_frame and "발행일정 제안" in sh.text_frame.text:
            set_shape_text(sh._element, f"이벤트·같은 등급군 수급을 고려하여 {md(issue)} 발행일정 제안")
    tb = next((sh for sh in sl.shapes if sh.name == "TextBox 5"), None)
    before = [f"{nm.split(' (')[0]}({date.fromisoformat(d).month}/{date.fromisoformat(d).day})" for d, nm in C.evs if sch["정정"].isoformat() < d]
    pre = [f"{nm.split(' (')[0]}({date.fromisoformat(d).month}/{date.fromisoformat(d).day})" for d, nm in C.evs if F["today"].isoformat() <= d <= sch["정정"].isoformat()]
    lines = [{"text": f"**{md(sch['수요예측'])} 수요예측 → {md(sch['정정'])} 금리 확정 — 이벤트 리스크 차단**", "lvl": 0}]
    if before:
        lines.append({"text": "·".join(before[:4]) + " 이전 발행금리 확정", "lvl": 1})
    if pre:
        lines.append({"text": "수요예측 전 " + "·".join(pre[:3]) + " 결과를 밴드에 반영", "lvl": 1})
    lines.append({"text": f"**수급: {X.band}급 만기·발행 일정**", "lvl": 0})
    if X.month_wall:
        lines.append({"text": f"{issue.month}월 {X.band}급 공모채 만기 {num(X.month_wall)}억원" + (f"(수요예측~납입 전후 2주 {num(X.wall)}억)" if X.wall and abs(X.wall - X.month_wall) > 1 else "") + " — 상환 자금의 재투자 수요", "lvl": 1})
    if X.pipe:
        lines.append({"text": f"수요예측 ±1주 {X.band}급 경합 {len(X.pipe)}건(" + ", ".join(f"{n} {ymd(d)}" for n, d, _ in X.pipe[:3]) + ") — 주문 분산 가능성", "lvl": 1})
    else:
        lines.append({"text": f"수요예측 ±1주 현재 공시된 {X.band}급 경합 발행 없음 (신고서 제출 전 재확인)", "lvl": 1})
    n_, a_, r_ = X.ly
    if n_:
        lines.append({"text": f"작년 {issue.month}월 {X.band}급 {n_}건·{num(a_)}억 발행" + (f", 평균 경쟁률 {r_:.1f}:1" if r_ else ""), "lvl": 1})
    sl_ = spread_line_short(X)
    if sl_:
        lines.append({"text": f"**금리: {sl_}**", "lvl": 0})
    if getattr(X, "outlook", None) and X.cv:
        from tools import proposal_outlook as OL
        g3, _ = OL.grade_path(X.outlook, X.grade, 3, X.cv)
        lab_i, v_i = OL.at(X.outlook, g3, issue)
        lab_n, v_n = OL.at(X.outlook, g3, date(issue.year, 6, 30) if issue.month <= 3 else issue + timedelta(days=120))
        if v_i is not None:
            lines.append({"text": f"리서치 전망: {X.grade} 3년 {lab_i} {v_i:.2f}%" + (f" → {lab_n} {v_n:.2f}%" if v_n is not None and lab_n != lab_i else "")
                          + f" (현재 {X.cv.y(X.grade, 3):.2f}%)", "lvl": 1})
    lines.append({"text": f"**→ {md(sch['신고'])} 신고 · {md(sch['수요예측'])} 수요예측 · {md(issue)} 납입**", "lvl": 0})
    if tb is not None:
        fill_tf(tb.text_frame, lines, size=9, space_after=2)
    st = next((sh for sh in sl.shapes if sh.has_table), None)
    if st is None:
        return
    t = st.table
    for r in range(1, len(t.rows), 2):
        for c in range(1, len(t.columns)):
            cell = t.cell(r + 1, c) if r + 1 < len(t.rows) else None
            if cell is None:
                continue
            txt = cell.text_frame.text
            col = None
            if "정정" in txt:
                fill_tf(cell.text_frame, [{"text": "발행금리 확정", "align": PP_ALIGN.CENTER, "bold": True}, {"text": "(정정신고서 제출)", "align": PP_ALIGN.CENTER, "bold": True}],
                        size=9, color="C55A11", space_after=0)
                col = "FCE4D6"
            elif "발행 및 납입" in txt:
                fill_tf(cell.text_frame, [{"text": x, "align": PP_ALIGN.CENTER, "bold": True} for x in txt.split("\n")], size=9, color="FFFFFF", space_after=0)
                col = NAVY
            if col:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(col)
    notes(sl, "[작성 가이드] 달력형 일정표. 이벤트는 docs/proposal_spec.md 의 events, 수급은 같은 등급군 만기·발행 예정 자동 집계.")


def build_story(prs, LAY, PR, F, I, X):
    s = new_slide(prs, LAY)
    title(s, PR, "크레딧 스토리 및 투자자 예상 질문")
    pts, qa, msg = credit_story(F, X, I)
    dr = draft_story(F, I)
    if dr:
        pts = [tuple(x[:2]) for x in dr["points"][:4]]
        qa = [tuple(x[:2]) for x in dr["qa"][:5]]
    H(s, f"수요예측 메시지: {msg}")
    bar(s, PR, "투자 포인트 (크레딧 스토리)", X0, 1800000, CW)
    for i, (k, v) in enumerate(pts):
        y = 2110000 + i * 900000
        rect(s, X0, y, CW, 800000, fill=PALE)
        rect(s, X0 + 100000, y + 200000, 400000, 400000, fill=NAVY, shape=2)
        textbox(s, X0 + 100000, y + 200000, 400000, 400000, [{"text": str(i + 1), "align": PP_ALIGN.CENTER, "bold": True}],
                size=13, color="FFFFFF", anchor=MSO_ANCHOR.MIDDLE, margin=0)
        textbox(s, X0 + 620000, y + 40000, CW - 700000, 720000, [{"text": k, "bold": True, "color": NAVY, "size": 10.5}, {"text": v}],
                size=9, anchor=MSO_ANCHOR.MIDDLE, margin=0)
    bar(s, PR, "투자자 예상 질문 및 답변 방향", X2, 1800000, CW)
    table(s, [["예상 질문", "답변 방향"]] + [list(x) for x in qa], X2, 2110000, CW, [1.5, 2.7], rowh=[300000] + [640000] * 5, size=8.5)
    textbox(s, X2, 5700000, CW, 650000, [{"text": "투자자 배포 자료 원칙", "bold": True, "color": RED},
                                          {"text": "숫자는 공시 기준 — 일시적 요인은 대응 논리와 함께 설명"}],
            size=9, fill="FBEAEA", margin=110000, anchor=MSO_ANCHOR.MIDDLE)
    src = "자료: 사업보고서·분기보고서(DART), 신용평가사, 금융투자협회 장외 체결" + (", 팝업 메모" if (I.get("credit_note") or I.get("issues_note")) else "")
    source(s, src + (" · Claude 초안(검토 필요)" if dr else ""), w=9000000)
    notes(s, "[작성 가이드] 크레딧 스토리 4개(사업·실적·재무·신용도) + 예상 Q&A 5개는 공시 수치로 자동 작성. 팝업 '크레딧 포인트' 메모가 있으면 맨 앞에 반영.")
    return s


def build_dashboard(prs, LAY, PR, F, X):
    s = new_slide(prs, LAY)
    title(s, PR, "시장 한눈에 보기 (Market Dashboard)")
    rt = F["rates"]
    cv = X.cv

    def row(lab, key, unit="%", scale=1, dec=3):
        s_ = rt.get(key) or []
        if not s_:
            return [lab, "-", "-", "-", "-"]
        d_now, v = s_[-1]
        dn = date.fromisoformat(d_now)
        m1 = series_at(s_, (dn - timedelta(days=31)).isoformat())
        y1 = series_at(s_, (dn - timedelta(days=365)).isoformat())
        f = lambda x: f"{x * scale:.{dec}f}{unit}" if x is not None else "-"  # noqa: E731
        return [lab, f(v), f(m1), f(y1), f"{(v - m1) * 100:+.0f}bp" if m1 is not None else "-"]

    def crow(lab, g, t, spread=False):
        if not cv:
            return [lab, "-", "-", "-", "-"]
        d1 = cv.last
        dm = cv.asof((date.fromisoformat(d1) - timedelta(days=31)).isoformat())
        dy = cv.asof((date.fromisoformat(d1) - timedelta(days=365)).isoformat())
        f = (lambda d: cv.sp(g, t, d)) if spread else (lambda d: cv.y(g, t, d))
        fm = (lambda v: f"{v:.0f}bp") if spread else (lambda v: f"{v:.3f}%")
        a, b, c = f(d1), f(dm) if dm else None, f(dy) if dy else None
        ch = (a - b) if spread else (a - b) * 100 if a is not None and b is not None else None
        return [lab, fm(a) if a is not None else "-", fm(b) if b is not None else "-", fm(c) if c is not None else "-", f"{ch:+.0f}bp" if ch is not None else "-"]
    H(s, (spread_line_short(X) or "") if cv else "금리 레벨·스프레드 점검")
    rows = [["지표", "현재", "1개월 전", "1년 전", "1개월 변화"],
            row("한은 기준금리", "base", dec=2), crow("국고 3년", "KTB", 3), crow("국고 10년", "KTB", 10),
            crow(f"{X.grade} 3년 (동사 등급)", X.grade, 3), crow(f"{X.grade} 3년 스프레드", X.grade, 3, spread=True)]
    if X.grade != "AA-":
        rows.append(crow("AA- 3년 스프레드 (참고)", "AA-", 3, spread=True))
    rows += [row("CP 91일", "cp91", dec=2), row("美 FF 상단", "ffr_hi", dec=2), row("美 국채 2년", "ust2", dec=2), row("美 국채 10년", "ust10", dec=2)]
    table(s, rows, X0, 1800000, W, [2.4, 1.5, 1.5, 1.5, 1.5], rowh=330000, size=10,
          align_cols={i: PP_ALIGN.CENTER for i in range(1, 5)}, bold_cells={(r, 1) for r in range(1, len(rows))},
          fills={(r, c): PALEBLUE for r in (4, 5) for c in range(5)})
    imp = []
    if cv:
        p = cv.pct_rank(X.grade, 3)
        c1 = cv.change(X.grade, 3, 30, spread=True)
        k1 = cv.change("KTB", 3, 30)
        imp.append(f"{X.grade} 스프레드 1년 분위 {p}%" + (f", 1개월 {c1:+.0f}bp" if c1 is not None else "") + (f" / 국고 3년 1개월 {k1:+.0f}bp" if k1 is not None else ""))
        if c1 is not None and k1 is not None:
            imp.append("금리·스프레드 동반 하락 — 발행 여건 우호" if c1 <= 0 and k1 <= 0 else
                       "국고 상승을 스프레드 축소가 일부 상쇄" if c1 < 0 < k1 else
                       "스프레드 확대 — 밴드 상단 관리 필요" if c1 > 0 >= k1 else "금리·스프레드 동반 상승 — 조기 확정 유리")
    textbox(s, X0, 5850000, W, 600000, [{"text": "발행 영향: " + " — ".join(imp) if imp else "발행 영향: ##"}], size=10, fill=PALEBLUE,
            margin=110000, anchor=MSO_ANCHOR.MIDDLE)
    source(s, f"자료: 금융투자협회 채권시가평가수익률({cv.last if cv else ''}), 한국은행 ECOS, FRED", w=9000000)
    notes(s, "[작성 가이드] 국고·회사채는 금융투자협회 시가평가(평가사 평균), 기준금리·CP는 ECOS, 미국은 FRED. 모두 자동.")
    return s


def rewrite_credit(sl, PR, F, X, spec):
    cv = X.cv
    if not cv:
        return
    one = (date.fromisoformat(cv.last) - timedelta(days=365)).isoformat()
    ds = [d for d in cv.dates if d >= one]
    ch = next((sh for sh in sl.shapes if getattr(sh, "has_chart", False) and sh.has_chart), None)
    if ch is not None:
        cd = CategoryChartData()
        cd.categories = [date.fromisoformat(d) for d in ds]
        cd.add_series("SP", [round(cv.sp(X.grade, 3, d), 1) for d in ds])
        cd.add_series("국고 3년", [round(cv.y("KTB", 3, d), 3) for d in ds])
        cd.add_series(f"{X.grade} 3년", [round(cv.y(X.grade, 3, d), 3) for d in ds])
        try:
            ch.chart.replace_data(cd)
        except Exception:  # noqa: BLE001
            pass
    for sh in sl.shapes:
        if sh.name == "모서리가 둥근 직사각형 8" and "추이" in (sh.text_frame.text if sh.has_text_frame else ""):
            set_shape_text(sh._element, f"최근 1년 금리 및 스프레드 추이 ({X.grade} 3년)")
        if sh.has_text_frame and sh.text_frame.text.startswith("자료: 한국은행"):
            set_shape_text(sh._element, "자료: 금융투자협회 채권시가평가수익률(평가사 평균)")
    rng = cv.range(X.grade, 3)
    c3 = cv.change(X.grade, 3, 91, spread=True)
    L = [f"{X.grade} 3년 스프레드 {cv.sp(X.grade, 3):.0f}bp — 1년 범위 {rng[0][1]:.0f}bp({rng[0][0][2:7].replace('-', '.')})~{rng[1][1]:.0f}bp({rng[1][0][2:7].replace('-', '.')}) 중 {cv.pct_rank(X.grade, 3)}% 위치"]
    sub = []
    if c3 is not None:
        sub.append(f"3개월 {c3:+.0f}bp")
    if X.grade != "AA-":
        d_now = cv.sp(X.grade, 3) - cv.sp("AA-", 3)
        d_old = cv.sp(X.grade, 3, ds[0]) - cv.sp("AA-", 3, ds[0])
        sub.append(f"AA- 대비 {d_now:+.0f}bp(1년 전 {d_old:+.0f}bp) — 등급 간 차별화 {'확대' if d_now > d_old else '축소'}")
    if sub:
        L.append("- " + ", ".join(sub))
    tot = sum(v["n"] for v in X.dem.values())
    if tot:
        under = sum(v.get("under", 0) for v in X.dem.values())
        L.append(f"최근 6개월 {X.band}급 공모채 {tot}건, 주문 {M.eok(sum(v['orders'] for v in X.dem.values()))}원" + (f", 미매각 {under}건" if under else ", 미매각 없음"))
        L.append("- " + " / ".join(f"{k} 평균 {v['avg_ratio']:.1f}:1" for k, v in sorted(X.dem.items()) if v.get("avg_ratio")))
    aj = adj_line(X, F["name"])
    if aj:
        L.append(f"동사: {aj}")
    L = (L[:5] + fresh_view(spec, "credit", F["today"], X.rules))[:7]
    tb = next((sh for sh in sl.shapes if sh.name == "TextBox 17"), None)
    if tb is not None:
        P3.set_bul(tb, L)


def rewrite_kr(sl, F, X, spec):
    L, _ = kr_lines(F["rates"])
    L = (L + fresh_view(spec, "kr", F["today"], X.rules))[:6]
    tb = next((sh for sh in sl.shapes if sh.name == "TextBox 25"), None)
    if tb is not None and L:
        P3.set_bul(tb, L)


def rewrite_us(sl, F, X, spec):
    L, _ = us_lines(F["rates"])
    L = (L + fresh_view(spec, "us", F["today"], X.rules))[:6]
    tb = next((sh for sh in sl.shapes if sh.name == "TextBox 17"), None)
    if tb is not None and L:
        P3.set_bul(tb, L)


def _chart_top(slide):
    bars = [sh for sh in slide.shapes if sh.name == "모서리가 둥근 직사각형 8"]
    return min(b.top for b in bars) if bars else 3440000


def panel_credit(sl, PR, F, X):
    cv = X.cv
    if not cv:
        return
    end = cv.last
    st = (date.fromisoformat(end) - timedelta(days=365)).isoformat()
    ev = []
    bs = F["rates"].get("base") or []
    for (d0, v0), (d1, v1) in zip(bs, bs[1:]):
        if v1 != v0 and d1 >= st:
            ev.append((d1, f"한은 {'인상' if v1 > v0 else '인하'} {v0:.2f}→{v1:.2f}%"))
    rng = cv.range(X.grade, 3)
    if rng:
        ev.append((rng[1][0], f"{X.grade} 스프레드 1년 고점"))
        ev.append((rng[0][0], f"{X.grade} 스프레드 1년 저점"))
    ev.append((end, "현재"))
    ev = sorted(set(ev))[-6:]
    rows = [["일자", "이벤트", f"{X.grade} 3년", "스프레드"]]
    for d, nm in ev:
        rows.append([f"'{d[2:4]}.{int(d[5:7])}/{int(d[8:])}", nm, f"{cv.y(X.grade, 3, d):.3f}%", f"{cv.sp(X.grade, 3, d):.0f}bp"])
    y0 = _chart_top(sl)
    bar(sl, PR, f"주요 이벤트와 {X.grade} 3년 금리", RX, y0, RW, h=276999, size=12)
    table(sl, rows, RX, y0 + 330000, RW, [0.9, 2.1, 0.9, 0.9], rowh=300000, size=9,
          align_cols={2: PP_ALIGN.CENTER, 3: PP_ALIGN.CENTER}, fills={(len(rows) - 1, c): PALEBLUE for c in range(4)})
    p = cv.pct_rank(X.grade, 3)
    msg = ("스프레드가 1년 상단 — 금리 확정 시점 앞당김 유리" if p is not None and p >= X.rules["spread_high_pct"] else
           "스프레드가 1년 하단 — 우호적 발행 구간" if p is not None and p <= X.rules["spread_low_pct"] else "스프레드 1년 중간 — 수요예측 흥행 시 밴드 하단 확정 여지")
    textbox(sl, RX, y0 + 330000 + len(rows) * 300000 + 150000, RW, 800000,
            [{"text": "발행 시사점", "bold": True, "color": NAVY}, {"text": msg}], size=9.5, fill=PALEBLUE, margin=100000)


def panel_kr(sl, PR, F, X):
    rt = F["rates"]
    bs = rt.get("base") or []
    k3, k10 = rt.get("ktb3") or [], rt.get("ktb10") or []
    if not bs or not k3:
        return
    ch = [(d1, v0, v1) for (d0, v0), (d1, v1) in zip(bs, bs[1:]) if v1 != v0][-5:]
    rows = [["일자", "기준금리", "국고 3년", "국고 10년"]]
    for d, v0, v1 in ch:
        rows.append([f"'{d[2:4]}.{int(d[5:7])}/{int(d[8:])}", f"{'인상' if v1 > v0 else '인하'} → {v1:.2f}%", f"{series_at(k3, d):.3f}%", f"{series_at(k10, d):.3f}%"])
    d = k3[-1][0]
    rows.append([f"'{d[2:4]}.{int(d[5:7])}/{int(d[8:])}", f"현재 {bs[-1][1]:.2f}%", f"{k3[-1][1]:.3f}%", f"{series_at(k10, d):.3f}%"])
    y0 = _chart_top(sl)
    bar(sl, PR, "기준금리 사이클과 국고채 금리", RX, y0, RW, h=276999, size=12)
    table(sl, rows, RX, y0 + 330000, RW, [0.9, 1.5, 1.0, 1.0], rowh=290000, size=9,
          align_cols={2: PP_ALIGN.CENTER, 3: PP_ALIGN.CENTER}, fills={(len(rows) - 1, c): PALEBLUE for c in range(4)})
    L, gap = kr_lines(rt)
    msg = (f"국고 3년-기준금리 {gap:+.0f}bp — " + ("추가 인상이 상당 부분 반영돼 수요예측 이후 금리 상승 여지 제한" if gap >= 40 else
                                               "동결 중심 반영 — 이벤트 결과에 따라 변동" if gap > -15 else "인하 기대 반영 — 발행 지연 시 금리 하락 가능")) if gap is not None else "##"
    textbox(sl, RX, y0 + 330000 + len(rows) * 290000 + 120000, RW, 700000,
            [{"text": "발행 시사점", "bold": True, "color": NAVY}, {"text": msg}], size=9.5, fill=PALEBLUE, margin=100000)


def panel_us(sl, PR, F, spec, X):
    rt = F["rates"]
    lo = (rt.get("ffr_lo") or [[None, None]])[-1][1]
    hi = (rt.get("ffr_hi") or [[None, None]])[-1][1]
    fomc = [d for d, nm in sorted(spec["events"].items()) if "FOMC" in nm and d >= F["today"].isoformat()][:5]
    if lo is None or not fomc:
        return
    fw = spec.get("fedwatch") or {}
    rngs = [(lo + 0.25 * k, hi + 0.25 * k) for k in range(-1, 4)]
    rows = [["FOMC"] + [f"{a:.2f}~{b:.2f}" for a, b in rngs]]
    for d in fomc:
        vals = (fw.get("probs") or {}).get(d) or {}
        rows.append([f"'{d[2:4]}.{int(d[5:7])}/{int(d[8:])}"] + [f"{vals[f'{a:.2f}']:.1f}" if f"{a:.2f}" in vals else "##" for a, _ in rngs])
    y0 = _chart_top(sl)
    bar(sl, PR, "CME FedWatch 정책금리 확률 (%)", RX, y0, RW, h=276999, size=12)
    table(sl, rows, RX, y0 + 330000, RW, [1.0] + [1] * len(rngs), rowh=300000, size=9, align_cols={i: PP_ALIGN.CENTER for i in range(1, len(rngs) + 1)})
    L, gap = us_lines(rt)
    msg = (f"美 2년물-정책금리 중간값 {gap:+.0f}bp — " + ("추가 인상 반영, 국내 장기물 금리에 상방 압력" if gap >= 25 else
                                                  "동결 중심 반영" if gap > -25 else "인하 기대 반영, 국내 금리 하방 요인")) if gap is not None else "##"
    asof = fw.get("asof") or "##"
    textbox(sl, RX, y0 + 330000 + len(rows) * 300000 + 120000, RW, 1050000, [
        {"text": "발행 시사점", "bold": True, "color": NAVY}, {"text": msg},
        {"text": f"현재 목표금리 {lo:.2f}~{hi:.2f}% / FedWatch 기준일 {asof} (설정 fedwatch 에 입력)", "size": 8, "color": SUB}],
        size=9.5, fill=PALEBLUE, margin=100000, space_after=3)


# ───────────────────────── 실행 ─────────────────────────
def auto_split(I, spec, today=None):
    """'미정' 체크: 만기만 받아 균형 배분 → I['split']·I['priority']·I['_alloc'] 채움."""
    from collections import defaultdict
    from tools import proposal_alloc as AL
    from tools import proposal_outlook as OL
    td = date.fromisoformat(today) if today else date.today()
    issue = date.fromisoformat(I["issue"])
    tenors = [int(x.split(":")[0]) for x in str(I["split"]).replace(" ", "").split(",") if x]
    base = int(I.get("base") or 0)
    C = P3.load_ctx(I["corp"], td)
    co, mine = C["co"], C["mine"]
    grade = M.gkey(I.get("grade")) or (M.gkey(mine[-1]["_grade"]) if mine and mine[-1].get("_grade") else "") or M.gkey(co.get("rating"))
    last = [r for r in mine if r.get("_dem")]
    lk = last[-1].get("original_rcept") if last else None
    prev = [r for r in mine if r.get("original_rcept") == lk] if lk else []
    lad = defaultdict(float)
    for x in (co.get("bonds") or {}).get("ladder") or []:
        lad[int(x["year"])] += (x.get("공모") or 0) / 1e8
    cv = M.Curves() if grade else M.Curves({})
    isins = [b.get("isin") for b in ((co.get("bonds") or {}).get("list") or []) if b.get("isin")]
    adj = M.issuer_spread(co.get("name") or "", grade, cv, isins) if cv else None
    try:
        ol = OL.load_latest(os.environ.get("RESEARCH_DIR"), td)
    except Exception:  # noqa: BLE001
        ol = None
    rules = (spec.get("rules") or {}).get("alloc") or {}
    res = AL.allocate(tenors, base, grade, C["pubs"], prev, lad, issue, td, cv, ol, adj, rules)
    I["split"] = AL.split_str(res["split"])
    I["priority"] = res["priority"]
    I["_alloc"] = res
    return res


def run(template, inputs, out, today=None, tighten=True):
    spec = load_spec()
    I = dict(inputs)
    if str(I.get("split_auto") or "").lower() in ("1", "true", "on", "yes"):
        auto_split(I, spec, today)
    tenors = [int(x.split(":")[0]) for x in I["split"].split(",")]
    base = int(I.get("base") or sum(int(x.split(":")[1]) for x in I["split"].split(",")))
    args = SimpleNamespace(template=template, corp=I["corp"], issue=I["issue"], base=base, max=int(I.get("max") or base),
                           tenors=",".join(map(str, tenors)), grade=I.get("grade") or None, band=(I.get("band") or None),
                           today=today, split=I["split"], out=None)
    with tempfile.TemporaryDirectory() as td:
        args.out = os.path.join(td, "v3.pptx")
        res = P3.build(args)
        F = res["facts"]
        if I.get("demand"):
            F["sch"]["수요예측"] = date.fromisoformat(I["demand"])
        transform(args.out, F, I, spec, out)
    if tighten and (spec.get("layout") or {}).get("tighten", True):
        try:
            from tools.pptx_tighten import tighten_file
            tighten_file(out, out, pad=(spec.get("layout") or {}).get("pad_emu", 90000))
        except Exception as e:  # noqa: BLE001
            print("tighten skipped:", type(e).__name__)
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--inputs", required=True, help="입력 JSON 파일 경로")
    ap.add_argument("--out", required=True)
    ap.add_argument("--today")
    ap.add_argument("--no-tighten", action="store_true")
    a = ap.parse_args()
    print(run(a.template, json.load(open(a.inputs, encoding="utf-8")), a.out, a.today, not a.no_tighten))
