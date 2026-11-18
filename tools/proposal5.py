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
    summ = None
    if on.get("exec_summary", True):
        summ = new_slide(prs, LAY_SUM)
        title(summ, PR, "권고안 요약")
        H(summ, f"{month_word} {'·'.join(str(t) for t in tenors)}년 {num(F['base'])}억 발행 권고 — 수요예측 {md(sch['수요예측'])}, 납입 {md(issue)}")
        bar(summ, PR, "권고안", X0, 1800000, CW)
        table(summ, [
            ["항목", "제안"],
            ["발행금액", f"{num(F['base'])}억원 (증액 시 최대 {num(F['max'])}억원)"],
            ["만기 구성", f"{tenor_txt} — 증액 시 {C.pri}년 우선"],
            ["금리밴드", f"{band}" + (f" ({band_note})" if band_note else "")],
            ["예상 발행금리", " / ".join(f"{t}년 ##%" for t in tenors)],
            ["일정", f"{md(sch['신고'])} 신고 → {md(sch['수요예측'])} 수요예측 → {md(issue)} 납입"],
            ["신용등급", f"{grade}"],
            ["인수·수수료", f"{num(F['base'])}억 {dflt.get('underwrite', '총액인수 확약')} / {fee}"],
        ], X0, 2110000, CW, [1.1, 3.6], rowh=300000, size=9.5)
        bar(summ, PR, "권고 근거", X2, 1800000, CW)
        reasons = []
        if F["prev"]:
            tot = sum(r["_dem"]["total"] for r in F["prev"] if r.get("_dem"))
            init = sum(r["_dem"]["total"] / r["_dem"]["ratio"] for r in F["prev"] if r.get("_dem") and r["_dem"].get("ratio"))
            sp = [r["_spread"] for r in F["prev"] if r.get("_spread") is not None]
            reasons.append(("수요", f"직전 회차 {num(init)}억 모집에 {num(tot)}억 주문({tot / init:.1f}:1)" + (f", 가산금리 {min(sp):+d}~{max(sp):+d}bp" if sp else "")))
        else:
            reasons.append(("수요", "AA급 최근 6개월 경쟁률 " + " / ".join(f"{C.mkt[t][2]} {C.mkt[t][0]:.1f}:1" for t in tenors if C.mkt[t][0])))
        dl, aa_now = rates_now(rt, "aa3")
        if dl:
            hi_d, hi_v = max(((d, v) for d, v in rt["aa3"] if d >= (date.fromisoformat(dl) - timedelta(days=92)).isoformat()), key=lambda x: x[1])
            reasons.append(("금리", f"AA- 3년 {aa_now:.3f}%({dl[5:].replace('-', '/')}) — 3개월 고점({hi_v:.3f}%) 대비 {(hi_v - aa_now) * 100:.0f}bp " + ("하락" if hi_v > aa_now else "")))
        ev_txt = "·".join(f"{nm.split(' (')[0]}({date.fromisoformat(d).month}/{date.fromisoformat(d).day})" for d, nm in C.evs if d > sch["수요예측"].isoformat())[:60]
        reasons.append(("이벤트", f"{ev_txt} 전 금리 확정" if ev_txt else f"{md(sch['정정'])} 정정신고로 금리 확정"))
        for i, (k, v) in enumerate(reasons[:3]):
            y = 2110000 + i * 680000
            rect(summ, X2, y, CW, 600000, fill=PALE)
            rect(summ, X2 + 90000, y + 120000, 360000, 360000, fill=NAVY, shape=2)
            textbox(summ, X2 + 90000, y + 120000, 360000, 360000, [{"text": str(i + 1), "align": PP_ALIGN.CENTER, "bold": True}],
                    size=12, color="FFFFFF", anchor=MSO_ANCHOR.MIDDLE, margin=0)
            textbox(summ, X2 + 560000, y + 40000, CW - 620000, 520000,
                    [{"text": k, "bold": True, "color": NAVY, "size": 10.5}, {"text": v}], size=9.5, anchor=MSO_ANCHOR.MIDDLE, margin=0)
        bar(summ, PR, "리스크 및 대응", X0, 4720000, CW)
        _, sp_now = None, F.get("sp_now")
        risks = [[f"AA- 스프레드 {sp_now:.0f}bp 수준" if sp_now else "스프레드 변동", "당일 가산금리 컷 기준 사전 합의"]]
        if F["prev"] and (issue - date.fromisoformat(F["prev"][0]["pay_dt"])).days < 120:
            risks.append(["직전 회차 투자자 한도 소진", "한도 잔여 점검 + 신규 기관 발굴"])
        else:
            risks.append(["동일 시기 차환 발행 경합", "수요예측 일정 선점"])
        risks.append(["한·미 정책 이벤트 전후 금리 변동", "이벤트 전 금리 확정, 결과 반영 후 재점검(##)"])
        table(summ, [["리스크", "대응"]] + risks, X0, 5030000, CW, [1.7, 2.6], rowh=400000, size=9.5)
        bar(summ, PR, "개별민평 및 국고채 금리 추이 (최근 6개월)", X2, 4140000, CW)
        a3, k3 = rt.get("aa3") or [], dict(map(tuple, rt.get("ktb3") or []))
        if a3:
            st = (date.fromisoformat(a3[-1][0]) - timedelta(days=183)).isoformat()
            pts = [(d, v) for d, v in a3 if d >= st and d in k3][::3]
            if pts and pts[-1][0] != a3[-1][0] and a3[-1][0] in k3:
                pts.append((a3[-1][0], a3[-1][1]))
            cd0 = CategoryChartData()
            cd0.categories = [d[5:].replace("-", "/") for d, _ in pts]
            cd0.add_series("AA- 3년 등급민평", [v for _, v in pts])
            cd0.add_series("국고 3년", [k3[d] for d, _ in pts])
            c0 = summ.shapes.add_chart(XL_CHART_TYPE.LINE, Emu(X2), Emu(4440000), Emu(CW), Emu(2050000), cd0).chart
            style_line(c0, [NAVY, SKY])
        textbox(summ, X2, 6500000, CW, 230000, ["주) ## 동사 개별민평 시계열 추가 (차트 → 데이터 편집), 자료: 한국은행 ECOS"], size=8, color="7F7F7F", margin=0)
        notes(summ, "[작성 가이드] 권고안(숫자) / 근거 3개 / 리스크·대응 / 개별민평·국고채 추이. ## 는 최종 확정 시 입력.")

    # ── 발행조건 ──
    build_cond(cond, F, I, C, spec, band, band_note, fee)

    # ── 예상 발행금리·프라이싱 ──
    rate = build_pricing(prs, LAY_I, PR, F, C) if on.get("pricing", True) else None
    # ── 만기구조 ──
    funds = build_maturity(prs, LAY_I, PR, F, C) if on.get("maturity", True) else None
    # ── 트랜치 ──
    build_tranche(tranche, PR, F, C)
    # ── 옵션 ──
    opt = build_options(prs, LAY_I, PR, F, C) if on.get("options", True) else None
    # ── 일정 ──
    build_schedule(sched, F, C)
    # ── 직전 회차 ──
    build_prev(prev_s, PR, F)
    # ── 크레딧 스토리 ──
    story = build_story(prs, LAY_II, PR, F, I) if on.get("story", True) else None
    # ── 시장: 대시보드·전망·우측 패널 ──
    dash = build_dashboard(prs, LAY_III, PR, F) if on.get("dashboard", True) else None
    outlook = build_outlook(prs, LAY_III, PR, F) if on.get("outlook", True) else None
    relayout(mpc, LAY_III)
    relayout(fed, LAY_III)
    R(fed, "Wiliams", "Williams")
    R(fed, "Washburm", "Washburn")
    R(mpc, "대응와", "대응과")
    panel_credit(credit, PR, F)
    panel_kr(kr, PR, F)
    panel_us(us, PR, F, spec)
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


def build_cond(cond, F, I, C, spec, band, band_note, fee):
    old = [sh for sh in cond.shapes if sh.has_table][0]
    tx, ty, tw = old.left, old.top, old.width
    old._element.getparent().remove(old._element)
    name, issue, sch = F["name"], F["issue"], F["sch"]
    tenors, split = C.tenors, F["split"]
    amt = []
    if F["due12"]:
        amt.append(B(f"**[만기도래 사채]** 향후 12개월 내 공모사채 만기 {num(F['due12'] / 1e8)}억원 — 차환 재원 선제 확보 필요"))
    if (I.get("issues_note") or "").strip():
        amt.append(B("**[자금 수요]** " + I["issues_note"].strip().splitlines()[0][:110]))
    elif F["capex_h"]:
        amt.append(B(f"**[자금 수요]** {F['period']} 설비투자 {num(F['capex_h'])}억원 vs 영업현금흐름 {num(F['ocf_h'])}억원 — 외부조달 수요 상존"))
    else:
        amt.append(B("**[자금 수요]** ##"))
    if F["prev"]:
        tot = sum(r["_dem"]["total"] for r in F["prev"] if r.get("_dem"))
        amt.append(B(f"**[시장 수급]** 직전 회차 주문 {num(tot)}억원 — {num(F['base'])}억원 신고·최대 {num(F['max'])}억원 증액은 소화 가능한 규모"))
    else:
        amt.append(B(f"**[시장 수급]** 최근 6개월 AA급 수요예측 참여 " + " / ".join(f"{C.mkt[t][2]} {C.mkt[t][1] / 1e4:.2f}조" for t in tenors if C.mkt[t][1])))
    mat = []
    pf = C.pri
    if F.get("sp_now"):
        mat.append(B(f"**[금리]** AA- 3년 스프레드 {F['sp_now']:.0f}bp — 증액 시 {pf}년물 우선 배정으로 금리·만기 균형"))
    dem_bits = [f"{t}년 {C.pdm[t]['_dem']['ratio']:.1f}:1" for t in tenors if t in C.pdm and C.pdm[t].get("_dem")]
    mk_bits = [f"{C.mkt[t][2]} {C.mkt[t][0]:.1f}:1" for t in tenors if C.mkt[t][0]]
    mat.append(B("**[수요]** " + (("직전 회차 " + " / ".join(dem_bits)) if dem_bits else "AA급 6개월 " + " / ".join(mk_bits)) + " — 만기별 물량 소화 충분"))
    mat.append(B("**[만기 분산]** " + " / ".join(f"'{(issue.year + t) % 100}년 기존 {num(C.lad.get(issue.year + t, 0))}억" for t in tenors) + " — 신규 만기 분산"))
    agencies = ""
    trig = (F.get("co") or {}).get("triggers") or {}
    if trig.get("agencies"):
        agencies = "·".join(a.get("agency", "") for a in trig["agencies"])
    ev_bits = [f"{nm}({date.fromisoformat(d).month}/{date.fromisoformat(d).day})" for d, nm in C.evs if d >= sch["신고"].isoformat()][:4]
    rows = [
        ["구분", "조건", "상세 내용"],
        ["발행금액", f"{num(F['base'])}억원\n(최대 {num(F['max'])}억원)", amt],
        ["인수", f"{num(F['base'])}억원\n{(spec.get('defaults') or {}).get('underwrite', '총액인수 확약')}", [B("수요예측 결과와 무관하게 신고금액 전액 인수 — 조달 확실성 확보")]],
        ["발행 만기", " / ".join(f"{t}년 {num(split[t])}억" for t in tenors), mat],
        ["금리밴드", band, [B(band_note or "희망 금리밴드")]],
        ["예상 발행금리", " / ".join(f"{t}년 ##%" for t in tenors), [B("개별민평 확인 후 입력 — 직전 회차 확정 가산금리 기준 참고 범위는 다음 장")]],
        ["신용등급", F["grade"], [B((f"{agencies} 기준 유효등급" if agencies else "평가사 ## 기준 유효등급") + " (본평가일 ##)")]],
        ["주요 일정", f"수요예측 {md(sch['수요예측'])}\n발행 {md(issue)}", [
            B(f"{md(sch['신고'])} 증권신고서 → {md(sch['수요예측'])} 수요예측 → {md(sch['정정'])} 정정신고서(금리 확정) → {md(issue)} 납입"),
            B(("·".join(ev_bits) + " 등 일정 고려") if ev_bits else "주요 이벤트 ##")]],
        ["인수수수료", fee, [B((spec.get("defaults") or {}).get("fee_note", "직전 회차와 동일"))]],
    ]
    table(cond, rows, tx, ty, tw, [1.0, 1.75, 7.0], rowh=[300000, 1020000, 340000, 1020000, 400000, 400000, 340000, 760000, 340000],
          size=9, align_cols={1: PP_ALIGN.CENTER}, name="ConditionTable")
    for sh in cond.shapes:
        if sh.has_text_frame and sh.text_frame.text.startswith("주) 상기 조건은"):
            sh.top = Emu(7060000)
    notes(cond, "[작성 가이드] 발행금액=만기도래 사채·자금 수요·시장 수급 / 만기=금리·수요·만기분산 / 금리밴드·수수료=직전 회차와 동일 / 신용등급=평가사별 유효등급. ## 는 최종 확정 시 입력.")


def build_pricing(prs, LAY, PR, F, C):
    s = new_slide(prs, LAY)
    title(s, PR, "예상 발행금리 및 프라이싱 전략")
    tenors = C.tenors
    sp = {t: C.pdm[t].get("_spread") for t in tenors if t in C.pdm}
    H(s, "직전 회차 확정 가산금리 기준 목표 설정 — 예상 확정금리 ##%대")
    bar(s, PR, "만기별 예상 확정금리 (개별민평 기준)", X0, 1800000, CW)
    rows = [["구분"] + [f"{t}년" for t in tenors],
            ["기준 민평"] + ["개별민평 ##%" for _ in tenors],
            ["밴드 하단"] + ["##%" for _ in tenors],
            ["직전 회차 수준"] + [(f"{sp[t]:+d}bp → ##%" if sp.get(t) is not None else "##%") for t in tenors],
            ["밴드 상단"] + ["##%" for _ in tenors]]
    dl, aa = rates_now(F["rates"], "aa3")
    if dl:
        rows.append([f"참고 ({dl[5:].replace('-', '/')})"] + [f"AA- 3년 {aa:.3f}%" if t == 3 else "-" for t in tenors])
    table(s, rows, X0, 2110000, CW, [1.4] + [1.1] * len(tenors), rowh=340000, size=9.5,
          align_cols={i: PP_ALIGN.CENTER for i in range(1, len(tenors) + 1)}, fills={(3, c): PALEBLUE for c in range(1, len(tenors) + 1)})
    # 동일·인접 등급 최근 사례
    g = F["grade"]
    since = (F["today"] - timedelta(days=150)).isoformat()
    cand = [r for r in F["pubs"] if r.get("_spread") is not None and (r.get("pay_dt") or "") >= since and same_band(r.get("_grade"), g)]
    cand.sort(key=lambda r: (r.get("pay_dt") or ""), reverse=True)
    mine = [r for r in F["prev"]]
    pick = mine + [r for r in cand if r not in mine][: max(0, 7 - len(mine))]
    bar(s, PR, "최근 우량물 공모채 프라이싱 사례", X2, 1800000, CW)
    prow = [["발행사 (발행일)", "만기", "경쟁률", "가산금리", "확정금리"]]
    for r in pick[:7]:
        d = r.get("_dem") or {}
        prow.append([f"{r['corp_name']}{r.get('series_raw', '')} ({(r.get('pay_dt') or '')[5:].replace('-', '/')})", tname(r["_tenor"]) if r.get("_tenor") else "-",
                     f"{d['ratio']:.1f}:1" if d.get("ratio") else "-", f"{r['_spread']:+d}bp", f"{r['_rate']:.3f}%" if r.get("_rate") else "##%"])
    table(s, prow, X2, 2110000, CW, [2.3, 0.6, 0.8, 0.85, 0.85], rowh=300000, size=8.5,
          align_cols={i: PP_ALIGN.CENTER for i in range(1, 5)}, fills={(i, c): PALEBLUE for i in range(1, len(mine) + 1) for c in range(5)})
    bar(s, PR, "프라이싱 시사점", X0, 4520000, W)
    csp = sorted(r["_spread"] for r in cand)
    med = csp[len(csp) // 2] if csp else None
    tgt = {t: (min(sp[t], -5) if sp.get(t) is not None else -5) for t in tenors}
    pc = [("동일 등급 시장 수준", f"최근 5개월 {g} 내외 {len(cand)}건 확정 가산금리 중앙값 {med:+d}bp" if med is not None else "동일 등급 최근 사례 ##"),
          ("동사 직전 프라이싱", (" / ".join(f"{t}년 {sp[t]:+d}bp" for t in tenors if sp.get(t) is not None) + " 확정") if sp else "직전 회차 사례 없음"),
          ("이번 목표 가산금리", " / ".join(f"{t}년 {tgt[t]:+d}bp 이내" for t in tenors) + " — 경쟁률 5:1 이상 시 추가 하향")]
    for i, (k, v) in enumerate(pc):
        x = X0 + i * (CW3 + 200000)
        textbox(s, x, 4840000, CW3, 1100000, [{"text": k, "bold": True, "color": NAVY, "size": 11}, {"text": v}],
                size=10, fill=PALEBLUE if i == 2 else PALE, margin=110000, anchor=MSO_ANCHOR.MIDDLE, space_after=4)
    source(s, "자료: 기업공시시스템, 한국은행 ECOS / 확정금리 미반영 건은 ##")
    notes(s, "[작성 가이드] 좌=만기별 예상 확정금리(개별민평 기준) / 우=동사 직전 회차 + 동일·인접 등급 최근 사례 / 하단=목표 가산금리.")
    return s


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
        {"text": f"'{years[mx] % 100}년 발행 후 만기 {num(after[mx])}억원으로 최대 — 연간 만기 상한(##억원) 내 관리", "lvl": 0},
        {"text": f"'{lo % 100}년 기존 만기 {num(C.lad.get(lo, 0))}억원으로 여유 — 증액분 배정 후보", "lvl": 0},
        {"text": f"향후 12개월 공모사채 만기 {num((F['due12'] or 0) / 1e8)}억원 차환 계획과 연계", "lvl": 0}],
        size=10, fill=PALE, margin=120000, space_after=4)
    source(s, "주) 공모사채 잔액 기준(채권기본정보·공시), 이번 발행은 기본 금액 기준 / 자료: 기업공시시스템, 금융위원회", w=9000000)
    notes(s, "[작성 가이드] 동사 공모사채 만기 사다리 + 이번 발행 반영 후 변화.")
    return s


def build_tranche(sl, PR, F, C):
    for sh in list(sl.shapes):
        if sh.has_table or (sh.has_text_frame and sh.name.startswith("TextBox") and sh.top > 4500000 and sh.left > 5000000):
            sh._element.getparent().remove(sh._element)
    for sh in sl.shapes:
        if sh.has_text_frame and sh.text_frame.text.startswith("등급·만기별 수요와"):
            set_shape_text(sh._element, "만기별 제안 근거 요약")
        if sh.name == "모서리가 둥근 직사각형 8":
            for r_ in sh._element.iter(qn("a:rPr")):
                r_.set("sz", "1200")
    for sh in sl.shapes:
        if sh.name == "TextBox 10" or (sh.has_text_frame and sh.top < 1700000 and sh.top > 1200000 and sh.width > 8000000):
            set_shape_text(sh._element, f"{C.pri}년물 중심 증액 권고 — 직전 회차 수요와 연간 만기 분산을 함께 반영")
            break
    issue = F["issue"]
    rows = [["구분", "발행 금액", "제안 근거"]]
    for t in C.tenors:
        y = issue.year + t
        r = C.pdm.get(t)
        bits = []
        if r and r.get("_dem"):
            bits.append(f"직전 회차 {t}년물에 {num(r['_dem']['total'])}억원({r['_dem']['ratio']:.1f}:1)이 몰렸고" + (f" 가산금리 {r['_spread']:+d}bp로 확정되었습니다." if r.get("_spread") is not None else "."))
        elif C.mkt[t][0]:
            bits.append(f"최근 6개월 AA급 {C.mkt[t][2]}물 평균 경쟁률은 {C.mkt[t][0]:.1f}:1입니다.")
        bits.append(f"만기('{y % 100}년)에는 기존 {num(C.lad.get(y, 0))}억원이 있어 {num(F['split'][t])}억원을 더하면 {num(C.lad.get(y, 0) + F['split'][t])}억원입니다.")
        if t == C.pri:
            bits.append("금리 인상 국면에서 고정금리를 확보하는 핵심 트랜치로, 증액 시 우선 배정을 권고합니다.")
        elif t == C.tenors[0]:
            bits.append("단기물 수요를 흡수해 낮은 가산금리를 먼저 확정하는 역할을 합니다.")
        rows.append([f"{t}년", f"{num(F['split'][t])}억원", " ".join(bits)])
    rh = max(450000, min(900000, 2300000 // len(C.tenors)))
    table(sl, rows, 451083, 1838003, 9790601, [0.7, 1.0, 8.1], rowh=[280000] + [rh] * len(C.tenors), size=9.5,
          align_cols={1: PP_ALIGN.CENTER}, fills={(1 + C.tenors.index(C.pri), c): PALEBLUE for c in (1, 2)})
    rr = [["항목"] + [f"{t}년 ({num(F['split'][t])}억)" for t in C.tenors],
          ["직전 회차 수요"] + [(f"{C.pdm[t]['_dem']['ratio']:.1f}:1" if t in C.pdm and C.pdm[t].get("_dem") else "-") for t in C.tenors],
          ["직전 가산금리"] + [(f"{C.pdm[t]['_spread']:+d}bp" if t in C.pdm and C.pdm[t].get("_spread") is not None else "-") for t in C.tenors],
          ["만기 연도"] + [f"'{(issue.year + t) % 100} (기존 {num(C.lad.get(issue.year + t, 0))})" for t in C.tenors],
          ["시장 수요(AA급 6M)"] + [(f"{C.mkt[t][1] / 1e4:.2f}조" if C.mkt[t][1] else "-") for t in C.tenors]]
    table(sl, rr, 5489407, 4716722, 4752279, [1.3] + [1.2] * len(C.tenors), rowh=[280000] + [400000] * 4, size=8.5,
          align_cols={i: PP_ALIGN.CENTER for i in range(1, len(C.tenors) + 1)})
    notes(sl, "[작성 가이드] 상단=만기별 '그래서 무엇을 우선하라는가' / 하단 우측=근거 요약 / 하단 좌측=만기별 시장 수요 차트.")


def build_options(prs, LAY, PR, F, C):
    s = new_slide(prs, LAY)
    title(s, PR, "만기 전략 옵션 비교")
    t0, tp = C.tenors[0], C.pri
    y = F["issue"].year

    def mat_line(sp):
        return " / ".join(f"'{(y + t) % 100} {num(C.lad.get(y + t, 0) + sp[t])}" for t in C.tenors)
    a_lab = f"옵션 A  ·  {tp}년 우선 증액 (권고)"
    b_lab = f"옵션 B  ·  {t0}년 우선 증액"
    H(s, f"옵션 A({tp}년 우선 증액) 권고 — 금리 고정과 만기 분산, 직전 회차 수요를 함께 반영")
    da, db = C.pdm.get(tp), C.pdm.get(t0)
    cards = [
        (a_lab, NAVY, PALEBLUE, [["최대 구성", " / ".join(f"{t}년 {num(C.a_max[t])}" for t in C.tenors) + "억"], ["만기 (억원)", mat_line(C.a_max)]],
         [f"{tp}년 고정금리 확보 — 인상 국면에서 차환 금리 노출 축소",
          (f"직전 회차 {tp}년물 {da['_dem']['ratio']:.1f}:1 수요 확인" if da and da.get("_dem") else f"{C.mkt[tp][2]} 시장 수요 활용"),
          "신규 만기를 뒤로 분산"],
         ["단기물 대비 기간 프리미엄(##bp)만큼 평균 금리 상승", f"'{(y + tp) % 100}년 만기 {num(C.lad.get(y + tp, 0) + C.a_max[tp])}억원으로 증가"],
         "만기 장기화와 금리 고정이 조달비용보다 우선일 때"),
        (b_lab, DGREY, PALE, [["최대 구성", " / ".join(f"{t}년 {num(C.b_max[t])}" for t in C.tenors) + "억"], ["만기 (억원)", mat_line(C.b_max)]],
         ["단기물 비중이 높아 평균 조달금리 최저",
          (f"직전 회차 {t0}년물 {db['_dem']['ratio']:.1f}:1 수요" if db and db.get("_dem") else "증권사·운용사 단기물 수요 흡수"),
          "기간 프리미엄 절감"],
         [f"'{(y + t0) % 100}년 만기 {num(C.lad.get(y + t0, 0) + C.b_max[t0])}억원으로 증가", "차환 시점 금리 재노출"],
         "금리 고점을 인식하고 단기 조달비용 최소화가 우선일 때"),
    ]
    for i, (hd, col, bg, kv, pros, cons, fit) in enumerate(cards):
        x = X0 if i == 0 else X2
        rect(s, x, 1800000, CW, 420000, fill=col)
        textbox(s, x, 1800000, CW, 420000, [{"text": hd, "align": PP_ALIGN.CENTER, "bold": True}], size=12, color="FFFFFF",
                anchor=MSO_ANCHOR.MIDDLE, margin=0)
        table(s, kv, x, 2290000, CW, [1.0, 2.6], rowh=330000, size=9.5, header_rows=0, align_cols={1: PP_ALIGN.CENTER},
              bold_cells={(0, 1), (1, 1)})
        textbox(s, x, 3060000, CW, 1350000, [{"text": "장점", "bold": True, "color": NAVY, "size": 10.5}] + [B(t) for t in pros],
                size=10.5, fill=bg, margin=110000, space_after=4)
        textbox(s, x, 4490000, CW, 1150000, [{"text": "단점", "bold": True, "color": RED, "size": 10.5}] + [B(t) for t in cons],
                size=10.5, fill="FBEAEA", margin=110000, space_after=4)
        textbox(s, x, 5720000, CW, 500000, [{"text": "**적합한 상황**  " + fit}], size=10, margin=60000, anchor=MSO_ANCHOR.MIDDLE)
    textbox(s, X0, 6330000, W, 420000, [{"text": f"**신한 의견: 옵션 A** — {tp}년물 경쟁률 3:1 미만 시 옵션 B로 전환, 연간 만기 상한 ##억원 이내", "align": PP_ALIGN.CENTER}],
            size=10.5, color=NAVY, fill=PALEBLUE, margin=60000, anchor=MSO_ANCHOR.MIDDLE)
    notes(s, "[작성 가이드] 옵션별 카드: 구성·만기 영향 → 장점 → 단점 → 적합한 상황 → 하단 의견·전환 조건.")
    return s


def build_schedule(sl, F, C):
    sch, issue = F["sch"], F["issue"]
    R(sl, re.sub(r".*", "", ""), "")
    for sh in sl.shapes:
        if sh.has_text_frame and "발행일정 제안" in sh.text_frame.text:
            set_shape_text(sh._element, f"한·미 주요 이벤트를 고려하여 {md(issue)} 발행일정 제안")
    tb = next((sh for sh in sl.shapes if sh.name == "TextBox 5"), None)
    before = [f"{nm}({date.fromisoformat(d).month}/{date.fromisoformat(d).day})" for d, nm in C.evs if sch["정정"].isoformat() < d]
    after_due = (F.get("due12") or 0) / 1e8
    lines = [{"text": f"**{md(sch['수요예측'])} 수요예측 — 이벤트 전 금리 확정**", "lvl": 0}]
    if before:
        lines.append({"text": "·".join(before[:4]) + f" 이전, {md(sch['정정'])} 정정신고서로 발행금리 확정", "lvl": 1})
    if F["prev"]:
        p0 = F["prev"][0]
        tot = sum(r["_dem"]["total"] for r in F["prev"] if r.get("_dem"))
        lines.append({"text": f"직전 회차({(p0.get('_demand_dt') or p0.get('pay_dt') or '')[:10]}) 주문 {num(tot)}억원 — 같은 투자자군 재공략", "lvl": 1})
    lines.append({"text": "**시장 여건 점검**", "lvl": 0})
    lines.append({"text": "결과 반영 후 일정·밴드 최종 점검 (##)", "lvl": 1})
    lines.append({"text": f"**→ {md(sch['신고'])} 신고 · {md(sch['수요예측'])} 수요예측 · {md(issue)} 납입" + (f" — 향후 12개월 만기 {num(after_due)}억원 차환 재원" if after_due else "") + "**", "lvl": 0})
    if tb is not None:
        fill_tf(tb.text_frame, lines, size=9.5, space_after=3)
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
    notes(sl, "[작성 가이드] 달력형 일정표. 신고서·수요예측·금리 확정(정정신고)·납입을 색으로 구분, 한·미 이벤트는 docs/proposal_spec.md 의 events.")


def build_prev(sl, PR, F):
    prev = F["prev"]
    if not prev or not all(r.get("_dem") for r in prev):
        return
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


def build_story(prs, LAY, PR, F, I):
    s = new_slide(prs, LAY)
    title(s, PR, "크레딧 스토리 및 투자자 예상 질문")
    H(s, "수요예측 메시지: ## (회사 핵심 강점 한 줄)")
    dr = draft_story(F, I)
    co = F.get("co") or {}
    if dr:
        pts = [tuple(x[:2]) for x in dr["points"][:4]]
        qa = [tuple(x[:2]) for x in dr["qa"][:5]]
    else:
        notes_ = [x.strip() for x in ((I.get("credit_note") or "") + "\n" + (I.get("issues_note") or "")).splitlines() if x.strip()]
        pts = [("사업 경쟁력", notes_[0] if notes_ else "##"), ("그룹 내 위상·재무 지원", notes_[1] if len(notes_) > 1 else "##"),
               ("최근 이슈", notes_[2] if len(notes_) > 2 else "##"),
               ("시장 수용성", f"신용등급 {F['grade']}" + (f", 직전 회차 {sum(r['_dem']['total'] for r in F['prev'] if r.get('_dem')):,.0f}억원 주문" if F["prev"] else ""))]
        qa = [("실적 변동 요인은?", "##"), ("차입 부담은?", "##"), ("투자 계획은?", "##"), ("등급 전망은?", "##"), ("추가 발행 계획은?", "##")]
    bar(s, PR, "투자 포인트 (크레딧 스토리)", X0, 1800000, CW)
    for i, (k, v) in enumerate(pts):
        y = 2110000 + i * 900000
        rect(s, X0, y, CW, 800000, fill=PALE)
        rect(s, X0 + 100000, y + 200000, 400000, 400000, fill=NAVY, shape=2)
        textbox(s, X0 + 100000, y + 200000, 400000, 400000, [{"text": str(i + 1), "align": PP_ALIGN.CENTER, "bold": True}],
                size=13, color="FFFFFF", anchor=MSO_ANCHOR.MIDDLE, margin=0)
        textbox(s, X0 + 620000, y + 60000, CW - 700000, 680000, [{"text": k, "bold": True, "color": NAVY, "size": 10.5}, {"text": v}],
                size=9.5, anchor=MSO_ANCHOR.MIDDLE, margin=0)
    bar(s, PR, "투자자 예상 질문 및 답변 방향", X2, 1800000, CW)
    table(s, [["예상 질문", "답변 방향"]] + [list(x) for x in qa], X2, 2110000, CW, [1.45, 2.75], rowh=[300000] + [600000] * 5, size=9)
    textbox(s, X2, 5500000, CW, 800000, [{"text": "투자자 배포 자료 원칙", "bold": True, "color": RED},
                                          {"text": "신용지표 원수치 대신 일시적 요인과 대응 논리 중심으로 기재"}],
            size=9.5, fill="FBEAEA", margin=110000, anchor=MSO_ANCHOR.MIDDLE)
    source(s, "자료: 공시, 신용평가사, 팝업 메모" + (", Claude 초안(검토 필요)" if dr else ""))
    notes(s, "[작성 가이드] 크레딧 스토리 4개 + 예상 Q&A 5개는 항상 채워서 제출." + (" 이 장은 Claude가 초안을 작성했으니 사실관계를 확인할 것." if dr else ""))
    return s


def build_dashboard(prs, LAY, PR, F):
    s = new_slide(prs, LAY)
    title(s, PR, "시장 한눈에 보기 (Market Dashboard)")
    rt = F["rates"]
    H(s, "금리 레벨·스프레드·이벤트 점검 — 발행 영향 요약")

    def row(lab, key, unit="%", scale=1, dec=3):
        s_ = rt.get(key) or []
        if not s_:
            return [lab, "##", "##", "##", "##"]
        d_now, v = s_[-1]
        dn = date.fromisoformat(d_now)
        m1 = series_at(s_, (dn - timedelta(days=31)).isoformat())
        y1 = series_at(s_, (dn - timedelta(days=365)).isoformat())
        f = lambda x: f"{x * scale:.{dec}f}{unit}" if x is not None else "-"  # noqa: E731
        ch = f"{(v - m1) * 100:+.0f}bp" if m1 is not None and unit == "%" else (f"{(v - m1) * scale:+.1f}bp" if m1 is not None else "-")
        return [lab, f(v), f(m1), f(y1), ch]
    sp = []
    a3, k3 = rt.get("aa3") or [], dict(map(tuple, rt.get("ktb3") or []))
    sp_series = [[d, (v - k3[d])] for d, v in a3 if d in k3]
    rows = [["지표", "현재", "1개월 전", "1년 전", "1개월 변화"],
            row("한은 기준금리", "base", dec=2), row("국고 3년", "ktb3"), row("국고 10년", "ktb10"), row("AA- 3년", "aa3")]
    rt["_sp"] = sp_series
    rows.append(row("AA- 스프레드", "_sp", unit="bp", scale=100, dec=1))
    rows += [row("CP 91일", "cp91", dec=2), row("美 FF 상단", "ffr_hi", dec=2), row("美 국채 2년", "ust2", dec=2), row("美 국채 10년", "ust10", dec=2)]
    rt.pop("_sp", None)
    del sp
    table(s, rows, X0, 1800000, W, [2.4, 1.5, 1.5, 1.5, 1.5], rowh=360000, size=10,
          align_cols={i: PP_ALIGN.CENTER for i in range(1, 5)}, bold_cells={(r, 1) for r in range(1, len(rows))})
    textbox(s, X0, 5700000, W, 700000, [{"text": "발행 영향: ## (금리 레벨·스프레드·이벤트 종합 판단 한 줄)"}], size=10, fill=PALEBLUE,
            margin=110000, anchor=MSO_ANCHOR.MIDDLE)
    source(s, "자료: 한국은행 ECOS, FRED / 기준일은 각 지표 최신 수집일", w=9000000)
    notes(s, "[작성 가이드] 모든 수치는 data/ref/mkt/rates.json 자동. FedWatch 확률은 미국 장표에서 수기 입력.")
    return s


def build_outlook(prs, LAY, PR, F):
    s = new_slide(prs, LAY)
    title(s, PR, "향후 전망 시나리오 (3·6개월)")
    H(s, "향후 3·6개월 금리 전망: ##")
    rt = F["rates"]
    cur = lambda k, d=2: (f"{rt[k][-1][1]:.{d}f}%" if rt.get(k) else "##")  # noqa: E731
    table(s, [["지표", "현재", "3개월 후 기본", "3개월 후 낙관", "3개월 후 비관", "6개월 후 기본"],
              ["한은 기준금리", cur("base"), "##", "##", "##", "##"], ["국고 3년", cur("ktb3"), "##", "##", "##", "##"],
              ["AA- 3년", cur("aa3"), "##", "##", "##", "##"], ["AA- 스프레드", f"{F['sp_now']:.0f}bp" if F.get("sp_now") else "##", "##", "##", "##", "##"],
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


def _chart_top(slide):
    bars = [sh for sh in slide.shapes if sh.name == "모서리가 둥근 직사각형 8"]
    return min(b.top for b in bars) if bars else 3440000


def panel_credit(sl, PR, F):
    rt = F["rates"]
    a3, k3 = rt.get("aa3") or [], dict(map(tuple, rt.get("ktb3") or []))
    if not a3:
        return
    end = a3[-1][0]
    st = (date.fromisoformat(end) - timedelta(days=365)).isoformat()
    ev = []
    bs = rt.get("base") or []
    for (d0, v0), (d1, v1) in zip(bs, bs[1:]):
        if v1 != v0 and d1 >= st:
            ev.append((d1, f"한은 {'인상' if v1 > v0 else '인하'} {v0:.2f}→{v1:.2f}%"))
    hi = max(((d, v) for d, v in a3 if d >= st), key=lambda x: x[1])
    ev.append((hi[0], "AA- 3년 1년 내 고점"))
    ev.append((end, "현재"))
    ev = sorted(set(ev))[-6:]
    rows = [["일자", "이벤트", "AA- 3년", "스프레드"]]
    for d, nm in ev:
        v = series_at(a3, d)
        kv = series_at(sorted(k3.items()), d)
        rows.append([f"'{d[2:4]}.{int(d[5:7])}/{int(d[8:])}", nm, f"{v:.3f}%", f"{(v - kv) * 100:.1f}bp" if kv else "-"])
    y0 = _chart_top(sl)
    bar(sl, PR, "주요 이벤트와 AA- 3년 금리", RX, y0, RW, h=276999, size=12)
    table(sl, rows, RX, y0 + 330000, RW, [0.9, 2.1, 0.9, 0.9], rowh=300000, size=9,
          align_cols={2: PP_ALIGN.CENTER, 3: PP_ALIGN.CENTER}, fills={(len(rows) - 1, c): PALEBLUE for c in range(4)})
    textbox(sl, RX, y0 + 330000 + len(rows) * 300000 + 150000, RW, 800000,
            [{"text": "발행 시사점", "bold": True, "color": NAVY}, {"text": "## (스프레드·금리 레벨 판단 한 줄)"}], size=9.5, fill=PALEBLUE, margin=100000)


def panel_kr(sl, PR, F):
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
    textbox(sl, RX, y0 + 330000 + len(rows) * 290000 + 120000, RW, 700000,
            [{"text": "발행 시사점", "bold": True, "color": NAVY}, {"text": f"국고 3년-기준금리 {(k3[-1][1] - bs[-1][1]) * 100:.0f}bp — ## (선반영 판단)"}],
            size=9.5, fill=PALEBLUE, margin=100000)


def panel_us(sl, PR, F, spec):
    rt = F["rates"]
    lo = (rt.get("ffr_lo") or [[None, None]])[-1][1]
    hi = (rt.get("ffr_hi") or [[None, None]])[-1][1]
    fomc = [d for d, nm in sorted(spec["events"].items()) if "FOMC" in nm and d >= F["today"].isoformat()][:5]
    if lo is None or not fomc:
        return
    rngs = [(lo + 0.25 * k, hi + 0.25 * k) for k in range(0, 5)]
    rows = [["FOMC"] + [f"{a:.2f}~{b:.2f}" for a, b in rngs]]
    for d in fomc:
        rows.append([f"'{d[2:4]}.{int(d[5:7])}/{int(d[8:])}"] + ["##"] * len(rngs))
    y0 = _chart_top(sl)
    bar(sl, PR, "CME FedWatch 정책금리 확률 (%)", RX, y0, RW, h=276999, size=12)
    table(sl, rows, RX, y0 + 330000, RW, [1.0] + [1] * len(rngs), rowh=300000, size=9, align_cols={i: PP_ALIGN.CENTER for i in range(1, len(rngs) + 1)})
    textbox(sl, RX, y0 + 330000 + len(rows) * 300000 + 120000, RW, 1050000, [
        {"text": "발행 시사점", "bold": True, "color": NAVY}, {"text": "## (회의별 최빈 구간과 발행 일정의 관계)"},
        {"text": f"현재 목표금리 {lo:.2f}~{hi:.2f}% / 기준일 ## (CME FedWatch → Probabilities)", "size": 8, "color": SUB}],
        size=9.5, fill=PALEBLUE, margin=100000, space_after=3)


# ───────────────────────── 실행 ─────────────────────────
def run(template, inputs, out, today=None, tighten=True):
    spec = load_spec()
    I = dict(inputs)
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
