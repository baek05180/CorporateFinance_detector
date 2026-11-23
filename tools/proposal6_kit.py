"""제안서 v6 장표 도구 — 리서치형 DCM 자료 형식(핵심 메시지 · 요점 · 표/차트 · 시사점 · Note/Source).

페이지 틀: 섹션 이름 / 제목 / 구분선 / 따옴표 핵심 메시지 / 요점 / 본문 패널 / 시사점 상자 / Note·Source / 쪽번호.
좌표 단위는 EMU (A4 가로 템플릿 10,693,400 × 7,562,850).
"""
import math
import re

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from tools.pptx_kit import (  # noqa: F401
    BODY_FONT, LGREY, MSO_ANCHOR, NAVY, PALE, PALEBLUE, PP_ALIGN, RED, SKY, SUB, TXT, RGBColor, _bullet, add_runs, etree,
    rect, table, textbox)

SW, SH = 10693400, 7562850
L, R = 451084, 10242400          # 좌우 여백
CW = R - L                        # 본문 폭
Y_HEAD = 900000                   # 핵심 메시지
Y_BOT = 6700000                   # 본문 하단
GREY = "7F7F7F"
DEEP = "0B2D4F"                   # 패널 제목 띠
ACC = "F2A900"                    # 강조(노랑)
C_SERIES = [NAVY, SKY, "9DB7C9", LGREY, "1F7A8C", ACC]
GAP = 160000


def lines_needed(text, w, size):
    """글자 수로 줄 수 추정 (한글 1em, 영문·숫자 0.55em)."""
    t = re.sub(r"\*\*", "", text or "")
    em = sum(1.0 if ord(ch) > 0x2E80 else 0.55 for ch in t)
    per = max(1.0, (w / 12700 - 10) / (size * 1.0))
    return max(1, math.ceil(em / per))


def text_h(paras, w, size, gap_pt=3):
    h = 0
    for p in paras:
        s = p.get("size", size) if isinstance(p, dict) else size
        t = p["text"] if isinstance(p, dict) else p
        h += lines_needed(t, w - (140000 if isinstance(p, dict) and p.get("bul") else 0), s) * s * 1.32 * 12700 + gap_pt * 12700
    return int(h) + 60000


def para(tf, text, size=10, color=TXT, bold=False, bul=None, align=None, after=3, first=False, indent=150000):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    if bul:
        _bullet(p, bul, indent, indent)
    if align is not None:
        p.alignment = align
    p.space_after = Pt(after)
    add_runs(p, text, size, bold, color)
    return p


def box(slide, x, y, w, h, paras, size=10, color=TXT, fill=None, margin=50000, anchor=MSO_ANCHOR.TOP, line=None, name=None):
    """paras: str 또는 dict(text, size, bold, color, bul, align, after)."""
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = Emu(margin)
    tf.margin_top = tf.margin_bottom = Emu(margin // 2)
    tf.vertical_anchor = anchor
    if fill:
        tb.fill.solid()
        tb.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line:
        tb.line.color.rgb = RGBColor.from_string(line)
        tb.line.width = Emu(6350)
    for i, p in enumerate(paras):
        p = {"text": p} if isinstance(p, str) else p
        para(tf, p["text"], p.get("size", size), p.get("color", color), p.get("bold", False), p.get("bul"),
             p.get("align"), p.get("after", 3), first=(i == 0))
    return tb


def hline(slide, x, y, w, color=NAVY, pt=1.5):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(x), Emu(y), Emu(x + w), Emu(y))
    ln.line.color.rgb = RGBColor.from_string(color)
    ln.line.width = Pt(pt)
    return ln


class Page:
    """한 장. 위에서부터 차곡차곡 쌓는다 (self.y = 다음 빈 자리)."""

    def __init__(self, slide, section, title, headline=None, bullets=None, key=None):
        self.s, self.key = slide, key
        self.ybot = Y_BOT
        box(slide, L, 200000, 6000000, 230000, [{"text": section, "size": 9, "color": GREY}], margin=0)
        box(slide, R - 2400000, 200000, 2400000, 230000, [{"text": "Strictly Confidential", "size": 9, "color": GREY,
                                                            "align": PP_ALIGN.RIGHT}], margin=0)
        box(slide, L, 400000, CW, 400000, [{"text": title, "size": 19, "bold": True, "color": NAVY}], margin=0,
            name="Title")
        hline(slide, L, 830000, CW, NAVY, 1.75)
        self.y = Y_HEAD
        if headline:
            hh = text_h([{"text": headline, "size": 12.5}], CW - 300000, 12.5, 0) - 20000
            box(slide, L + 150000, self.y, CW - 300000, hh, [{"text": f"“{headline}”", "size": 12.5, "bold": True,
                                                              "color": DEEP}], margin=0, name="Headline")
            self.y += hh + 60000
        if bullets:
            self.bullets(bullets)

    def bullets(self, items, x=None, w=None, size=10, y=None):
        """items: (굵은 머리, 본문) 또는 문자열."""
        x = L if x is None else x
        w = CW if w is None else w
        y0 = self.y if y is None else y
        ps = []
        for it in items:
            if not it:
                continue
            t = f"**{it[0]}:** {it[1]}" if isinstance(it, (tuple, list)) else it
            ps.append({"text": t, "size": size, "bul": "·", "after": 3})
        h = text_h(ps, w, size)
        box(self.s, x, y0, w, h, ps, margin=0)
        if y is None:
            self.y = y0 + h + 40000
        return y0 + h

    def panel(self, x, y, w, title, unit=None):
        """패널 제목 띠 — 아래 내용 시작 y 를 돌려준다."""
        rect(self.s, x, y, w, 250000, fill=DEEP, shape=2)
        box(self.s, x + 60000, y, w - 120000, 250000, [{"text": title, "size": 9.5, "bold": True, "color": "FFFFFF"}],
            margin=30000, anchor=MSO_ANCHOR.MIDDLE)
        if unit:
            box(self.s, x + w - 1800000, y, 1740000, 250000, [{"text": unit, "size": 8, "color": "D9E2EA",
                                                               "align": PP_ALIGN.RIGHT}], margin=30000,
                anchor=MSO_ANCHOR.MIDDLE)
        return y + 300000

    def insight(self, title, items, y=None, x=None, w=None, size=9, max_bottom=None):
        x = L if x is None else x
        w = CW if w is None else w
        max_bottom = max_bottom or self.ybot
        ps = [{"text": title, "size": 10, "bold": True, "color": NAVY, "after": 2}]
        for it in items:
            if not it:
                continue
            t = f"**{it[0]}:** {it[1]}" if isinstance(it, (tuple, list)) else it
            ps.append({"text": t, "size": size, "bul": "·", "after": 2})
        h = text_h(ps, w - 160000, size, 2) + 40000
        y0 = (max_bottom - h) if y is None else y
        rect(self.s, x, y0, w, h, fill="EEF3F7")
        rect(self.s, x, y0, 45000, h, fill=NAVY)
        box(self.s, x + 110000, y0 + 30000, w - 160000, h - 40000, ps, margin=0)
        return y0, y0 + h

    def kpis(self, y, items, h=720000):
        """큰 숫자 칸: (숫자, 설명)."""
        n = len(items)
        gw = (CW - (n - 1) * 120000) // n
        for i, (big, lab) in enumerate(items):
            x = L + i * (gw + 120000)
            rect(self.s, x, y, gw, h, fill="F4F6F8")
            rect(self.s, x, y, gw, 35000, fill=NAVY)
            box(self.s, x, y + 70000, gw, 330000, [{"text": big, "size": 17, "bold": True, "color": NAVY,
                                                    "align": PP_ALIGN.CENTER}], margin=20000)
            box(self.s, x, y + 400000, gw, h - 420000, [{"text": lab, "size": 8, "color": SUB,
                                                         "align": PP_ALIGN.CENTER}], margin=30000)
        return y + h

    def foot(self, note=None, source=None):
        ps = []
        if note:
            ps.append({"text": f"Note: {note}", "size": 7, "color": GREY, "after": 0})
        if source:
            ps.append({"text": f"Source: {source}", "size": 7, "color": GREY, "after": 0})
        if ps:
            h = text_h(ps, 8100000, 7, 0)
            box(self.s, L, max(6760000, 7230000 - h), 8100000, h, ps, margin=0, anchor=MSO_ANCHOR.BOTTOM, name="Foot")


def tbl(slide, rows, x, y, w, colw, rowh=250000, size=8.5, **kw):
    """표 — 숫자 열 가운데 정렬 기본. 높이를 돌려준다."""
    ncol = len(rows[0])
    al = kw.pop("align_cols", None) or {c: PP_ALIGN.CENTER for c in range(1, ncol)}
    rh = rowh if isinstance(rowh, (list, tuple)) else [rowh] * len(rows)
    table(slide, rows, x, y, w, colw, rowh=list(rh), size=size, align_cols=al, **kw)
    return y + sum(rh)


def chart(slide, kind, x, y, w, h, cats, series, fmt="#,##0", labels=True, legend=None, colors=None, gap=70,
          overlap=None, ymin=None, ymax=None, label_size=8, axis=False, font=8.5, label_pos=None):
    """kind: col, colstack, bar, line. series: [(이름, 값들)]."""
    cd = CategoryChartData()
    cd.categories = cats
    for nm, vals in series:
        cd.add_series(nm, vals)
    ct = {"col": XL_CHART_TYPE.COLUMN_CLUSTERED, "colstack": XL_CHART_TYPE.COLUMN_STACKED,
          "bar": XL_CHART_TYPE.BAR_CLUSTERED, "line": XL_CHART_TYPE.LINE}[kind]
    gf = slide.shapes.add_chart(ct, Emu(x), Emu(y), Emu(w), Emu(h), cd)
    c = gf.chart
    c.has_title = False
    c.font.size = Pt(font)
    c.font.name = BODY_FONT
    leg = legend if legend is not None else len(series) > 1
    c.has_legend = leg
    if leg:
        c.legend.position = XL_LEGEND_POSITION.TOP
        c.legend.include_in_layout = False
        c.legend.font.size = Pt(8)
    colors = colors or C_SERIES
    plot = c.plots[0]
    if kind != "line":
        plot.gap_width = gap
        if overlap is not None:
            plot.overlap = overlap
        elif kind == "colstack":
            plot.overlap = 100
    for i, s in enumerate(plot.series):
        col = RGBColor.from_string(colors[i % len(colors)])
        if kind == "line":
            s.format.line.color.rgb = col
            s.format.line.width = Pt(1.75)
            s.smooth = False
            from pptx.enum.chart import XL_MARKER_STYLE
            s.marker.style = XL_MARKER_STYLE.NONE
        else:
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = col
    if labels and kind != "line":
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(label_size)
        dl.number_format = fmt
        dl.number_format_is_linked = False
        if label_pos is not None:
            dl.position = label_pos
        elif kind in ("col", "bar"):
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
        else:
            dl.position = XL_LABEL_POSITION.CENTER
    va = c.value_axis
    va.has_major_gridlines = kind == "line"
    if kind == "line":
        va.major_gridlines.format.line.color.rgb = RGBColor.from_string("E3E3E3")
    va.visible = axis or kind == "line"
    va.tick_labels.font.size = Pt(8)
    va.tick_labels.number_format = fmt
    va.tick_labels.number_format_is_linked = False
    va.format.line.fill.background()
    if ymin is not None:
        va.minimum_scale = ymin
    if ymax is not None:
        va.maximum_scale = ymax
    ca = c.category_axis
    ca.tick_labels.font.size = Pt(8)
    ca.format.line.color.rgb = RGBColor.from_string("BFBFBF")
    ca.has_major_gridlines = False
    return c


def replace_tokens(prs, mapping):
    """'{p:key}' → 쪽번호."""
    pat = re.compile(r"\{p:([a-z0-9_]+)\}")
    for s in prs.slides:
        for t in s._element.iter(qn("a:t")):
            if t.text and "{p:" in t.text:
                t.text = pat.sub(lambda m: str(mapping.get(m.group(1), "")), t.text).replace("(p.)", "")


class Page5(Page):
    """원래 제안서(v5 원본 템플릿) 디자인으로 같은 내용을 그린다.

    제목은 템플릿 제목 틀, 핵심 메시지는 가운데 한 줄 헤드라인, 패널 제목은 템플릿 둥근 띠,
    요점·시사점은 회색 상자, 각주는 슬라이드 노트로, 하단엔 '자료:' 한 줄.
    """
    YB = 6820000

    def __init__(self, slide, proto, title, headline=None, bullets=None, key=None):
        from tools.pptx_kit import HEAD_FONT, bar as _bar, headline as _hl, title as _title  # noqa: F401
        self.s, self.key, self.PR = slide, key, proto
        self.ybot = self.YB
        _title(slide, proto, title)
        self.y = 1330000
        if headline:
            size = 15
            for sz in (15, 14, 13, 12):
                size = sz
                if lines_needed(headline, CW, sz) <= 1:
                    break
            n = lines_needed(headline, CW, size)
            h = int(size * 1.35 * 12700 * n) + 60000
            tb = _hl(slide, headline, y=self.y, size=size, h=h)
            for p in tb.text_frame.paragraphs:
                for r in p.runs:
                    rPr = r._r.get_or_add_rPr()
                    for t in ("a:latin", "a:ea"):
                        rPr.find(qn(t)).set("typeface", HEAD_FONT)
            self.y += h + 130000
        else:
            self.y = 1800000
        if bullets:
            self.bullets(bullets)

    def bullets(self, items, x=None, w=None, size=9.5, y=None):
        x = L if x is None else x
        w = CW if w is None else w
        y0 = self.y if y is None else y
        ps = []
        for it in items:
            if not it:
                continue
            t = f"**{it[0]}:** {it[1]}" if isinstance(it, (tuple, list)) else it
            ps.append({"text": t, "size": size, "bul": "•", "after": 3})
        h = text_h(ps, w - 160000, size) + 60000
        box(self.s, x, y0, w, h, ps, fill=PALE, margin=90000, anchor=MSO_ANCHOR.MIDDLE)
        if y is None:
            self.y = y0 + h + 140000
        return y0 + h

    def panel(self, x, y, w, title, unit=None):
        from tools.pptx_kit import bar as _bar
        _bar(self.s, self.PR, f"{title} {unit}" if unit else title, x, y, w, h=270000, size=11)
        return y + 310000

    def insight(self, title, items, y=None, x=None, w=None, size=9, max_bottom=None):
        x = L if x is None else x
        w = CW if w is None else w
        max_bottom = max_bottom or self.ybot
        ps = []
        for it in items:
            if not it:
                continue
            t = f"**{it[0]}:** {it[1]}" if isinstance(it, (tuple, list)) else it
            ps.append({"text": t, "size": size, "bul": "•", "after": 2})
        h = text_h(ps, w - 200000, size, 2) + 80000
        y0 = (max_bottom - h - 310000) if y is None else y
        yb = self.panel(x, y0, w, title)
        box(self.s, x, yb, w, h, ps, fill=PALE, margin=90000, anchor=MSO_ANCHOR.MIDDLE)
        return y0, yb + h

    def kpis(self, y, items, h=720000):
        n = len(items)
        gw = (CW - (n - 1) * 120000) // n
        for i, (big, lab) in enumerate(items):
            x = L + i * (gw + 120000)
            rect(self.s, x, y, gw, h, fill=PALEBLUE)
            box(self.s, x, y + 60000, gw, 340000, [{"text": big, "size": 17, "bold": True, "color": NAVY,
                                                    "align": PP_ALIGN.CENTER}], margin=20000)
            box(self.s, x, y + 400000, gw, h - 420000, [{"text": lab, "size": 8, "color": SUB,
                                                         "align": PP_ALIGN.CENTER}], margin=30000)
        return y + h

    def foot(self, note=None, source=None):
        from tools.pptx_kit import source as _src
        if source:
            _src(self.s, "자료: " + source.rstrip("."), w=9000000)
        if note:
            nt = self.s.notes_slide.notes_text_frame
            nt.text = ((nt.text + "\n") if nt.text else "") + "[Note] " + note
