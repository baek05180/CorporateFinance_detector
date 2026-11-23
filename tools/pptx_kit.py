import copy, re
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

NAVY = "05507D"; SKY = "00A5E5"; DGREY = "4B5151"; LGREY = "BDBDBA"; RED = "DA291C"
TXT = "262626"; SUB = "54585A"; PALE = "F2F2F2"; PALEBLUE = "E6EEF3"
BODY_FONT = "맑은 고딕"; HEAD_FONT = "원신한 Bold"
X0 = 451084; W = 9790000; YTOP = 1290000; YBOT = 6880000

PH = re.compile(r"(【[^】]*】|##)")


def _set_font(run, size=None, bold=None, color=None, font=BODY_FONT, italic=None):
    f = run.font
    if size: f.size = Pt(size)
    if bold is not None: f.bold = bold
    if italic is not None: f.italic = italic
    if color: f.color.rgb = RGBColor.from_string(color)
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", font)
    # order: latin must precede ea; python-pptx placed solidFill earlier - fine
    lat = rPr.find(qn("a:latin")); ea = rPr.find(qn("a:ea"))
    rPr.remove(lat); rPr.remove(ea)
    # insert after fill elements
    rPr.append(lat); rPr.append(ea)


def add_runs(par, text, size=10, bold=False, color=TXT, font=BODY_FONT):
    """text may contain **bold** segments and 【placeholder】 (red bold)."""
    parts = re.split(r"(\*\*[^*]+\*\*)", text)
    for part in parts:
        if not part:
            continue
        b = bold
        if part.startswith("**") and part.endswith("**"):
            part = part[2:-2]; b = True
        for seg in PH.split(part):
            if not seg:
                continue
            r = par.add_run(); r.text = seg
            if PH.fullmatch(seg):
                _set_font(r, size, True, RED, font)
            else:
                _set_font(r, size, b, color, font)


def _bullet(par, char, indent_emu, level_marl):
    pPr = par._p.get_or_add_pPr()
    pPr.set("marL", str(level_marl)); pPr.set("indent", str(-indent_emu))
    for t in ("a:buNone", "a:buChar", "a:buFont"):
        for el in pPr.findall(qn(t)):
            pPr.remove(el)
    if char:
        bf = etree.SubElement(pPr, qn("a:buFont")); bf.set("typeface", "Arial")
        bc = etree.SubElement(pPr, qn("a:buChar")); bc.set("char", char)
    else:
        etree.SubElement(pPr, qn("a:buNone"))


def fill_tf(tf, paras, size=10, color=TXT, align=None, space_after=2, line=None):
    """paras: list of str or dict(text, lvl, size, bold, color, align)."""
    tf.clear()
    first = True
    for p in paras:
        if isinstance(p, str):
            p = {"text": p}
        par = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        lvl = p.get("lvl", None)
        sz = p.get("size", size)
        if lvl == 0:
            _bullet(par, "•", 140000, 140000)
        elif lvl == 1:
            _bullet(par, "-", 120000, 300000)
        if p.get("align", align) is not None:
            par.alignment = p.get("align", align)
        par.space_after = Pt(p.get("space_after", space_after))
        if line:
            par.line_spacing = line
        add_runs(par, p["text"], sz, p.get("bold", False), p.get("color", color))
    return tf


def textbox(slide, x, y, w, h, paras, size=10, color=TXT, align=None, anchor=MSO_ANCHOR.TOP,
            margin=36000, fill=None, line_color=None, wrap=True, space_after=2, name=None):
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    if name: tb.name = name
    tf = tb.text_frame; tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = Emu(margin); tf.margin_top = tf.margin_bottom = Emu(margin)
    tf.vertical_anchor = anchor
    tf.auto_size = None
    if fill:
        tb.fill.solid(); tb.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line_color:
        tb.line.color.rgb = RGBColor.from_string(line_color); tb.line.width = Emu(6350)
    fill_tf(tf, paras, size, color, align, space_after)
    return tb


def rect(slide, x, y, w, h, fill=PALE, line=None, shape=1, name=None):
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE if shape == 1 else MSO_SHAPE.ROUNDED_RECTANGLE,
                                Emu(x), Emu(y), Emu(w), Emu(h))
    if name: sh.name = name
    if fill:
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = RGBColor.from_string(line); sh.line.width = Emu(9525)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


class Proto:
    """Holds prototype XML elements copied from the original deck."""
    def __init__(self, prs):
        s4 = prs.slides[3]; s6 = prs.slides[5]
        self.title = [sh for sh in s4.shapes if sh.name == "텍스트 개체 틀 3"][0]._element
        self.bar = [sh for sh in s6.shapes if sh.name == "모서리가 둥근 직사각형 8"][0]._element


def _next_id(slide):
    ids = [int(e.get("id")) for e in slide.shapes._spTree.iter() if e.tag == qn("p:cNvPr")]
    return max(ids + [1]) + 1


def _clone(slide, el):
    new = copy.deepcopy(el)
    for c in new.iter(qn("p:cNvPr")):
        c.set("id", str(_next_id(slide)))
        for ext in c.findall(qn("a:extLst")):
            c.remove(ext)
    slide.shapes._spTree.append(new)
    return new


def set_shape_text(el, text):
    """Replace text of a cloned shape keeping first run formatting; for multi-paragraph title use last para."""
    ps = el.findall(".//" + qn("a:p"))
    target = None
    for p in ps:
        if p.findall(qn("a:r")):
            target = p
    if target is None:
        target = ps[-1]
    runs = target.findall(qn("a:r"))
    for r in runs[1:]:
        target.remove(r)
    runs[0].find(qn("a:t")).text = text


def title(slide, proto, text):
    el = _clone(slide, proto.title)
    set_shape_text(el, text)
    return el


def bar(slide, proto, text, x, y, w, h=270000, size=12):
    el = _clone(slide, proto.bar)
    off = el.find(".//" + qn("a:off")); ext = el.find(".//" + qn("a:ext"))
    off.set("x", str(x)); off.set("y", str(y)); ext.set("cx", str(w)); ext.set("cy", str(h))
    set_shape_text(el, text)
    for r in el.iter(qn("a:rPr")):
        r.set("sz", str(size * 100))
    for r in el.iter(qn("a:endParaRPr")):
        r.set("sz", str(size * 100))
    return el


def headline(slide, text, y=1330000, size=15, h=340000):
    return textbox(slide, X0, y, W, h, [{"text": text, "align": PP_ALIGN.CENTER}], size=size,
                   color="0B1966", anchor=MSO_ANCHOR.MIDDLE, margin=0, name="Headline")


def headline_font(tb):
    for r in tb.text_frame.paragraphs[0].runs:
        rPr = r._r.get_or_add_rPr()
        for t in ("a:latin", "a:ea"):
            rPr.find(qn(t)).set("typeface", HEAD_FONT)


def source(slide, text, y=6905000, x=X0, w=6200000):
    return textbox(slide, x, y, w, 200000, [text], size=8, color="7F7F7F", margin=0, name="Source")


def _border(cell, color="BFBFBF", w=6350):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    els = []
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = etree.Element(qn(tag)); ln.set("w", str(w)); ln.set("cmpd", "sng")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        c = etree.SubElement(sf, qn("a:srgbClr")); c.set("val", color)
        etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")
        els.append(ln)
    # borders must precede fill in tcPr
    for i, ln in enumerate(els):
        tcPr.insert(i, ln)


def table(slide, data, x, y, w, colw, rowh=270000, size=9.5, header_rows=1, first_col=True,
          align_cols=None, fills=None, bold_cells=None, merges=(), name=None, colors=None):
    """data: list of rows (strings). colw: relative widths. fills: dict (r,c)->hex.
    colors: dict (r,c)->hex text color."""
    rows, cols = len(data), len(data[0])
    gf = slide.shapes.add_table(rows, cols, Emu(x), Emu(y), Emu(w), Emu(sum(rowh) if isinstance(rowh, (list, tuple)) else rowh * rows))
    if name: gf.name = name
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    for a in ("firstRow", "bandRow"):
        tblPr.set(a, "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No Style, No Grid
    tot = sum(colw)
    acc = 0
    for i, cw in enumerate(colw):
        wi = int(w * cw / tot) if i < cols - 1 else w - acc
        tbl.columns[i].width = Emu(wi); acc += wi
    rhs = list(rowh) if isinstance(rowh, (list, tuple)) else [rowh] * rows
    for r in range(rows):
        tbl.rows[r].height = Emu(rhs[r])
    for (r0, c0, r1, c1) in merges:
        tbl.cell(r0, c0).merge(tbl.cell(r1, c1))
    fills = fills or {}; bold_cells = bold_cells or set(); colors = colors or {}
    for r in range(rows):
        for c in range(cols):
            cell = tbl.cell(r, c)
            if cell.is_spanned:
                _border(cell)
                continue
            txt = data[r][c]
            hdr = r < header_rows
            fc = c == 0 and first_col and not hdr
            cell.margin_left = cell.margin_right = Emu(60000)
            cell.margin_top = cell.margin_bottom = Emu(25000)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            col = colors.get((r, c), "FFFFFF" if hdr else TXT)
            paras = txt.split("\n") if isinstance(txt, str) else txt
            al = PP_ALIGN.CENTER if (hdr or fc) else (align_cols or {}).get(c, PP_ALIGN.LEFT)
            fill_tf(cell.text_frame, [{"text": t} if isinstance(t, str) else t for t in paras], size=size,
                    color=col, align=al, space_after=0)
            if hdr or fc or (r, c) in bold_cells:
                for p in cell.text_frame.paragraphs:
                    for rn in p.runs:
                        if rn.font.color.rgb != RGBColor.from_string(RED):
                            rn.font.bold = True
            f = fills.get((r, c), NAVY if hdr else (PALE if fc else None))
            if f:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor.from_string(f)
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
            _border(cell)
    return gf


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def replace_in_slide(slide, old, new, count_required=1):
    """Replace text within paragraphs; merges runs of the paragraph if needed (keeps first run format)."""
    n = 0
    for p in slide.shapes._spTree.iter(qn("a:p")):
        runs = p.findall(qn("a:r"))
        if not runs:
            continue
        joined = "".join((r.find(qn("a:t")).text or "") for r in runs)
        if old not in joined:
            continue
        # try single run
        done = False
        for r in runs:
            t = r.find(qn("a:t"))
            if t.text and old in t.text:
                t.text = t.text.replace(old, new); done = True; break
        if not done:
            # find run index where old starts; merge from that run to the run where old ends
            starts = []; pos = 0
            for r in runs:
                starts.append(pos); pos += len(r.find(qn("a:t")).text or "")
            i0 = joined.index(old); i1 = i0 + len(old)
            a = max(k for k in range(len(runs)) if starts[k] <= i0)
            b = max(k for k in range(len(runs)) if starts[k] < i1)
            seg = "".join((runs[k].find(qn("a:t")).text or "") for k in range(a, b + 1))
            runs[a].find(qn("a:t")).text = seg.replace(old, new, 1)
            for k in range(a + 1, b + 1):
                p.remove(runs[k])
        n += 1
    if n < count_required:
        raise ValueError(f"not found: {old!r}")
    return n
