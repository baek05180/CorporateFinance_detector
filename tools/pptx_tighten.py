"""상단 회색 요약 박스를 마지막 줄까지만 남기고, 아래 차트·표를 그만큼 키운다.

LibreOffice(soffice)와 pdftoppm 이 있어야 동작한다. 없으면 아무것도 바꾸지 않는다.
"""
import glob
import os
import shutil
import subprocess
import sys
import tempfile

from pptx import Presentation
from pptx.util import Emu

DPI = 110


def _soffice():
    wrapper = os.environ.get("SOFFICE_WRAPPER")
    if wrapper and os.path.exists(wrapper):
        return [sys.executable, wrapper]
    for b in ("soffice", "libreoffice"):
        if shutil.which(b):
            return [b]
    return None


def render(path, outdir):
    cmd = _soffice()
    if not cmd or not shutil.which("pdftoppm"):
        raise RuntimeError("soffice/pdftoppm not found")
    subprocess.run(cmd + ["--headless", "--convert-to", "pdf", "--outdir", outdir, path], capture_output=True, timeout=300)
    pdf = glob.glob(f"{outdir}/*.pdf")[0]
    subprocess.run(["pdftoppm", "-r", str(DPI), "-png", pdf, f"{outdir}/p"], check=True)
    return sorted(glob.glob(f"{outdir}/p-*.png"))


def e2p(e):
    return int(e / 914400 * DPI)


def text_bottom(img, sh):
    x0, y0 = e2p(sh.left) + 6, e2p(sh.top) + 4
    x1, y1 = e2p(sh.left + sh.width) - 6, e2p(sh.top + sh.height) - 4
    px = img.load()
    last = None
    for y in range(y0, min(y1, img.height)):
        for x in range(x0, min(x1, img.width), 2):
            r, g, b = px[x, y][:3]
            if r < 140 and g < 140 and b < 140:
                last = y
                break
    return last


def scale_table(gf, new_h):
    tot = sum(r.height for r in gf.table.rows)
    f = new_h / tot
    for r in gf.table.rows:
        r.height = Emu(int(r.height * f))
    gf.height = Emu(new_h)


def summary_box(slide):
    c = [sh for sh in slide.shapes if sh.has_text_frame and sh.top < 1900000 and sh.width > 9000000 and sh.height > 900000]
    return c[0] if c else None


def tighten_slide(slide, img, box, pad):
    lb = text_bottom(img, box)
    if lb is None:
        return 0
    need = int((lb + 1) / DPI * 914400) - box.top + pad
    delta = box.height - need
    if delta < 150000:
        return 0
    old_bottom = box.top + box.height
    box.height = Emu(need)
    below = [sh for sh in slide.shapes if sh is not box and sh.top >= old_bottom - 50000]
    big_all = sorted([sh for sh in below if (getattr(sh, "has_chart", False) or getattr(sh, "has_table", False))], key=lambda x: x.top)
    big = []
    for sh in big_all:  # 열마다 가장 위의 차트·표만 늘린다
        if not any(o.left < sh.left + sh.width and sh.left < o.left + o.width for o in big):
            big.append(sh)
    charts = [sh for sh in big if sh.has_chart]
    for sh in below:
        if sh in big_all:
            continue
        host = next((c for c in charts if c.left - 50000 <= sh.left <= c.left + c.width and c.top - 300000 <= sh.top <= c.top + c.height), None)
        if host is not None and sh.top > host.top:
            f = (host.height + delta) / host.height
            sh.top = Emu(int(host.top - delta + (sh.top - host.top) * f))
        elif not any(sh.top >= b.top + b.height - 10000 for b in big if b.left <= sh.left < b.left + b.width) \
                and not (sh.has_text_frame and sh.text_frame.text.startswith(("자료", "주)"))):
            sh.top = Emu(sh.top - delta)
    for sh in big:
        sh.top = Emu(sh.top - delta)
        if sh.has_table:
            scale_table(sh, sh.height + delta)
        else:
            sh.height = Emu(sh.height + delta)
    return delta


def tighten_file(path, out, pad=90000):
    from PIL import Image
    prs = Presentation(path)
    targets = [(i, s) for i, s in enumerate(prs.slides) if summary_box(s) is not None]
    if not targets:
        if out != path:
            prs.save(out)
        return []
    with tempfile.TemporaryDirectory() as td:
        imgs = render(path, td)
        done = []
        for i, s in targets:
            img = Image.open(imgs[i]).convert("RGB")
            d = tighten_slide(s, img, summary_box(s), pad)
            if d:
                done.append((i + 1, round(d / 914400, 2)))
    prs.save(out)
    return done


if __name__ == "__main__":
    print(tighten_file(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else sys.argv[1]))
