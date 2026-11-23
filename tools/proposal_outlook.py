"""리서치 금리 전망 엑셀 읽기.

비공개 저장소 proposal-vault 의 research/ 폴더에 있는 가장 최근 .xlsx 를 읽는다 (공개 저장소에는 두지 않는다).

엑셀 모양 (리서치센터 정기 전망)
  - 시트 하나가 전망 하나. 시트 이름이 전망일('11.10' = 11월 10일). 가장 앞(가장 최근) 시트를 쓴다.
  - 3행: 연도(2027년·2028년·2029년 …), 4행: 기간(11월·12월·1Q~4Q·기말), 마지막 열: '전망 근거 Comment'
  - C열: 구분(한국·미국·환율·CP·CD …, 아래로 이어짐), D열: 지표(기준금리·국채 3Y·AA- 3Y …)
행·열 위치가 바뀌어도 글자로 찾으므로 그대로 읽힌다.
"""
import glob
import os
import re
from datetime import date, timedelta

from tools import proposal_mkt as M


def _qend(y, q):
    return date(y, 3 * q, 31 if q in (1, 4) else 30)


def _mend(y, m):
    return (date(y + m // 12, m % 12 + 1, 1) - timedelta(days=1))


def find_file(folder):
    if not folder or not os.path.isdir(folder):
        return None
    fs = [f for f in glob.glob(os.path.join(folder, "*.xlsx")) if not os.path.basename(f).startswith("~$")]
    return max(fs, key=os.path.getmtime) if fs else None


def load_latest(folder, today=None):
    """폴더 안 엑셀 중 전망일이 가장 최근인 것 (git 체크아웃은 파일 시각이 같아 시각으로 고르지 않음)."""
    if not folder or not os.path.isdir(folder):
        return None
    best = None
    for f in glob.glob(os.path.join(folder, "*.xlsx")):
        if os.path.basename(f).startswith("~$"):
            continue
        try:
            o = load(f, today)
        except Exception:  # noqa: BLE001
            continue
        if o and (best is None or o["vintage"] > best["vintage"]):
            best = o
    return best


def _vintage(sheet_name, file_name, today):
    m = re.match(r"\s*(\d{1,2})\.(\d{1,2})\s*$", sheet_name or "")
    if not m:
        return None
    mo, dd = int(m.group(1)), int(m.group(2))
    fy = re.search(r"(\d{2})(\d{2})(\d{2})", os.path.basename(file_name or ""))
    y = 2000 + int(fy.group(1)) if fy and int(fy.group(2)) == mo else today.year
    try:
        d = date(y, mo, dd)
    except ValueError:
        return None
    if d > today + timedelta(days=1):
        d = date(y - 1, mo, dd)
    return d


def load(path, today=None):
    """→ {"vintage": date, "cols": [(label, end_date)], "rows": {key: [값…]}, "comments": {kr, credit, us, fx}} 또는 None"""
    if not path or not os.path.exists(path):
        return None
    import openpyxl
    today = today or date.today()
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    ws, vint = None, None
    for w in wb.worksheets:  # 앞쪽이 최신
        v = _vintage(w.title, path, today)
        if v:
            ws, vint = w, v
            break
    if ws is None:
        return None
    # 머리글: '1Q' 같은 기간이 있는 행
    hdr = None
    for r in range(1, 12):
        vals = [str(c.value or "").strip() for c in ws[r]]
        if any(re.fullmatch(r"[1-4]Q", v) for v in vals):
            hdr = r
            break
    if hdr is None:
        return None
    year_row = [str(c.value or "").strip() for c in ws[hdr - 1]]
    per_row = [str(c.value or "").strip() for c in ws[hdr]]
    ncol = len(per_row)
    cols, comment_col, cur_y = [], None, None
    for j in range(ncol):
        yv = re.search(r"(20\d{2})", year_row[j] if j < len(year_row) else "")
        if yv:
            cur_y = int(yv.group(1))
        p = per_row[j]
        head = (year_row[j] if j < len(year_row) else "") + p
        if "Comment" in head or "코멘트" in head or "근거" in head:
            comment_col = j
            continue
        mm = re.fullmatch(r"(\d{1,2})월", p)
        if mm:
            m_ = int(mm.group(1))
            y_ = vint.year if m_ >= vint.month else vint.year + 1
            cols.append((j, f"'{y_ % 100}.{m_}월", _mend(y_, m_)))
        elif re.fullmatch(r"[1-4]Q", p) and cur_y:
            cols.append((j, f"'{cur_y % 100}.{p}", _qend(cur_y, int(p[0]))))
        elif p == "기말" and cur_y:
            cols.append((j, f"'{cur_y % 100}말", date(cur_y, 12, 31)))
    if comment_col is None:
        for j in range(ncol):
            if "Comment" in str(ws.cell(hdr - 1, j + 1).value or ""):
                comment_col = j
    # 본문
    rows, comments, sec = {}, {}, ""
    label_col = None
    for r in range(hdr + 1, ws.max_row + 1):
        cells = [ws.cell(r, j + 1).value for j in range(ncol)]
        texts = [(j, str(v).strip()) for j, v in enumerate(cells) if isinstance(v, str) and v.strip()]
        if not texts:
            continue
        if label_col is None:  # 숫자가 처음 나오는 열 바로 앞 글자가 지표 이름
            first_num = next((j for j, v in enumerate(cells) if isinstance(v, (int, float))), None)
            if first_num is None:
                continue
            label_col = first_num - 1
        lab = str(cells[label_col] or "").strip()
        sc = next((t for j, t in texts if j < label_col), None)
        if sc:
            sec = re.sub(r"\s+", " ", sc)
        key = _key(sec, lab)
        if key:
            rows[key] = [cells[j] if isinstance(cells[j], (int, float)) else None for j, _, _ in cols]
        if comment_col is not None and isinstance(cells[comment_col], str) and cells[comment_col].strip():
            ck = "credit" if key and key[0] == "CR" else {"KR": "kr", "US": "us", "FX": "fx"}.get(key[0] if key else "", None)
            if ck and ck not in comments:
                comments[ck] = cells[comment_col].strip()
    if not rows:
        return None
    return {"vintage": vint, "cols": [(lab, d) for _, lab, d in cols], "rows": rows, "comments": comments,
            "file": os.path.basename(path), "sheet": ws.title}


def _key(sec, lab):
    s, l_ = sec.replace(" ", ""), re.sub(r"\s+", " ", lab).strip()
    m = re.match(r"^(AAA|AA[+\-0]?|A[+\-0]?|BBB[+\-0]?)\s*(\d{1,2})Y$", l_.replace(" ", " "))
    if m:
        return ("CR", M.gkey(m.group(1)), int(m.group(2)))
    m = re.match(r"^국채\s*(\d{1,2})Y$", l_)
    if s.startswith("한국") and m:
        return ("KR", "KTB", int(m.group(1)))
    if s.startswith("한국") and "기준금리" in l_:
        return ("KR", "base")
    if s.startswith("미국") and m:
        return ("US", "UST", int(m.group(1)))
    if s.startswith("미국") and "기준금리" in l_:
        return ("US", "base")
    if s.startswith("환율"):
        return ("FX", "usdkrw")
    if s.startswith("CP"):
        return ("CP", 91)
    if s.startswith("CD"):
        return ("CD", 91)
    return None


def series(O, *key):
    return O["rows"].get(tuple(key))


def grade_path(O, grade, tenor, curves=None):
    """동사 등급·만기 전망 경로. 등급 행이 없으면 국고 전망 + 현재 스프레드."""
    g = M.gkey(grade)
    tenors = sorted({k[2] for k in O["rows"] if k[0] == "CR" and k[1] == g})
    if tenors:
        lo = max([t for t in tenors if t <= tenor] or [tenors[0]])
        hi = min([t for t in tenors if t >= tenor] or [tenors[-1]])
        a, b = O["rows"][("CR", g, lo)], O["rows"][("CR", g, hi)]
        w = 0 if hi == lo else (tenor - lo) / (hi - lo)
        return [None if x is None or y is None else x + (y - x) * w for x, y in zip(a, b)], "리서치 등급 전망"
    k = ktb_path(O, tenor)
    if k and curves:
        sp = curves.sp(g, tenor)
        if sp is not None:
            return [None if x is None else x + sp / 100 for x in k], "국고 전망 + 현재 스프레드"
    return None, ""


def ktb_path(O, tenor):
    ks = sorted({k[2] for k in O["rows"] if k[0] == "KR" and k[1] == "KTB"})
    if not ks:
        return None
    lo = max([t for t in ks if t <= tenor] or [ks[0]])
    hi = min([t for t in ks if t >= tenor] or [ks[-1]])
    a, b = O["rows"][("KR", "KTB", lo)], O["rows"][("KR", "KTB", hi)]
    w = 0 if hi == lo else (tenor - lo) / (hi - lo)
    return [None if x is None or y is None else x + (y - x) * w for x, y in zip(a, b)]


def at(O, path, d):
    """d 가 속한 기간(그 날 이후 처음 끝나는 칸)의 값."""
    if not path:
        return None, None
    for (lab, end), v in zip(O["cols"], path):
        if end >= d and v is not None:
            return lab, v
    return None, None


def first_sentences(txt, n=170):
    """코멘트 앞부분 요약 (문장 단위로 n자 이내)."""
    t = re.sub(r"\s+", " ", str(txt or "")).strip().lstrip("-").strip()
    out = ""
    for s in re.split(r"(?<=[.다])\s+", t):
        if len(out) + len(s) > n:
            break
        out += (" " if out else "") + s
    return out or (t[: n - 1] + "…")
