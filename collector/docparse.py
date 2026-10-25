"""DART 공시 원문(XML) 파서.

사업보고서 'II. 사업의 내용'에서 사업 개요·주요 제품·주요 매출처를 뽑고,
채무증권 증권신고서에서 수요예측일·청약일·납입일을 뽑는다.
원문 서식이 회사마다 달라서 '찾으면 보여주고, 못 찾으면 비운다'는 원칙으로 만든다.
"""
import html
import re

TAG = re.compile(r"<[^>]+>")
WS = re.compile(r"[ \t\r\f\v 　]+")


def text_of(fragment):
    t = TAG.sub(" ", fragment)
    t = html.unescape(t)
    return WS.sub(" ", t).strip()


def blocks(xml):
    """XML 조각 → [("p", text) | ("table", [[cell,...], ...])] 순서대로."""
    out = []
    pos = 0
    pat = re.compile(r"<TABLE\b.*?</TABLE>|<P\b[^>]*>.*?</P>", re.S | re.I)
    for m in pat.finditer(xml):
        frag = m.group(0)
        if frag[:6].upper() == "<TABLE":
            rows = []
            for tr in re.findall(r"<TR\b.*?</TR>", frag, re.S | re.I):
                cells = [text_of(c) for c in re.findall(r"<(?:TD|TH|TE|TU)\b[^>]*>(.*?)</(?:TD|TH|TE|TU)>", tr, re.S | re.I)]
                if any(cells):
                    rows.append(cells)
            if rows:
                out.append(("table", rows))
        else:
            t = text_of(frag)
            if t:
                out.append(("p", t))
        pos = m.end()
    return out


def sections(xml, level=2):
    """{제목: 본문 XML} — 같은 수준의 다음 섹션 시작 전까지를 본문으로 본다 (닫는 태그가 없어도 동작)."""
    out = {}
    starts = [m.start() for m in re.finditer(rf"<SECTION-{level}\b", xml, re.I)]
    for i, st in enumerate(starts):
        body = xml[st: starts[i + 1] if i + 1 < len(starts) else len(xml)]
        t = re.search(r"<TITLE[^>]*>(.*?)</TITLE>", body, re.S | re.I)
        title = text_of(t.group(1)) if t else ""
        out.setdefault(title, body)
    return out


def find_section(secs, *keywords):
    for title, body in secs.items():
        norm = title.replace(" ", "")
        if all(k.replace(" ", "") in norm for k in keywords):
            return body
    return None


def _clean_para(t):
    t = re.sub(r"^\s*[가-하]\.\s*", "", t)            # '가. ' 같은 머리기호
    t = re.sub(r"^\s*\(\s*\d+\s*\)\s*", "", t)
    t = re.sub(r"^(?:\d+\.\s*)?(?:사업의\s*개요|업계의\s*현황|회사의\s*현황)\s*", "", t)
    return t.strip()


BOILER = ("작성기준일", "해당사항", "기재하였습니다", "참고하시기", "상기", "단위 :", "(단위", "※", "기업공시서식",
          "연결실체", "본 보고서", "기재생략", "해당없음")


def business_overview(body, limit=900):
    paras = []
    body = re.sub(r"<TITLE[^>]*>.*?</TITLE>", "", body, count=1, flags=re.S | re.I)
    for kind, val in blocks(body):
        if kind != "p":
            continue
        t = _clean_para(val)
        if len(t) < 40 or any(b in t[:40] for b in BOILER):
            continue
        paras.append(t)
        if sum(len(p) for p in paras) >= limit:
            break
    text = "\n".join(paras)
    return text[:limit].rstrip() + ("…" if len(text) > limit else "")


def _num(s):
    s = (s or "").replace(",", "").replace("%", "").strip()
    s = s.replace("(", "-").replace(")", "")
    try:
        return float(s)
    except ValueError:
        return None


PRODUCT_HEAD = ("품목", "제품", "주요제품", "주요 제품", "상품", "서비스", "구분", "사업부문", "부문", "품 목", "제 품")
SHARE_HEAD = ("비율", "비중", "%")


def products(body, max_rows=8):
    """주요 제품 표 → [{"name":..., "desc":..., "share":...}]"""
    best = None
    for kind, rows in blocks(body):
        if kind != "table" or len(rows) < 2:
            continue
        head = [c.replace(" ", "") for c in rows[0]]
        if not any(any(h.replace(" ", "") in c for h in PRODUCT_HEAD) for c in head):
            continue
        share_col = next((i for i, c in enumerate(head) if any(s in c for s in SHARE_HEAD)), None)
        score = (share_col is not None) * 2 + min(len(rows), 10) / 10
        if not best or score > best[0]:
            best = (score, rows, share_col)
    if not best:
        return []
    _, rows, share_col = best
    head = [c.replace(" ", "") for c in rows[0]]
    name_col = next((i for i, c in enumerate(head) if any(k in c for k in ("품목", "제품", "상품", "서비스"))), 0)
    desc_col = next((i for i, c in enumerate(head) if any(k in c for k in ("용도", "기능", "특징", "주요내용", "설명"))), None)
    seg_col = next((i for i, c in enumerate(head) if any(k in c for k in ("사업부문", "부문", "구분"))), None)
    out = []
    prev = rows[0]
    width = len(rows[0])
    seen = set()
    for r in rows[1:]:
        if len(r) < width and len(prev) >= width:
            r = prev[:width - len(r)] + r          # 앞 칸이 세로 병합(rowspan)된 행
        prev = r
        if len(r) <= name_col:
            continue
        name = r[name_col]
        joined = " ".join(r).replace(" ", "")
        if (not name or any(t in joined for t in ("합계", "총계", "소계", "내부거래"))
                or re.fullmatch(r"[\d,.\-()%\s]+", name)):
            continue
        ident = (r[seg_col] if seg_col is not None and seg_col < len(r) else "", name)
        if ident in seen:          # 세로 병합으로 같은 제품이 반복된 행
            continue
        seen.add(ident)
        item = {"name": name[:60]}
        if seg_col is not None and seg_col != name_col and seg_col < len(r) and r[seg_col] and r[seg_col] != name:
            item["segment"] = r[seg_col][:30]
        if desc_col is not None and desc_col < len(r) and r[desc_col]:
            item["desc"] = r[desc_col][:80]
        if share_col is not None and share_col < len(r):
            v = _num(r[share_col])
            if v is not None and 0 <= v <= 100:
                item["share"] = round(v, 1)
        out.append(item)
        if len(out) >= max_rows:
            break
    return out


CUSTOMER_KEYS = ("주요매출처", "주요 매출처", "주요거래처", "주요 거래처", "주요고객", "주요 고객", "매출처")


def customers(body, limit=400):
    """매출 및 수주상황 → 주요 매출처 문장/표 발췌 (문자열)"""
    bl = blocks(body)
    for i, (kind, val) in enumerate(bl):
        if kind == "p" and any(k in val for k in CUSTOMER_KEYS):
            # 제목성 문장이면 다음 문단/표까지 붙인다
            parts = [val]
            for k2, v2 in bl[i + 1:i + 4]:
                if k2 == "p":
                    parts.append(v2)
                else:
                    parts.append(" / ".join(" ".join(c for c in r if c) for r in v2[:6]))
                if sum(len(p) for p in parts) > limit:
                    break
            text = " ".join(parts)
            return text[:limit].rstrip() + ("…" if len(text) > limit else "")
        if kind == "table":
            head = " ".join(val[0]).replace(" ", "")
            if any(k.replace(" ", "") in head for k in CUSTOMER_KEYS):
                text = " / ".join(" ".join(c for c in r if c) for r in val[:6])
                return text[:limit]
    return ""


def annual_business(xml):
    """사업보고서 원문 → {"overview","products","customers"}"""
    part = None
    for m in re.finditer(r"<TITLE[^>]*>(.*?)</TITLE>", xml, re.S | re.I):
        if re.sub(r"\s", "", text_of(m.group(1))).endswith("사업의내용"):
            nxt = re.search(r"<SECTION-1\b", xml[m.end():], re.I)
            part = xml[m.end(): m.end() + nxt.start()] if nxt else xml[m.end():]
            break
    if part is None:
        return {}
    secs = sections(part, 2)
    out = {}
    ov = find_section(secs, "사업의 개요")
    if ov:
        out["overview"] = business_overview(ov)
    pr = find_section(secs, "주요 제품")
    if pr:
        out["products"] = products(pr)
    sa = find_section(secs, "매출")
    if sa:
        out["customers"] = customers(sa)
    if not out.get("customers"):
        for body in secs.values():
            c = customers(body)
            if c:
                out["customers"] = c
                break
    return {k: v for k, v in out.items() if v}


# ── 채무증권 증권신고서 일정 ──
DATE = r"(20\d{2})\s*[년.\-/]\s*(\d{1,2})\s*[월.\-/]\s*(\d{1,2})"


def _iso(m):
    y, mo, d = (int(x) for x in m.groups()[:3])
    return f"{y:04d}-{mo:02d}-{d:02d}"


def bond_schedule(xml):
    """증권신고서(채무증권) 원문 → {"demand_dt": "YYYY-MM-DD"} (찾지 못하면 빈 dict)"""
    text = text_of(xml)
    for p in (r"수요예측\s*(?:기간|일시|일자|실시일|기일|일)\s*[:：]?\s*", r"수요예측[^0-9]{0,15}"):
        m = re.search(p + DATE, text)
        if m:
            y, mo, d = (int(x) for x in m.groups()[-3:])
            return {"demand_dt": f"{y:04d}-{mo:02d}-{d:02d}"}
    return {}
