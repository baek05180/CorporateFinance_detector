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
        if name.replace(" ", "") in HEADER_WORDS:   # 2줄짜리 머리글의 두 번째 줄
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


HEADER_WORDS = {"매출액", "비율", "비중", "금액", "품목", "구분", "제품", "주요제품", "용도", "사업부문", "매출유형", "구체적용도", "(%)", "%"}


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


def _flatten(body, limit):
    """섹션 본문 → 문단과 표(행마다 'a | b | c')를 이어 붙인 텍스트. AI 요약 재료용."""
    body = re.sub(r"<TITLE[^>]*>.*?</TITLE>", "", body, count=1, flags=re.S | re.I)
    lines, n = [], 0
    for kind, val in blocks(body):
        if kind == "p":
            t = val.strip()
            if len(t) < 2:
                continue
            lines.append(t)
        else:
            for row in val[:40]:
                cells = [c for c in row if c]
                if cells:
                    lines.append(" | ".join(cells))
        n = sum(len(x) + 1 for x in lines)
        if n >= limit:
            break
    return "\n".join(lines)[:limit]


RAW_PARTS = [("overview", ("사업의 개요",), 5000), ("products", ("주요 제품",), 2500),
             ("materials", ("원재료",), 1500), ("sales", ("매출",), 3500)]


def business_raw(xml):
    """사업보고서 'II. 사업의 내용' 1~4절을 압축한 원문 {overview, products, materials, sales}"""
    part = _business_part(xml)
    if part is None:
        return {}
    secs = sections(part, 2)
    out = {}
    for key, kws, limit in RAW_PARTS:
        body = find_section(secs, *kws)
        if body:
            out[key] = _flatten(body, limit)
    return {k: v for k, v in out.items() if v}


def _business_part(xml):
    for m in re.finditer(r"<TITLE[^>]*>(.*?)</TITLE>", xml, re.S | re.I):
        if re.sub(r"\s", "", text_of(m.group(1))).endswith("사업의내용"):
            nxt = re.search(r"<SECTION-1\b", xml[m.end():], re.I)
            return xml[m.end(): m.end() + nxt.start()] if nxt else xml[m.end():]
    return None


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


AGENCY = r"(한국기업평가|한국신용평가|NICE\s*신용평가|나이스신용평가|서울신용평가|한기평|한신평|NICE)"
GRADE = r"(?<![A-Za-z])(AAA|AA[+\-]|AA0|AA|A[+\-]|A0|BBB[+\-]|BBB0|BBB|BB[+\-]|BB0|BB|B[+\-]|B0)(?![A-Za-z0-9가-힣])"
BOND_DOC_VERSION = 2


def bond_schedule(xml):
    """증권신고서(채무증권) 원문 → {"demand_dt": "YYYY-MM-DD", "rating": "AA-", "v": 2} (못 찾은 값은 None)"""
    text = text_of(xml)
    out = {"demand_dt": None, "rating": None, "v": BOND_DOC_VERSION}
    for p in (r"수요예측\s*(?:기간|일시|일자|실시일|기일|일)\s*[:：]?\s*", r"수요예측[^0-9]{0,15}"):
        m = re.search(p + DATE, text)
        if m:
            y, mo, d = (int(x) for x in m.groups()[-3:])
            out["demand_dt"] = f"{y:04d}-{mo:02d}-{d:02d}"
            break
    votes = {}
    for m in re.finditer(AGENCY, text):
        g = re.search(GRADE, text[m.end(): m.end() + 80])
        if g:
            grade = g.group(1).replace("0", "")
            votes[grade] = votes.get(grade, 0) + 1
    if votes:
        out["rating"] = max(votes.items(), key=lambda kv: kv[1])[0]
    return out


# ── 메자닌(CB·BW·EB) 주요사항보고서: 투자자·콜/풋·할증률 ──
MEZZ_DOC_VERSION = 3
_KINDS = [  # (구분, 이름에 들어가는 말) — 위에서부터 먼저 맞는 것
    ("PEF", ("사모투자합자회사", "기관전용사모", "사모투자전문회사", " PEF", "프라이빗에쿼티")),
    ("자산운용", ("자산운용", "투자신탁", "투자자문", "펀드", "신탁업자", "전문투자형사모", "일반사모")),
    ("증권", ("증권", "투자증권", "금융투자")),
    ("신기사·VC", ("인베스트먼트", "벤처", "신기술", "창업투자", "파트너스", "투자조합", "조합", "캐피탈 파트너")),
    ("캐피탈", ("캐피탈",)),
    ("은행", ("은행",)),
    ("보험", ("생명", "화재", "손해보험", "보험")),
    ("공제·연기금", ("공제", "연금", "기금", "공사")),
]
_CORP = re.compile(r"(주식회사|\(주\)|㈜|유한회사|합자회사|Ltd|LTD|Inc|INC|Co\.|Corp|L\.P|LP|LLC|Fund|FUND|GmbH|S\.A)")


def investor_kind(name, relation=""):
    n = name or ""
    if any(k in (relation or "") for k in ("최대주주", "특수관계", "계열", "임원", "대표이사")) and not _CORP.search(n) and len(n.replace(" ", "")) <= 5:
        return "개인"
    for kind, keys in _KINDS:
        if any(k in n for k in keys):
            return kind
    if _CORP.search(n) or len(n.replace(" ", "")) > 5:
        return "법인"
    return "개인"


def _amount(cells):
    best = None
    for c in cells:
        s = c.replace(",", "").replace(" ", "")
        m = re.fullmatch(r"(\d{7,})(원)?", s)
        if m and (best is None or int(m.group(1)) > best):
            best = int(m.group(1))
    return best


def _months(seg):
    m = re.search(r"(\d{1,3})\s*개월", seg)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d{1,2})\s*년\s*(?:이|후|경과|이 되는|되는)", seg)
    return int(m.group(1)) * 12 if m else None


def _start(seg):
    """'발행일로부터 12개월이 되는 날부터' → (12, None) · '2027년 10월 29일부터' → (None, '2027-10-29')"""
    m = re.search(r"발행일\s*(?:로부터|부터|이후)\s*(\d{1,3})\s*개월", seg)
    if m:
        return int(m.group(1)), None
    d = re.search(DATE + r"\s*일?\s*(?:부터|이후)", seg)
    m = re.search(r"(?<!매)(?<!매 )(\d{1,3})\s*개월\s*(?:이|째)\s*되는\s*날", seg)
    if m and (not d or m.start() < d.start()):
        return int(m.group(1)), None
    m = re.search(r"발행일\s*(?:로부터|부터)\s*(\d{1,2})\s*년", seg)
    if m:
        return int(m.group(1)) * 12, None
    m = re.search(DATE + r"\s*일?\s*(?:부터|이후)", seg)
    if m:
        y, mo, d = (int(x) for x in m.groups()[:3])
        return None, f"{y:04d}-{mo:02d}-{d:02d}"
    return None, None


def mezz_doc(xml):
    """전환·신주인수권부·교환사채 발행결정 원문 → 투자자 목록·콜옵션 비율/시작·풋 시작·할증률."""
    out = {"investors": [], "call_pct": None, "call_m": None, "call_from": None, "put_m": None, "put_from": None, "premium": None, "base_price": None,
           "v": MEZZ_DOC_VERSION}
    seen = set()
    for kind, rows in blocks(xml):
        if kind != "table" or len(rows) < 2:
            continue
        head = " ".join(rows[0]).replace(" ", "")
        if "대상자" not in head or "매도청구" in head or "권면" not in head and "발행대상자" not in head:
            continue
        hcells = [c.replace(" ", "") for c in rows[0]]
        rel_i = next((i for i, c in enumerate(hcells) if "관계" in c), None)
        for r in rows[1:]:
            if not r or not r[0] or any(k in r[0].replace(" ", "") for k in ("합계", "합 계", "계", "-")) and len(r[0]) <= 3:
                continue
            amt = _amount(r[1:])
            name = r[0].strip()
            if not amt or name in seen:
                continue
            seen.add(name)
            rel = r[rel_i].strip() if rel_i is not None and rel_i < len(r) else ""
            k = investor_kind(name, rel)
            out["investors"].append({"name": None if k == "개인" else name, "kind": k, "rel": (rel[:30] if rel not in ("-", "—", "") else None), "amt": amt})
    t = text_of(xml)
    m = re.search(r"매도\s*청구권\s*\(?\s*Call\s*Option\s*\)?\s*에\s*관한\s*사항\s*\]?", t)
    if m:
        seg = t[m.end(): m.end() + 1500]
        if re.match(r"\s*[-:：]?\s*(해당\s*사항\s*(없음|없습니다)|없음|-\s*\d)", seg):
            out["call_pct"] = 0.0
        else:
            p = (re.search(r"(?:권면|총액|원금|전자등록\s*(?:총액|금액)|발행\s*금액|사채\s*금액)[^%]{0,25}?(\d{1,3}(?:\.\d+)?)\s*%", seg)
                 or re.search(r"(\d{1,3}(?:\.\d+)?)\s*%\s*(?:에\s*해당|를\s*초과하지|를\s*총\s*한도|를\s*한도|범위|이내|까지|이하)", seg))
            if p and 0 < float(p.group(1)) <= 100:
                out["call_pct"] = float(p.group(1))
            out["call_m"], out["call_from"] = _start(seg)
    if out["call_pct"] is None:
        p = re.search(r"(?:Call\s*Option|콜\s*옵션|매도\s*청구)[^%\d]{0,40}?(\d{1,3}(?:\.\d+)?)\s*%", t)
        if p and 0 < float(p.group(1)) <= 100:
            out["call_pct"] = float(p.group(1))
    m = re.search(r"조기\s*상환\s*청구권\s*\(?\s*Put\s*Option\s*\)?\s*에\s*관한\s*사항\s*\]?", t)
    if m:
        seg = t[m.end(): m.end() + 1200]
        if not re.match(r"\s*[-:：]?\s*(해당\s*사항\s*(없음|없습니다)|없음)", seg):
            out["put_m"], out["put_from"] = _start(seg)
    m = re.search(r"기준\s*(?:주가|가격|가액)\s*의\s*(\d{2,3}(?:\.\d+)?)\s*%", t)
    if m:
        out["premium"] = round(float(m.group(1)) - 100, 1)
    m = None if out["premium"] is not None else re.search(r"할증\s*률?\s*[:：]?\s*(\d{1,3}(?:\.\d+)?)\s*%", t) or re.search(r"(\d{1,3}(?:\.\d+)?)\s*%\s*(?:를|을)?\s*할증", t)
    if m:
        out["premium"] = float(m.group(1))
    m = re.search(r"기준\s*주가\s*[:：]?\s*([\d,]{3,})\s*원", t)
    if m:
        out["base_price"] = int(m.group(1).replace(",", ""))
    return out
