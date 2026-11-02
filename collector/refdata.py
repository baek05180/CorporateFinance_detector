"""기준 데이터: 상장사 목록(KIND), 주가·시가총액(공공데이터포털)."""
import html
import logging
import re
import time
from datetime import timedelta
from urllib.parse import unquote

import requests

from . import config
from .datagokr import redact

log = logging.getLogger(__name__)

MARKET = {"유가증권": "Y", "유가": "Y", "코스피": "Y", "코스닥": "K", "코넥스": "N"}


def _cells(tr):
    return [html.unescape(re.sub(r"<[^>]+>", "", c)).strip()
            for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.S | re.I)]


def kind_listed(session=None):
    """KIND 상장법인목록 → {종목코드: {name, market, industry, products, listed, fiscal_month, ceo, homepage, region}}"""
    s = session or requests.Session()
    r = s.post(config.KIND_LIST_URL, timeout=60, headers={"User-Agent": "Mozilla/5.0"},
               data={"method": "download", "searchType": "13", "orderMode": "3", "orderStat": "D"})
    r.raise_for_status()
    text = r.content.decode("euc-kr", "ignore")
    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", text, re.S | re.I)
    if not rows:
        raise ValueError("KIND 목록 형식이 바뀌었습니다")
    head = _cells(rows[0])
    idx = {h: i for i, h in enumerate(head)}
    need = ["회사명", "시장구분", "종목코드", "업종", "주요제품"]
    if not all(n in idx for n in need):
        raise ValueError(f"KIND 목록 열 이름이 다릅니다: {head}")
    out = {}
    for tr in rows[1:]:
        c = _cells(tr)
        if len(c) < len(head):
            continue
        code = c[idx["종목코드"]].strip().zfill(6)
        get = lambda k: c[idx[k]] if k in idx else ""  # noqa: E731
        out[code] = {
            "name": get("회사명"),
            "market": MARKET.get(get("시장구분"), get("시장구분")),
            "industry": get("업종"),
            "products": get("주요제품"),
            "listed": get("상장일"),
            "fiscal_month": get("결산월"),
            "ceo": get("대표자명"),
            "homepage": get("홈페이지"),
            "region": get("지역"),
        }
    return out


class StockPriceError(Exception):
    pass


def _key():
    key = (config.DATA_GO_KR_KEY or "").strip()
    if not key:
        raise StockPriceError("DATA_GO_KR_KEY 없음")
    return unquote(key) if "%" in key else key


def prices_on(d, session=None, key=None):
    """특정 거래일 전 종목 시세 → {종목코드: {close, mcap, shares}} (휴장일이면 빈 dict)"""
    key = key or _key()
    s = session or requests.Session()
    out, page = {}, 1
    while True:
        for attempt in range(4):
            time.sleep(0.2 if attempt == 0 else 5 * attempt)
            try:
                r = s.get(config.STOCK_PRICE_URL, timeout=40, params=dict(
                    serviceKey=key, resultType="json", numOfRows=1000, pageNo=page, basDt=d.strftime("%Y%m%d")))
            except requests.RequestException as e:
                err = StockPriceError(f"연결 실패: {redact(e)[:200]}")
                continue
            if r.status_code in (502, 503, 504) or "SERVICETIMEOUT" in r.text[:400]:
                err = StockPriceError(f"HTTP {r.status_code}: 일시 오류")
                continue
            break
        else:
            raise err
        if r.status_code >= 400 or r.text.lstrip().startswith("<"):
            raise StockPriceError(f"HTTP {r.status_code}: {redact(r.text[:300])}")
        body = r.json().get("response", {}).get("body", {})
        items = (body.get("items") or {})
        items = items.get("item", []) if isinstance(items, dict) else (items or [])
        if isinstance(items, dict):
            items = [items]
        total = int(body.get("totalCount") or 0)
        for it in items:
            code = str(it.get("srtnCd") or "").strip()
            if len(code) == 7 and code[0] == "A":
                code = code[1:]
            try:
                out[code] = {"close": int(float(it.get("clpr"))), "mcap": int(float(it.get("mrktTotAmt"))),
                             "shares": int(float(it.get("lstgStCnt"))),
                             "vol": int(float(it.get("trqu") or 0)), "val": int(float(it.get("trPrc") or 0))}
            except (TypeError, ValueError):
                continue
        if not items or len(out) >= total or page >= 6:
            return out
        page += 1


def stock_prices(today, session=None):
    """가장 최근 거래일의 전 종목 종가·시가총액 → (기준일, {종목코드: {close, mcap, shares}})"""
    key, s = _key(), session or requests.Session()
    d = today
    for _ in range(10):
        d -= timedelta(days=1)
        if d.weekday() >= 5:
            continue
        out = prices_on(d, s, key)
        if out:
            return d.isoformat(), out
    return None, {}


def update_price_history(hist, today, latest=None, days=180, max_fetch=40, session=None):
    """종가 이력(최근 days일) 갱신. hist = {"dates": [...], "close": {code: [...]}} (dates 오름차순, 값은 dates와 같은 순서).
    아직 없는 평일만 조회하며, 한 번에 최대 max_fetch일까지만 채운다(나머지는 다음 실행 때)."""
    key, s = _key(), session or requests.Session()
    have = dict(zip(hist.get("dates") or [], range(len(hist.get("dates") or []))))
    rows = {d: {c: v[i] for c, v in (hist.get("close") or {}).items() if i < len(v) and v[i] is not None}
            for d, i in have.items()}
    if latest:
        rows[latest[0]] = {c: v["close"] for c, v in latest[1].items()}
    start = today - timedelta(days=days)
    tried = set(hist.get("closed") or [])
    fetched, d = 0, today - timedelta(days=1)
    while d >= start and fetched < max_fetch:
        iso = d.isoformat()
        if d.weekday() < 5 and iso not in rows and iso not in tried:
            out = prices_on(d, s, key)
            fetched += 1
            if out:
                rows[iso] = {c: v["close"] for c, v in out.items()}
            elif d < today - timedelta(days=4):
                tried.add(iso)  # 휴장일 (최근 며칠은 아직 미반영일 수 있어 다시 시도)
        d -= timedelta(days=1)
    dates = sorted(x for x in rows if x >= start.isoformat())
    codes = sorted({c for x in dates for c in rows[x]})
    close = {c: [rows[x].get(c) for x in dates] for c in codes}
    return {"dates": dates, "close": close, "closed": sorted(x for x in tried if x >= start.isoformat())}, fetched


def update_px(today, latest=None, years=3, max_fetch=40, session=None):
    """3년 일별 시세를 월별 파일(ref/px/YYYY-MM.json)에 쌓는다: {"dates", "c": 종가, "v": 거래량, "a": 거래대금(백만원), "sh": 상장주식수}.
    지난 달 파일은 다 차면 다시 쓰지 않으므로 저장소가 크게 늘지 않는다. 최근 날짜부터 채운다."""
    from .store import dump_json, load_json
    key, s = _key(), session or requests.Session()
    months, tried_all = {}, load_json("ref/px/closed.json", [])
    tried = set(tried_all)

    def month(m):
        if m not in months:
            months[m] = load_json(f"ref/px/{m}.json", {"dates": [], "c": {}, "v": {}, "a": {}, "sh": {}})
        return months[m]

    def put(iso, out):
        f = month(iso[:7])
        if iso in f["dates"]:
            return
        f["dates"].append(iso)
        f["dates"].sort()
        i = f["dates"].index(iso)
        n = len(f["dates"])
        for k, fld in (("c", "close"), ("v", "vol"), ("a", "val")):
            for code in set(f[k]) | set(out):
                arr = f[k].setdefault(code, [None] * (n - 1))
                arr.insert(i, (out[code][fld] // 1_000_000 if fld == "val" else out[code][fld]) if code in out else None)
        for code, v in out.items():
            f["sh"][code] = v["shares"]

    if latest and latest[0]:
        put(latest[0], latest[1])
    start = today - timedelta(days=365 * years + 5)
    fetched, d = 0, today - timedelta(days=1)
    dirty = set(months)
    fails = 0
    while d >= start and fetched < max_fetch and fails < 15:
        iso = d.isoformat()
        if d.weekday() < 5 and iso not in tried and iso not in month(iso[:7])["dates"]:
            try:
                out = prices_on(d, s, key)
            except StockPriceError:
                fails += 1          # 이 날은 다음 실행에서 다시
                d -= timedelta(days=1)
                continue
            fetched += 1
            if out:
                put(iso, out)
                dirty.add(iso[:7])
            elif d < today - timedelta(days=4):
                tried.add(iso)
        d -= timedelta(days=1)
    for m in dirty:
        dump_json(f"ref/px/{m}.json", months[m], indent=None)
    dump_json("ref/px/closed.json", sorted(tried), indent=None)
    return fetched


# ── 공매도(KRX, 로그인 필요)·대차잔고(금융위) 일별 → 월별 파일 ──
LNB_URL = "https://apis.data.go.kr/1160100/service/GetCMStckLnbInfoService/getStckLnbDetail"
KRX_JSON = "https://data.krx.co.kr/comm/bldAttendant/getJsonData.cmd"


def krx_session():
    import os
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
                      "Referer": "https://data.krx.co.kr/contents/MDC/MDI/mdiLoader/index.cmd"})
    s.get("https://data.krx.co.kr/contents/MDC/MDI/mdiLoader/index.cmd", timeout=30)
    s.get("https://data.krx.co.kr/contents/MDC/COMS/client/MDCCOMS001.cmd", timeout=30)
    lg = dict(mbrNm="", telNo="", di="", certType="", mbrId=os.environ.get("KRX_ID", ""), pw=os.environ.get("KRX_PW", ""))
    url = "https://data.krx.co.kr/contents/MDC/COMS/client/MDCCOMS001D1.cmd"
    r = s.post(url, data=lg, timeout=30)
    if "CD011" in r.text:
        r = s.post(url, data={**lg, "skipDup": "Y"}, timeout=30)
    if "CD001" not in r.text:
        raise StockPriceError(f"KRX 로그인 실패: {r.text[:120]}")
    return s


def _int(x):
    try:
        return int(str(x).replace(",", ""))
    except (TypeError, ValueError):
        return None


def short_on(d, s):
    """그날 전 종목 공매도 거래량 {종목코드: 주수} (유가증권+코스닥)."""
    out = {}
    for mkt in ("STK", "KSQ"):
        for attempt in range(3):
            try:
                r = s.post(KRX_JSON, timeout=40, data={"bld": "dbms/MDC/STAT/srt/MDCSTAT30101", "locale": "ko_KR", "searchType": "1",
                                                        "mktId": mkt, "trdDd": d.strftime("%Y%m%d"), "inqCond": "STMFRTSCIFDRFS",
                                                        "share": "1", "money": "1", "csvxls_isNo": "false"})
                rows = r.json().get("OutBlock_1") or []
                break
            except (requests.RequestException, ValueError):
                time.sleep(5 * (attempt + 1))
        else:
            raise StockPriceError("KRX 공매도 조회 실패")
        for it in rows:
            v = _int(it.get("CVSRTSELL_TRDVOL"))
            if it.get("ISU_CD") and v is not None:
                out[it["ISU_CD"]] = v
        time.sleep(0.5)
    return out


def lnb_on(d, s, key):
    """그날 전 종목 대차잔고 {종목코드: 잔고주수}."""
    out, page, sample = {}, 1, None
    while page <= 8:
        for attempt in range(4):
            time.sleep(0.2 if attempt == 0 else 5 * attempt)
            try:
                r = s.get(LNB_URL, timeout=40, params=dict(serviceKey=key, resultType="json", numOfRows=1000, pageNo=page,
                                                            basDt=d.strftime("%Y%m%d")))
            except requests.RequestException:
                continue
            if r.status_code in (502, 503, 504) or "SERVICETIMEOUT" in r.text[:400]:
                continue
            break
        else:
            raise StockPriceError("대차 API 일시 오류")
        if r.status_code >= 400 or r.text.lstrip().startswith("<"):
            raise StockPriceError(f"대차 HTTP {r.status_code}: {redact(r.text[:200])}")
        body = r.json().get("response", {}).get("body", {})
        items = (body.get("items") or {})
        items = items.get("item", []) if isinstance(items, dict) else (items or [])
        if isinstance(items, dict):
            items = [items]
        sample = sample or (items[0] if items else None)
        for it in items:
            code = str(it.get("srtnCd") or it.get("stckItmsCd") or it.get("itmsCd") or "").strip()
            if not code and it.get("isinCd"):
                code = it["isinCd"][3:9]
            if code.startswith("A") and len(code) == 7:
                code = code[1:]
            v = _int(it.get("balnStckCnt"))
            if code and v is not None:
                out[code] = v
        total = int(body.get("totalCount") or 0)
        if not items or page * 1000 >= total:
            break
        page += 1
    return out, sample


def update_short_lnb(today, max_fetch=30, diag=None):
    """px에 있는 거래일 중 아직 없는 날의 공매도 거래량(ref/short)·대차잔고(ref/lnb)를 최근부터 채운다."""
    import glob
    import os
    from .store import dump_json, load_json
    days = sorted({x for f in glob.glob(os.path.join(config.DATA_DIR, "ref", "px", "2*.json"))
                   for x in load_json(os.path.relpath(f, config.DATA_DIR), {}).get("dates", [])}, reverse=True)
    diag = diag if diag is not None else {}
    for kind in ("short", "lnb"):
        months, n, fails = {}, 0, 0
        try:
            sess = krx_session() if kind == "short" else requests.Session()
        except Exception as e:  # noqa: BLE001
            diag[kind + "_error"] = redact(e)[:200]
            continue
        key = _key()
        for iso in days:
            if n >= max_fetch or fails >= 8:
                break
            m = iso[:7]
            if m not in months:
                months[m] = load_json(f"ref/{kind}/{m}.json", {"dates": [], "v": {}})
            f = months[m]
            if iso in f["dates"]:
                continue
            from datetime import date as _date
            d = _date.fromisoformat(iso)
            try:
                if kind == "short":
                    vals = short_on(d, sess)
                else:
                    vals, sample = lnb_on(d, sess, key)
                    diag.setdefault("lnb_sample", sample)
            except Exception as e:  # noqa: BLE001
                fails += 1
                diag[kind + "_error"] = redact(e)[:200]
                continue
            n += 1
            if not vals:
                continue
            f["dates"].append(iso)
            f["dates"].sort()
            i, L = f["dates"].index(iso), len(f["dates"])
            for code in set(f["v"]) | set(vals):
                f["v"].setdefault(code, [None] * (L - 1)).insert(i, vals.get(code))
            f["dirty"] = True
        for m, f in months.items():
            if f.pop("dirty", None):
                dump_json(f"ref/{kind}/{m}.json", f, indent=None)
        diag[kind + "_fetched"] = n
    return diag
