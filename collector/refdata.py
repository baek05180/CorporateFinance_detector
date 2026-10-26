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


def stock_prices(today, session=None):
    """가장 최근 거래일의 전 종목 종가·시가총액 → (기준일, {종목코드: {close, mcap, shares}})"""
    key = (config.DATA_GO_KR_KEY or "").strip()
    if not key:
        raise StockPriceError("DATA_GO_KR_KEY 없음")
    key = unquote(key) if "%" in key else key
    s = session or requests.Session()
    d = today
    for _ in range(10):
        d -= timedelta(days=1)
        if d.weekday() >= 5:
            continue
        out, page, total = {}, 1, None
        while True:
            time.sleep(0.2)
            r = s.get(config.STOCK_PRICE_URL, timeout=40, params=dict(
                serviceKey=key, resultType="json", numOfRows=1000, pageNo=page, basDt=d.strftime("%Y%m%d")))
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
                                 "shares": int(float(it.get("lstgStCnt")))}
                except (TypeError, ValueError):
                    continue
            if not items or len(out) >= total or page >= 6:
                break
            page += 1
        if out:
            return d.isoformat(), out
    return None, {}
