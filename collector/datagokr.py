"""공공데이터포털 금융위원회_채권기본정보 클라이언트 (사모 회사채용)."""
import re
import time
import logging
from urllib.parse import unquote

import requests

from . import config

log = logging.getLogger(__name__)


class DataGoKrError(Exception):
    pass


def redact(text):
    """오류 메시지에 인증키가 섞여 공개 저장소에 남지 않도록 지운다."""
    t = re.sub(r"((?:serviceKey|crtfc_key)=)[^&\s'\"]+", r"\1***", str(text))
    for k in (config.DATA_GO_KR_KEY, config.DART_API_KEY):   # 인코딩 여부와 관계없이 키 값 자체도 지운다
        if k and len(k) > 8:
            from urllib.parse import quote, unquote
            for v in {k, unquote(k), quote(unquote(k), safe="")}:
                t = t.replace(v, "***")
    return t


class BondInfoClient:
    def __init__(self, service_key=None, session=None):
        key = (service_key or config.DATA_GO_KR_KEY or "").strip()
        if not key:
            raise DataGoKrError("DATA_GO_KR_KEY 환경변수가 비어 있습니다.")
        # Encoding 키(%2B 등)를 넣어도 동작하도록 한 번 풀어준다 (requests가 다시 인코딩함)
        self.key = unquote(key) if "%" in key else key
        self.session = session or requests.Session()
        self.calls = 0

    def _page(self, bas_dt, page):
        params = dict(serviceKey=self.key, resultType="json", numOfRows=config.BOND_INFO_ROWS,
                      pageNo=page, basDt=bas_dt)
        last_err = None
        for attempt in range(3):
            try:
                time.sleep(0.2)
                r = self.session.get(config.BOND_INFO_URL, params=params, timeout=30)
                self.calls += 1
                if r.text.lstrip().startswith("<"):
                    # 인증 오류 등은 json 요청이어도 XML로 온다
                    raise DataGoKrError(f"XML 오류 응답(HTTP {r.status_code}): {redact(r.text[:300])}")
                if r.status_code >= 400:
                    raise DataGoKrError(f"HTTP {r.status_code}: {redact(r.text[:300])}")
                data = r.json()
                break
            except DataGoKrError:
                raise
            except (requests.RequestException, ValueError) as e:
                last_err = e
                time.sleep(2 * (attempt + 1))
        else:
            raise DataGoKrError(f"채권기본정보 호출 실패: {redact(last_err)}")

        resp = data.get("response", data)
        header = resp.get("header", {})
        if header.get("resultCode") not in ("00", "0", None):
            raise DataGoKrError(f"resultCode={header.get('resultCode')} {header.get('resultMsg')}")
        body = resp.get("body", {})
        items = (body.get("items") or {})
        items = items.get("item", []) if isinstance(items, dict) else items
        if isinstance(items, dict):
            items = [items]
        return items or [], int(body.get("totalCount") or 0)

    def bonds_for_basdt(self, bas_dt):
        """기준일(YYYYMMDD) 하루치 채권기본정보 전체."""
        out, page = [], 1
        while page <= config.BOND_INFO_MAX_PAGES:
            items, total = self._page(bas_dt, page)
            out.extend(items)
            if not items or len(out) >= total:
                break
            page += 1
        else:
            log.warning("기준일 %s: 최대 페이지(%s) 도달, 일부만 수집", bas_dt, config.BOND_INFO_MAX_PAGES)
        return out
