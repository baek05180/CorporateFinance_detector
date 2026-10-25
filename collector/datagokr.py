"""공공데이터포털 금융위원회_채권기본정보 클라이언트 (사모 회사채용)."""
import time
import logging

import requests

from . import config

log = logging.getLogger(__name__)


class DataGoKrError(Exception):
    pass


class BondInfoClient:
    def __init__(self, service_key=None, session=None):
        self.key = service_key or config.DATA_GO_KR_KEY
        if not self.key:
            raise DataGoKrError("DATA_GO_KR_KEY 환경변수가 비어 있습니다.")
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
                    raise DataGoKrError(f"XML 오류 응답: {r.text[:300]}")
                r.raise_for_status()
                data = r.json()
                break
            except DataGoKrError:
                raise
            except (requests.RequestException, ValueError) as e:
                last_err = e
                time.sleep(2 * (attempt + 1))
        else:
            raise DataGoKrError(f"채권기본정보 호출 실패: {last_err}")

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
