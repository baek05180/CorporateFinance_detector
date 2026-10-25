"""OpenDART API 클라이언트 (requests 기반, 재시도·한도 초과 처리 포함)."""
import time
import logging

import requests

from . import config

log = logging.getLogger(__name__)


class DartQuotaExceeded(Exception):
    """status 020: 일일 호출 한도 초과."""


class DartError(Exception):
    pass


class DartClient:
    def __init__(self, api_key=None, session=None):
        self.api_key = api_key or config.DART_API_KEY
        if not self.api_key:
            raise DartError("DART_API_KEY 환경변수가 비어 있습니다.")
        self.session = session or requests.Session()
        self.calls = 0

    def _get(self, endpoint, **params):
        params["crtfc_key"] = self.api_key
        url = f"{config.DART_BASE}/{endpoint}.json"
        last_err = None
        for attempt in range(3):
            try:
                time.sleep(config.DART_SLEEP_SEC)
                r = self.session.get(url, params=params, timeout=20)
                self.calls += 1
                r.raise_for_status()
                data = r.json()
                break
            except (requests.RequestException, ValueError) as e:
                last_err = e
                time.sleep(2 * (attempt + 1))
        else:
            raise DartError(f"{endpoint} 호출 실패: {last_err}")

        status = data.get("status")
        if status == "000":
            return data
        if status == "013":  # 조회된 데이터 없음
            return {**data, "list": [], "group": []}
        if status == "020":
            raise DartQuotaExceeded(data.get("message"))
        raise DartError(f"{endpoint} status={status} message={data.get('message')}")

    # ── 공시 목록 ──
    def list_filings(self, bgn_de, end_de, pblntf_ty, pblntf_detail_ty=None):
        """기간 내 공시 목록 전체(페이지 끝까지). 날짜는 YYYYMMDD."""
        out, page = [], 1
        while True:
            params = dict(bgn_de=bgn_de, end_de=end_de, pblntf_ty=pblntf_ty,
                          page_no=page, page_count=100)
            if pblntf_detail_ty:
                params["pblntf_detail_ty"] = pblntf_detail_ty
            data = self._get("list", **params)
            out.extend(data.get("list", []))
            total_page = int(data.get("total_page") or 1)
            if page >= total_page:
                return out
            page += 1

    # ── 주요사항보고서 상세 (CB/BW/EB/유상증자) ──
    def major_detail(self, endpoint, corp_code, bgn_de, end_de):
        return self._get(endpoint, corp_code=corp_code, bgn_de=bgn_de, end_de=end_de).get("list", [])

    # ── 증권신고서(채무증권) ──
    def bond_registration(self, corp_code, bgn_de, end_de):
        """그룹 구조 그대로 반환: {"일반사항": [...], "인수인정보": [...], ...}"""
        data = self._get("bdRs", corp_code=corp_code, bgn_de=bgn_de, end_de=end_de)
        groups = {}
        for g in data.get("group", []) or []:
            groups[g.get("title", "")] = g.get("list", []) or []
        return groups

    # ── 기업개황 (업종코드 확인용) ──
    def company(self, corp_code):
        return self._get("company", corp_code=corp_code)
