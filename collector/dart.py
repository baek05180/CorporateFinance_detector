"""OpenDART API 클라이언트 (requests 기반, 재시도·한도 초과 처리 포함)."""
import io
import threading
import time
import logging
import zipfile

import requests

from . import config

log = logging.getLogger(__name__)


class DartQuotaExceeded(Exception):
    """status 020: 일일 호출 한도 초과."""


class DartError(Exception):
    pass


class DartBudgetExceeded(Exception):
    """이번 실행에 정한 호출 예산을 다 씀 (다음 실행에서 이어서)."""


class DartClient:
    def __init__(self, api_key=None, session=None):
        self.api_key = api_key or config.DART_API_KEY
        if not self.api_key:
            raise DartError("DART_API_KEY 환경변수가 비어 있습니다.")
        self.session = session or requests.Session()
        if hasattr(self.session, "mount"):   # 병렬 다운로드용 연결 풀
            ad = requests.adapters.HTTPAdapter(pool_connections=12, pool_maxsize=12)
            self.session.mount("https://", ad)
        self.calls = 0
        self._lock = threading.Lock()
        self.budget = None   # None = 제한 없음
        self.deadline = None  # time.time() 기준 마감 시각

    def _spend(self):
        """호출 1회를 기록하고, 예산(횟수·시간)을 넘었으면 멈춘다. 여러 스레드에서 불려도 안전."""
        with self._lock:
            if self.budget is not None and self.calls >= self.budget:
                raise DartBudgetExceeded(f"호출 예산 {self.budget}회 소진")
            if self.deadline is not None and time.time() >= self.deadline:
                raise DartBudgetExceeded("시간 예산 소진")
            self.calls += 1

    def _get(self, endpoint, **params):
        params["crtfc_key"] = self.api_key
        url = f"{config.DART_BASE}/{endpoint}.json"
        last_err = None
        for attempt in range(3):
            try:
                self._spend()
                time.sleep(config.DART_SLEEP_SEC)
                r = self.session.get(url, params=params, timeout=20)
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

    # ── 정기보고서 목록 (사업·반기·분기) ──
    def periodic_reports(self, corp_code, bgn_de, end_de):
        data = self._get("list", corp_code=corp_code, bgn_de=bgn_de, end_de=end_de,
                         pblntf_ty="A", page_count=100, last_reprt_at="Y")
        return data.get("list", [])

    # ── 원문 파일 (zip 안의 XML들) ──
    def _download(self, endpoint, **params):
        params["crtfc_key"] = self.api_key
        last_err = None
        for attempt in range(3):
            try:
                self._spend()
                time.sleep(config.DART_SLEEP_SEC)
                r = self.session.get(f"{config.DART_BASE}/{endpoint}", params=params, timeout=120)
                r.raise_for_status()
                content = r.content
                break
            except requests.RequestException as e:
                last_err = e
                time.sleep(2 * (attempt + 1))
        else:
            raise DartError(f"{endpoint} 다운로드 실패: {last_err}")
        if content[:2] != b"PK":   # zip이 아니면 오류 응답(XML/JSON)
            text = content[:400].decode("utf-8", "ignore")
            if "020" in text and "제한" in text:
                raise DartQuotaExceeded(text)
            raise DartError(f"{endpoint} zip 아님: {text[:200]}")
        return zipfile.ZipFile(io.BytesIO(content))

    def document(self, rcept_no):
        """공시 원문 → {파일명: 텍스트}"""
        z = self._download("document.xml", rcept_no=rcept_no)
        out = {}
        for n in z.namelist():
            raw = z.read(n)
            for enc in ("utf-8", "cp949"):
                try:
                    out[n] = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                out[n] = raw.decode("utf-8", "ignore")
        return out

    def corp_codes(self):
        """고유번호 전체 → [{corp_code, corp_name, stock_code}] (상장사 아닌 곳은 stock_code 빈값)"""
        import re
        z = self._download("corpCode.xml")
        x = z.read(z.namelist()[0]).decode("utf-8")
        out = []
        for m in re.finditer(r"<list>(.*?)</list>", x, re.S):
            b = m.group(1)
            g = lambda t: (re.search(rf"<{t}>(.*?)</{t}>", b, re.S) or [None, ""])[1].strip()  # noqa: E731
            out.append({"corp_code": g("corp_code"), "corp_name": g("corp_name"), "stock_code": g("stock_code")})
        return out

    # ── 재무 ──
    def multi_accounts(self, corp_codes, year, reprt_code):
        data = self._get("fnlttMultiAcnt", corp_code=",".join(corp_codes), bsns_year=str(year), reprt_code=reprt_code)
        return data.get("list", [])

    def full_statements(self, corp_code, year, reprt_code):
        """연결 우선, 없으면 별도. (list, fs_div)"""
        for fs in ("CFS", "OFS"):
            data = self._get("fnlttSinglAcntAll", corp_code=corp_code, bsns_year=str(year),
                             reprt_code=reprt_code, fs_div=fs)
            if data.get("list"):
                return data["list"], fs
        return [], None

    def shareholders(self, corp_code, year, reprt_code):
        return self._get("hyslrSttus", corp_code=corp_code, bsns_year=str(year), reprt_code=reprt_code).get("list", [])

    def treasury(self, corp_code, year, reprt_code):
        return self._get("tesstkAcqsDspsSttus", corp_code=corp_code, bsns_year=str(year),
                         reprt_code=reprt_code).get("list", [])
