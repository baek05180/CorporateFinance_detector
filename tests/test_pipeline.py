"""가짜 API 응답으로 수집 → 저장 → 페이지 생성까지 돌려보는 점검 스크립트.

  python -m tests.test_pipeline
실제 API는 호출하지 않는다. 결과물은 임시 폴더에 만들어진다.
"""
import json
import os
import shutil
import tempfile
from datetime import timedelta
from urllib.parse import urlparse

from collector import config

TMP = tempfile.mkdtemp(prefix="ecm-test-")
config.DATA_DIR = os.path.join(TMP, "data")
config.PUBLIC_DIR = os.path.join(TMP, "public")
config.DART_SLEEP_SEC = 0
os.environ["DART_API_KEY"] = "x" * 40
os.environ["DATA_GO_KR_KEY"] = "test"
config.DART_API_KEY = "x" * 40
config.DATA_GO_KR_KEY = "test"

from collector import run, store as store_mod  # noqa: E402
from collector.store import now_kst  # noqa: E402

T = now_kst().date()
D = lambda n: (T - timedelta(days=n)).strftime("%Y%m%d")  # noqa: E731
ISO = lambda n: (T - timedelta(days=n)).isoformat()  # noqa: E731


def F(corp, code, cls, nm, rno, stock=""):
    return dict(corp_name=corp, corp_code=code, corp_cls=cls, report_nm=nm, rcept_no=rno,
                rcept_dt=rno[:8], stock_code=stock, flr_nm=corp, rm="")


LIST_B = [
    F("금호건설", "00000001", "Y", "주요사항보고서(전환사채권발행결정)", D(1) + "000370", "002990"),
    F("비나텍", "00000002", "K", "주요사항보고서(신주인수권부사채권발행결정)", D(1) + "000332", "126340"),
    F("엔비홀딩스", "00000003", "E", "주요사항보고서(교환사채권발행결정)", D(2) + "000100"),
    F("CSA 코스믹", "00000004", "K", "주요사항보고서(유상증자결정)", D(5) + "000200", "083660"),
    F("CSA 코스믹", "00000004", "K", "[기재정정]주요사항보고서(유상증자결정)", D(1) + "000500", "083660"),
    F("씨어스", "00000011", "K", "주요사항보고서(전환사채권발행결정)", D(0) + "000010", "458870"),
    F("아무개", "00000009", "K", "주요사항보고서(자기주식취득결정)", D(1) + "000900"),
    F("유무상테스트", "00000010", "K", "주요사항보고서(유무상증자결정)", D(1) + "000901"),
]
LIST_C = [
    F("한화솔루션", "00000005", "Y", "증권신고서(채무증권)", D(3) + "000111", "009830"),
    F("KB캐피탈", "00000006", "E", "일괄신고추가서류(채무증권)", D(2) + "000222"),
    F("에스케이렌터카", "00000007", "E", "증권신고서(채무증권)", D(2) + "000333"),
]

DETAIL = {
    ("cvbdIsDecsn", "00000001"): [dict(
        rcept_no=D(1) + "000370", corp_cls="Y", corp_code="00000001", corp_name="금호건설",
        bd_tm="304", bd_knd="무기명식 이권부 무보증 사모 전환사채", bd_fta="50,000,000,000",
        bd_intr_ex="0.0", bd_intr_sf="3.0", bd_mtd="2029년 10월 23일", bdis_mthn="사모",
        cv_prc="4,325", cvisstk_cnt="11,560,693", cvisstk_tisstk_vs="31.25",
        cvrqpd_bgd="2027년 10월 23일", cvrqpd_edd="2029년 09월 23일",
        act_mktprcfl_cvprc_lwtrsprc="3,028", fdpp_op="30,000,000,000", fdpp_dtrp="20,000,000,000",
        bddd="2026년 10월 20일", pymd="2026년 10월 23일", rpmcmp="-")],
    ("bdwtIsDecsn", "00000002"): [dict(
        rcept_no=D(1) + "000332", corp_cls="K", corp_code="00000002", corp_name="비나텍",
        bd_tm="5", bd_fta="29,000,000,000", bd_intr_ex="1.0", bd_intr_sf="4.0", bd_mtd="2031-10-27",
        bdis_mthn="공모", ex_prc="52,300", nstk_isstk_cnt="554,493", nstk_isstk_tisstk_vs="9.12",
        expd_bgd="2027-10-27", expd_edd="2031-09-27", bdwt_div_atn="분리", bddd="2026-10-20",
        fdpp_fclt="29,000,000,000", rpmcmp="키움증권")],
    ("exbdIsDecsn", "00000003"): [dict(
        rcept_no=D(2) + "000100", corp_cls="E", corp_code="00000003", corp_name="엔비홀딩스",
        bd_tm="1", bd_fta="20,300,000,000", bd_intr_ex="0", bd_intr_sf="2", bdis_mthn="사모",
        ex_prc="18,000", extg="(주)엔비 보통주", extg_stkcnt="1,127,777", extg_tisstk_vs="4.5",
        bddd="2026.10.14", fdpp_ocsa="20,300,000,000")],
    ("cvbdIsDecsn", "00000011"): [dict(
        rcept_no=D(0) + "000010", corp_cls="K", corp_code="00000011", corp_name="씨어스",
        bd_tm="2", bd_fta="110,000,000,000", bdis_mthn="사모", cv_prc="61,200", bddd=ISO(0))],
    ("piicDecsn", "00000004"): [
        dict(rcept_no=D(5) + "000200", corp_cls="K", corp_code="00000004", corp_name="CSA 코스믹",
             nstk_ostk_cnt="3,000,000", nstk_estk_cnt="-", fv_ps="500", bfic_tisstk_ostk="26,500,000",
             bfic_tisstk_estk="-", fdpp_op="1,800,000,000", ic_mthn="일반공모"),
        dict(rcept_no=D(1) + "000500", corp_cls="K", corp_code="00000004", corp_name="CSA 코스믹",
             nstk_ostk_cnt="3,000,000", nstk_estk_cnt="-", fv_ps="500", bfic_tisstk_ostk="26,500,000",
             bfic_tisstk_estk="-", fdpp_op="2,000,000,000", ic_mthn="일반공모"),
    ],
}

BDRS = {
    "00000005": {"group": [
        {"title": "일반사항", "list": [
            dict(rcept_no=D(3) + "000111", corp_cls="Y", corp_code="00000005", corp_name="한화솔루션",
                 tm="321-1", bdnmn="무기명식 이권부 무보증사채", slmth="일반공모", fta="100,000,000,000",
                 slta="100,000,000,000", intr="수요예측 결과에 따라 결정", rpd="2028년 10월 30일",
                 cdrt_int="AA-(한국신용평가)", sbd="2026년 10월 28일", pymd="2026년 10월 30일"),
            dict(rcept_no=D(3) + "000111", corp_cls="Y", corp_code="00000005", corp_name="한화솔루션",
                 tm="321-2", bdnmn="무기명식 이권부 무보증사채", slmth="일반공모", slta="50,000,000,000",
                 intr="3.512", rpd="2029-10-30", cdrt_int="AA-", sbd="2026-10-28", pymd="2026-10-30"),
        ]},
        {"title": "인수인정보", "list": [
            dict(rcept_no=D(3) + "000111", tm="321-1", actsen="대표", actnmn="NH투자증권", udtamt="50,000,000,000"),
            dict(rcept_no=D(3) + "000111", tm="321-1", actsen="대표", actnmn="KB증권", udtamt="30,000,000,000"),
            dict(rcept_no=D(3) + "000111", tm="321-1", actsen="인수", actnmn="유진투자증권", udtamt="20,000,000,000"),
            dict(rcept_no=D(3) + "000111", tm="321-2", actsen="대표", actnmn="NH투자증권", udtamt="50,000,000,000"),
        ]},
    ]},
    "00000006": {"group": [{"title": "일반사항", "list": [
        dict(rcept_no=D(2) + "000222", corp_code="00000006", corp_name="KB캐피탈", tm="500",
             bdnmn="무보증사채", slta="100,000,000,000")]}]},
    "00000007": {"group": [{"title": "일반사항", "list": [
        dict(rcept_no=D(2) + "000333", corp_code="00000007", corp_name="에스케이렌터카", tm="40-1",
             bdnmn="무기명식 이권부 무보증 전환사채", slta="30,000,000,000")]}]},
}

COMPANY = {"00000005": "20302", "00000006": "64992", "00000007": "76110"}


def bond_item(isin, name, issuer, method, kind, issue_ago, years, amt, pclr="주식관련해당사항없음", sic="제조업"):
    return dict(basDt=D(1), isinCd=isin, isinCdNm=name, bondIsurNm=issuer, bondOffrMcdNm=method,
                scrsItmsKcdNm=kind, pclrBondKcdNm=pclr, sicNm=sic, bondIssuDt=D(issue_ago),
                bondExprDt=(T - timedelta(days=issue_ago) + timedelta(days=365 * years)).strftime("%Y%m%d"),
                bondIssuAmt=str(amt), bondSrfcInrt="5.5", grnDcdNm="무보증", issuDptyNm="교보증권",
                cptUsgeDcdNm="운영자금")


BOND_ITEMS = [
    bond_item("KR6336261G01", "두산퓨얼셀11", "두산퓨얼셀", "사모", "일반회사채", 3, 2, 10000000000),
    bond_item("KR6012631G02", "쌍용건설120", "쌍용건설", "사모", "일반회사채", 2, 2, 12000000000),
    bond_item("KR6000001G03", "현대차500-1", "현대자동차", "공모", "일반회사채", 2, 3, 300000000000),
    bond_item("KR6000002G04", "OO캐피탈1234", "OO캐피탈", "사모", "금융채", 2, 2, 20000000000, sic="여신금융업"),
    bond_item("KR6000003G05", "단기물", "어떤회사", "사모", "일반회사채", 2, 0.5, 5000000000),
    bond_item("KR6000004G06", "OO전환1", "OO바이오", "사모", "일반회사채", 2, 3, 5000000000, pclr="CB"),
    bond_item("KR6000005G07", "옛날사모채", "오래된회사", "사모", "일반회사채", 120, 3, 5000000000),
]


class Resp:
    def __init__(self, obj):
        self._obj = obj
        self.text = json.dumps(obj, ensure_ascii=False)
        self.status_code = 200

    def raise_for_status(self):
        pass

    def json(self):
        return self._obj


class FakeSession:
    def __init__(self):
        self.calls = []

    def get(self, url, params=None, timeout=None):
        p = params or {}
        ep = urlparse(url).path.rsplit("/", 1)[-1].replace(".json", "")
        self.calls.append((ep, {k: v for k, v in p.items() if k not in ("crtfc_key", "serviceKey")}))
        ok = lambda **kw: Resp({"status": "000", "message": "정상", **kw})  # noqa: E731
        if ep == "list":
            src = LIST_B if p.get("pblntf_ty") == "B" else LIST_C
            rows = [f for f in src if p["bgn_de"] <= f["rcept_dt"] <= p["end_de"]]
            return ok(list=rows, total_page=1, total_count=len(rows))
        if ep in ("cvbdIsDecsn", "bdwtIsDecsn", "exbdIsDecsn", "piicDecsn"):
            rows = DETAIL.get((ep, p["corp_code"]), [])
            return ok(list=rows) if rows else Resp({"status": "013", "message": "없음"})
        if ep == "bdRs":
            return ok(**BDRS.get(p["corp_code"], {"group": []}))
        if ep == "company":
            return ok(induty_code=COMPANY.get(p["corp_code"], ""), stock_code="")
        if ep == "getBondBasiInfo_V2":
            items = BOND_ITEMS
            return Resp({"response": {"header": {"resultCode": "00"}, "body": {
                "totalCount": len(items), "items": {"item": items}}}})
        raise AssertionError(f"예상 못한 호출 {url}")


def main():
    fake = FakeSession()
    import collector.dart as dmod
    import collector.datagokr as gmod
    dmod.requests.Session = lambda: fake
    gmod.requests.Session = lambda: fake

    # 1) 실시간 감시 → 2) 같은 감시 다시 (중복 호출 없어야) → 3) 정기 갱신
    run.main(["--mode", "watch"])
    db0 = json.load(open(os.path.join(config.DATA_DIR, "records.json"), encoding="utf-8"))
    watch_caught = [r for r in db0["records"].values() if r["captured_by"] == "watch" and not r["confirmed"]]
    calls_after_first = len(fake.calls)
    run.main(["--mode", "watch"])
    second_run_calls = len(fake.calls) - calls_after_first
    run.main(["--mode", "daily", "--days", "7"])

    db = json.load(open(os.path.join(config.DATA_DIR, "records.json"), encoding="utf-8"))
    recs = db["records"]
    by_type = {}
    for r in recs.values():
        by_type.setdefault((r["type"], r.get("offering")), []).append(r["corp_name"])

    checks = []

    def check(name, cond):
        checks.append((name, bool(cond)))

    check("CB 금호건설 500억 사모", any(r["type"] == "CB" and r["amount"] == 50_000_000_000 and r["offering"] == "사모"
                                     and r["price"] == 4325 and r["refix_floor"] == 3028 for r in recs.values()))
    check("BW 비나텍 공모", any(r["type"] == "BW" and r["offering"] == "공모" for r in recs.values()))
    check("EB 엔비홀딩스", any(r["type"] == "EB" for r in recs.values()))
    rights = [r for r in recs.values() if r["type"] == "RIGHTS"]
    check("유상증자 정정공시는 원공시에 합쳐져 1건", len(rights) == 1)
    check("유상증자 금액 = 정정 후 20억", rights and rights[0]["amount"] == 2_000_000_000)
    check("유상증자 증자비율 11.32%", rights and abs(rights[0]["dilution"] - 11.32) < 0.01)
    pub = [r for r in recs.values() if r["type"] == "BOND" and r["offering"] == "공모"]
    check("공모채 한화솔루션 2개 트랜치", sorted(r["series_raw"] for r in pub) == ["321-1", "321-2"])
    check("공모채 대표주관 묶음", any(r["lead"] == "KB증권, NH투자증권" for r in pub))
    check("금융사(KB캐피탈) 공모채 제외", not any(r["corp_name"] == "KB캐피탈" for r in recs.values()))
    check("전환사채 명칭의 공모 채무증권 제외", not any(r["corp_name"] == "에스케이렌터카" for r in recs.values()))
    pvt = sorted(r["corp_name"] for r in recs.values() if r["type"] == "BOND" and r["offering"] == "사모")
    check("사모채 = 두산퓨얼셀·쌍용건설만", pvt == ["두산퓨얼셀", "쌍용건설"])
    check("자기주식·유무상증자 공시 무시", not any(r["corp_name"] in ("아무개", "유무상테스트") for r in recs.values()))
    check("두 번째 감시 실행은 목록 조회 2회뿐", second_run_calls == 2)
    check("실시간 감시가 오늘 CB(씨어스)를 미확정 상태로 잡음", [r["corp_name"] for r in watch_caught] == ["씨어스"])
    check("정기 갱신 후 확정 표시", all(r.get("confirmed") for r in recs.values()))

    from web import build
    build.main()
    html = open(os.path.join(config.PUBLIC_DIR, "index.html"), encoding="utf-8").read()
    check("페이지에 데이터 삽입", "금호건설" in html and "__DATA__" not in html)

    for name, ok in checks:
        print(("PASS " if ok else "FAIL ") + name)
    print(f"\n결과물: {TMP}")
    print("diag:", json.dumps(json.load(open(os.path.join(config.DATA_DIR, "diag.json"), encoding="utf-8")), ensure_ascii=False))
    return 0 if all(ok for _, ok in checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
