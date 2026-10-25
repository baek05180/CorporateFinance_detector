"""일회성 점검: 실제 API 응답 샘플을 data/probe/ 에 저장한다 (인증키는 저장하지 않음).

  python -m tools.probe
"""
import io
import json
import os
import re
import zipfile
from collections import Counter
from datetime import date, timedelta

import requests

from collector import config
from collector.dart import DartClient
from collector.datagokr import redact

OUT = os.path.join(config.DATA_DIR, "probe")
os.makedirs(OUT, exist_ok=True)


def save(name, obj):
    with open(os.path.join(OUT, name), "w", encoding="utf-8") as f:
        if isinstance(obj, str):
            f.write(obj)
        else:
            json.dump(obj, f, ensure_ascii=False, indent=1)


def safe(fn, *a, **kw):
    try:
        return fn(*a, **kw)
    except Exception as e:  # noqa: BLE001
        return {"error": redact(str(e))[:500]}


dart = DartClient()
today = date.today()
summary = {}

# 1) 채무증권 증권신고 목록: 보고서명별 건수, 상세 API 커버 여부
c002 = dart.list_filings((today - timedelta(days=60)).strftime("%Y%m%d"), today.strftime("%Y%m%d"), "C", "C002")
c002 += dart.list_filings((today - timedelta(days=120)).strftime("%Y%m%d"), (today - timedelta(days=61)).strftime("%Y%m%d"), "C", "C002")
summary["c002_report_names"] = Counter(re.sub(r"\[.*?\]", "", f["report_nm"]).strip() for f in c002).most_common(30)
summary["c002_prefixes"] = Counter((re.match(r"\[(.*?)\]", f["report_nm"]) or [None, ""])[1] for f in c002).most_common()
summary["c002_corp_cls"] = Counter(f.get("corp_cls") for f in c002).most_common()
shelf = [f for f in c002 if "일괄신고추가" in f["report_nm"]][:6]
plain = [f for f in c002 if "일괄신고추가" not in f["report_nm"] and "증권신고서" in f["report_nm"]][:3]
cover = []
for f in shelf + plain:
    rno = f["rcept_no"]
    g = safe(dart.bond_registration, f["corp_code"], rno[:8], rno[:8])
    hit = [x for x in (g.get("일반사항", []) if isinstance(g, dict) else []) if x.get("rcept_no") == rno]
    cover.append({"report_nm": f["report_nm"], "corp": f["corp_name"], "rcept_no": rno,
                  "bdRs_rows_same_rcept": len(hit), "bdRs_groups": list(g.keys()) if isinstance(g, dict) else g})
summary["bdRs_coverage"] = cover
save("c002_sample.json", c002[:80])

# 2) 공시 원문 (공모채 증권신고서) → 수요예측 문구 찾기
def document(rno):
    r = dart.session.get(f"{config.DART_BASE}/document.xml", params={"crtfc_key": dart.api_key, "rcept_no": rno}, timeout=60)
    dart.calls += 1
    z = zipfile.ZipFile(io.BytesIO(r.content))
    files = {}
    for n in z.namelist():
        raw = z.read(n)
        for enc in ("utf-8", "euc-kr", "cp949"):
            try:
                files[n] = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
    return files


for i, f in enumerate((shelf + plain)[:3]):
    docs = safe(document, f["rcept_no"])
    if isinstance(docs, dict) and "error" in docs:
        summary.setdefault("bond_doc_errors", []).append(docs)
        continue
    info = {"report_nm": f["report_nm"], "files": {k: len(v) for k, v in docs.items()}}
    main = max(docs.values(), key=len)
    hits = [m.start() for m in re.finditer("수요예측", main)][:8]
    info["demand_snippets"] = [re.sub(r"<[^>]+>", " ", main[max(0, h - 200):h + 300]) for h in hits]
    info["head"] = main[:3000]
    save(f"bond_doc_{i}.json", info)

# 3) 사업보고서 원문 구조 (대형 1, 코스닥 1)
for corp, name in [("00126380", "samsung"), ("01275665", "remed")]:
    lst = safe(dart.list_filings, "20260101", today.strftime("%Y%m%d"), "A", "A001") if False else None
    data = safe(dart._get, "list", corp_code=corp, bgn_de="20250101", end_de=today.strftime("%Y%m%d"),
                pblntf_ty="A", page_count=20)
    reports = [x for x in data.get("list", [])] if isinstance(data, dict) else []
    summary[f"{name}_periodic"] = [(x["report_nm"], x["rcept_no"]) for x in reports]
    annual = next((x for x in reports if "사업보고서" in x["report_nm"]), None)
    if not annual:
        continue
    docs = safe(document, annual["rcept_no"])
    if isinstance(docs, dict) and "error" in docs:
        summary[f"{name}_doc_error"] = docs
        continue
    main = max(docs.values(), key=len)
    titles = re.findall(r"<TITLE[^>]*>(.*?)</TITLE>", main, flags=re.S | re.I)
    summary[f"{name}_doc"] = {"files": {k: len(v) for k, v in docs.items()}, "titles": titles[:80]}
    i = main.find("사업의 개요")
    save(f"annual_{name}_excerpt.xml", main[max(0, i - 3000): i + 60000])

# 4) 재무 API 샘플
summary["multi"] = safe(dart._get, "fnlttMultiAcnt", corp_code="00126380,01275665,00106313",
                        bsns_year="2025", reprt_code="11011")
if isinstance(summary["multi"], dict) and "list" in summary["multi"]:
    summary["multi"] = summary["multi"]["list"][:40]
full = safe(dart._get, "fnlttSinglAcntAll", corp_code="01275665", bsns_year="2025", reprt_code="11011", fs_div="CFS")
if isinstance(full, dict) and full.get("status") != "000":
    full = safe(dart._get, "fnlttSinglAcntAll", corp_code="01275665", bsns_year="2025", reprt_code="11011", fs_div="OFS")
save("full_fs_remed.json", full.get("list", full) if isinstance(full, dict) else full)
summary["hyslr"] = safe(dart._get, "hyslrSttus", corp_code="01275665", bsns_year="2025", reprt_code="11011")
summary["tesstk"] = safe(dart._get, "tesstkAcqsDspsSttus", corp_code="00126380", bsns_year="2025", reprt_code="11011")
summary["company"] = safe(dart.company, "01275665")

# 5) 고유번호 전체 (크기만)
r = dart.session.get(f"{config.DART_BASE}/corpCode.xml", params={"crtfc_key": dart.api_key}, timeout=120)
z = zipfile.ZipFile(io.BytesIO(r.content))
x = z.read(z.namelist()[0]).decode("utf-8")
summary["corpcode"] = {"bytes": len(x), "n": x.count("<list>"), "head": x[:600]}

# 6) 주식시세 (공공데이터포털)
key = config.DATA_GO_KR_KEY
from urllib.parse import unquote  # noqa: E402
key = unquote(key) if "%" in key else key
sp = requests.get("https://apis.data.go.kr/1160100/service/GetStockSecuritiesInfoService/getStockPriceInfo",
                  params=dict(serviceKey=key, resultType="json", numOfRows=3, pageNo=1, likeSrtnCd="005930"), timeout=30)
sp2 = requests.get("https://apis.data.go.kr/1160100/GetStockSecuritiesInfoService_V2/getStockPriceInfo_V2",
                   params=dict(serviceKey=key, resultType="json", numOfRows=3, pageNo=1, likeSrtnCd="005930"), timeout=30)
summary["stock_v1"] = {"status": sp.status_code, "body": redact(sp.text[:800])}
summary["stock_v2"] = {"status": sp2.status_code, "body": redact(sp2.text[:800])}

# 7) KIND 상장사 목록
try:
    k = requests.post("https://kind.krx.co.kr/corpgeneral/corpList.do",
                      data={"method": "download", "searchType": "13", "orderMode": "3", "orderStat": "D"},
                      headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
    summary["kind"] = {"status": k.status_code, "bytes": len(k.content), "head": k.content[:400].decode("euc-kr", "ignore")}
except Exception as e:  # noqa: BLE001
    summary["kind"] = {"error": str(e)}

summary["dart_calls"] = dart.calls
save("summary.json", summary)
print("probe done, DART calls:", dart.calls)
