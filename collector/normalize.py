"""API 원본 응답 → 사이트에서 쓰는 공통 레코드로 변환."""
import re
from datetime import date

DART_VIEW = "https://dart.fss.or.kr/dsaf001/main.do?rcpNo="

PURPOSE_KEYS = {
    "fdpp_fclt": "시설",
    "fdpp_bsninh": "영업양수",
    "fdpp_op": "운영",
    "fdpp_dtrp": "채무상환",
    "fdpp_ocsa": "타법인증권",
    "fdpp_etc": "기타",
}


# ── 값 정리 도우미 ──
def num(v):
    """'1,000,000' / '-' / '' / None → int 또는 None"""
    if v is None:
        return None
    s = str(v).replace(",", "").strip()
    if s in ("", "-", "—", "해당사항없음", "해당없음"):
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def pct(v):
    if v is None:
        return None
    s = str(v).replace(",", "").replace("%", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def iso_date(v):
    """'2026년 10월 20일', '2026-10-20', '2026.10.20', '20261020' → '2026-10-20'"""
    if not v:
        return None
    s = str(v)
    m = re.search(r"(\d{4})\D{0,3}(\d{1,2})\D{0,3}(\d{1,2})", s)
    if not m:
        return None
    y, mo, d = (int(x) for x in m.groups())
    try:
        return date(y, mo, d).isoformat()
    except ValueError:
        return None


def series(v):
    """'제33회', '33', '33회 무기명식...' → '33' (숫자가 없으면 원문)"""
    if v is None:
        return ""
    m = re.search(r"\d+(?:-\d+)?", str(v))
    return m.group(0) if m else str(v).strip()


def offering_of(text):
    t = str(text or "")
    if "사모" in t:
        return "사모"
    if "공모" in t or "일괄신고" in t:
        return "공모"
    return t.strip() or None


def rcept_date(rcept_no):
    return iso_date(rcept_no[:8]) if rcept_no else None


def purposes(item):
    out = {}
    for k, label in PURPOSE_KEYS.items():
        v = num(item.get(k))
        if v:
            out[label] = v
    return out


def base(item, kind, filing=None):
    filing = filing or {}
    rno = item.get("rcept_no", "")
    return {
        "type": kind,
        "rcept_no": rno,
        "rcept_dt": rcept_date(rno),
        "url": DART_VIEW + rno if rno else None,
        "corp_code": item.get("corp_code"),
        "corp_name": item.get("corp_name") or filing.get("corp_name"),
        "corp_cls": item.get("corp_cls") or filing.get("corp_cls"),
        "stock_code": filing.get("stock_code") or None,
        "amended": "정정" in (filing.get("report_nm") or ""),
        "source": "DART",
    }


# ── 메자닌 (CB/BW/EB) ──
def mezzanine(item, kind, filing=None):
    r = base(item, kind, filing)
    r.update({
        "series": series(item.get("bd_tm")),
        "amount": num(item.get("bd_fta")),
        "offering": offering_of(item.get("bdis_mthn")),
        "method_raw": item.get("bdis_mthn"),
        "coupon": pct(item.get("bd_intr_ex")),
        "ytm": pct(item.get("bd_intr_sf")),
        "maturity": iso_date(item.get("bd_mtd")),
        "board_dt": iso_date(item.get("bddd")) or r["rcept_dt"],
        "pay_dt": iso_date(item.get("pymd")),
        "lead": item.get("rpmcmp") if item.get("rpmcmp") not in ("-", "", None) else None,
        "purposes": purposes(item),
        "refix_floor": num(item.get("act_mktprcfl_cvprc_lwtrsprc")),
    })
    if kind == "CB":
        r.update({
            "price": num(item.get("cv_prc")),
            "shares": num(item.get("cvisstk_cnt")),
            "dilution": pct(item.get("cvisstk_tisstk_vs")),
            "exercise_from": iso_date(item.get("cvrqpd_bgd")),
            "exercise_to": iso_date(item.get("cvrqpd_edd")),
        })
    elif kind == "BW":
        r.update({
            "price": num(item.get("ex_prc")),
            "shares": num(item.get("nstk_isstk_cnt")),
            "dilution": pct(item.get("nstk_isstk_tisstk_vs")),
            "exercise_from": iso_date(item.get("expd_bgd")),
            "exercise_to": iso_date(item.get("expd_edd")),
            "detachable": item.get("bdwt_div_atn"),
        })
    elif kind == "EB":
        r.update({
            "price": num(item.get("ex_prc")),
            "shares": num(item.get("extg_stkcnt")),
            "dilution": pct(item.get("extg_tisstk_vs")),
            "exercise_from": iso_date(item.get("exrqpd_bgd")),
            "exercise_to": iso_date(item.get("exrqpd_edd")),
            "exchange_target": item.get("extg"),
        })
    r["id"] = f"{kind}:{r['corp_code']}:{r['series'] or r['rcept_no']}"
    return r


# ── 유상증자 ──
def rights(item, filing=None):
    r = base(item, "RIGHTS", filing)
    o = num(item.get("nstk_ostk_cnt")) or 0
    e = num(item.get("nstk_estk_cnt")) or 0
    bo = num(item.get("bfic_tisstk_ostk")) or 0
    be = num(item.get("bfic_tisstk_estk")) or 0
    p = purposes(item)
    kinds = []
    if o:
        kinds.append("보통주")
    if e:
        kinds.append("종류주")
    r.update({
        "amount": sum(p.values()) or None,   # 유상증자 API엔 금액 필드가 없어 자금조달 목적 합계로 계산
        "new_common": o or None,
        "new_other": e or None,
        "stock_kind": "·".join(kinds) or None,
        "dilution": round((o + e) / (bo + be) * 100, 2) if (bo + be) and (o + e) else None,
        "method": item.get("ic_mthn"),
        "offering": "공모" if "공모" in str(item.get("ic_mthn") or "") else "사모·주주배정",
        "par_value": num(item.get("fv_ps")),
        "purposes": p,
        "board_dt": r["rcept_dt"],
    })
    # 유상증자는 회차가 없어 접수번호를 키로 쓴다. 정정공시는 store.merge_rights()에서 원공시에 합친다.
    r["id"] = f"RIGHTS:{r['corp_code']}:{r['rcept_no']}"
    return r


# ── 공모 회사채 (증권신고서 채무증권) ──
def public_bonds(groups, filing=None):
    general = groups.get("일반사항", [])
    underwriters = groups.get("인수인정보", [])
    uses = groups.get("자금의사용목적", [])
    out = []
    for item in general:
        tm = item.get("tm")
        uw = [u for u in underwriters if u.get("tm") == tm] or underwriters
        leads = sorted({u.get("actnmn") for u in uw if "대표" in str(u.get("actsen") or "") and u.get("actnmn")})
        others = sorted({u.get("actnmn") for u in uw if u.get("actnmn")} - set(leads))
        use = {u.get("se"): num(u.get("amt")) for u in uses if u.get("tm") == tm and num(u.get("amt"))}
        r = base(item, "BOND", filing)
        r.update({
            "series": series(tm),
            "series_raw": tm,
            "name": item.get("bdnmn"),
            "amount": num(item.get("slta")) or num(item.get("fta")),
            "offering": "공모",
            "method_raw": item.get("slmth"),
            "coupon_raw": item.get("intr"),
            "coupon": pct(item.get("intr")),
            "maturity": iso_date(item.get("rpd")),
            "rating": item.get("cdrt_int"),
            "sub_dt": iso_date(item.get("sbd")),
            "pay_dt": iso_date(item.get("pymd")),
            "guarantor": item.get("grt_int") if item.get("grt_int") not in ("-", "", None) else None,
            "lead": ", ".join(leads) or None,
            "underwriters": others,
            "uses": use,
        })
        r["id"] = f"BOND:{r['corp_code']}:{r['series_raw'] or r['series']}"
        out.append(r)
    return out


# ── 사모 회사채 (금융위원회_채권기본정보) ──
def private_bond(item):
    isin = item.get("isinCd")
    issue = iso_date(item.get("bondIssuDt"))
    method = str(item.get("bondOffrMcdNm") or "")
    return {
        "id": f"PBOND:{isin}",
        "type": "BOND",
        "offering": "사모" if "사모" in method else offering_of(method),
        "pcbo": "CBO" in method,                # P-CBO 기초자산 사모채 여부
        "source": "공공데이터포털",
        "isin": isin,
        "name": item.get("isinCdNm"),
        "corp_name": item.get("bondIsurNm"),
        "crno": item.get("crno"),
        "kind": item.get("scrsItmsKcdNm"),
        "industry": item.get("sicNm"),
        "amount": num(item.get("bondIssuAmt")),
        "coupon": pct(item.get("bondSrfcInrt")),
        "rate_type": item.get("irtChngDcdNm"),
        "issue_dt": issue,
        "maturity": iso_date(item.get("bondExprDt")),
        "guarantee": item.get("grnDcdNm"),
        "rank": item.get("bondRnknDcdNm"),
        "option": item.get("optnTcdNm"),
        "usage": item.get("cptUsgeDcdNm"),
        "lead": item.get("issuDptyNm") or None,
        "underwriter": item.get("bondUndtInstNm") or None,
        "rcept_dt": issue,
        "url": None,
    }
