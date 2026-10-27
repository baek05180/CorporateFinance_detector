"""수집기 설정값. 숫자나 키워드를 바꾸고 싶으면 이 파일만 고치면 된다."""
import os
from datetime import timedelta, timezone

KST = timezone(timedelta(hours=9))

# ── API 키 (GitHub Secrets → 환경변수로 들어온다. 코드에 직접 쓰지 말 것) ──
DART_API_KEY = os.environ.get("DART_API_KEY", "")
DATA_GO_KR_KEY = os.environ.get("DATA_GO_KR_KEY", "")  # 공공데이터포털 '일반 인증키(Decoding)'

# ── 경로 ──
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "data")
PUBLIC_DIR = os.path.join(ROOT, "public")

# ── OpenDART ──
DART_BASE = "https://opendart.fss.or.kr/api"
DART_SLEEP_SEC = 0.15          # 호출 간 간격 (분당 호출 제한 회피)
DETAIL_LOOKBACK_DAYS = 120     # 상세 API는 '최초접수일' 기준이라 정정공시를 잡으려면 넉넉히

# 공시 제목(공백 제거 후)에 들어 있으면 수집하는 키워드 → 유형
MAJOR_REPORT_KINDS = {
    "전환사채권발행결정": "CB",
    "신주인수권부사채권발행결정": "BW",
    "교환사채권발행결정": "EB",
    "유상증자결정": "RIGHTS",
}
# 유무상증자결정은 별도 API라 MVP에서는 제외
MAJOR_REPORT_EXCLUDE = ["유무상증자결정"]

DETAIL_ENDPOINT = {
    "CB": "cvbdIsDecsn",
    "BW": "bdwtIsDecsn",
    "EB": "exbdIsDecsn",
    "RIGHTS": "piicDecsn",
}

# 공모 회사채: 증권신고서(채무증권) 제목 키워드
BOND_REPORT_KEYWORDS = ["증권신고서(채무증권)", "일괄신고추가서류(채무증권)"]

# 일반 회사채만 남기기 위한 필터
# 1) 채무증권 명칭에 이 단어가 있으면 제외 (메자닌·파생결합·조건부자본증권 등)
BOND_NAME_EXCLUDE = ["전환", "신주인수권", "교환", "파생결합", "조건부자본", "상각형", "유동화"]
# 2) 발행사 업종코드(KSIC 앞 2자리)가 금융업이면 제외: 64 금융업, 65 보험·연금, 66 금융보조
FINANCIAL_KSIC_PREFIX = ("64", "65", "66")
# 3) 업종코드를 못 얻었을 때 회사명으로 보조 판별
FINANCIAL_NAME_HINTS = ["은행", "카드", "캐피탈", "증권", "보험", "생명", "화재", "금융지주",
                        "저축은행", "파이낸셜", "리스", "자산운용", "투자신탁", "유동화", "제일차", "제이차"]

# ── 공공데이터포털: 금융위원회_채권기본정보 (사모 회사채) ──
BOND_INFO_URL = "https://apis.data.go.kr/1160100/GetBondIssuInfoService_V2/getBondBasiInfo_V2"
BOND_INFO_ROWS = 1000          # 한 번에 받는 건수
BOND_INFO_MAX_PAGES = 30       # 기준일 하루치 최대 페이지 (호출 한도 보호)
BOND_INFO_BASDT_DAYS = 5       # 매일 저녁 최근 며칠치 기준일을 훑을지
BOND_ISSUE_LOOKBACK_DAYS = 1095  # 조달내역 화면용: 발행일이 최근 3년 안인 잔존 사모채를 저장

# 실제 데이터 확인 결과(2026-10-23 기준일 29,261건)에 맞춘 허용 목록
# 종목종류: 일반회사채만 (금융채·MBS·SLBS·유동화SPC채·유사집합투자기구채·특수채 제외)
PRIVATE_BOND_KINDS = ["일반회사채"]
# 특이채권종류: 주식 관련 없는 것만 (CB·EB·BW·이익참가 제외)
PRIVATE_BOND_PCLR_OK = ["주식관련해당사항없음", "", "None"]
# 모집방법: '사모' + 'CBO기초사모'(신용보증기금 P-CBO에 편입되는 사모채, 화면에 P-CBO 태그)
PRIVATE_BOND_METHODS = ["사모", "CBO기초사모"]

# 위 조건을 통과한 뒤에도 종목명·업종명에 이 단어가 있으면 제외
PRIVATE_BOND_EXCLUDE = ["금융", "은행", "카드", "캐피탈", "유동화", "ABS", "단기", "CD", "CP",
                        "국채", "지방채", "특수채", "통화안정", "보험", "전환", "교환", "신주인수권",
                        "조건부", "파생", "커버드"]


# ── 기업 프로필 (profiles 모드) ──
DART_RUN_BUDGET = 4000          # profiles 한 번 실행에서 쓸 최대 OpenDART 호출 수 (일일 한도 20,000, 밤사이 여러 번 실행)
PROFILE_TIME_BUDGET_MIN = 100  # profiles 한 번 실행의 최대 시간(분). 넘으면 저장하고 다음 실행에서 이어감
PROFILE_WORKERS = 6             # 원문·재무 동시 다운로드 수
MEZZ_DOC_DAYS = 730            # 메자닌 원문(투자자·콜/풋)을 읽을 범위(최근 N일)
MEZZ_DOC_MAX = 400             # 한 번 실행에 읽을 메자닌 원문 수
BOND_DOC_DAYS = 400             # 수요예측일·신용등급을 원문에서 읽을 공모채 범위(최근 N일)
PROFILE_RECHECK_DAYS = 7        # 프로필이 있는 회사는 이 주기로 새 정기보고서가 나왔는지 확인
COMPANY_INFO_RECHECK_DAYS = 60  # 기업개황(대표·주소·업종코드 등) 재확인 주기
PROFILE_TARGET_DAYS = 1095      # 최근 3년 안에 조달 기록이 있는 회사 = 프로필 대상
WATCHLIST_FILE = "watchlist.txt"  # 저장소 루트. 한 줄에 회사명 또는 종목코드 하나 (관심 회사)

KIND_MIN_ROWS = 1500         # 이보다 적게 받으면 형식 변경으로 보고 기존 목록 유지
KIND_LIST_URL = "https://kind.krx.co.kr/corpgeneral/corpList.do"
STOCK_PRICE_URL = "https://apis.data.go.kr/1160100/GetStockSecuritiesInfoService_V2/getStockPriceInfo_V2"

# 다중회사 주요계정 → 사이트에서 쓰는 이름
MULTI_ACCOUNTS = {
    "유동자산": "cur_assets", "비유동자산": "noncur_assets", "자산총계": "assets",
    "유동부채": "cur_liab", "비유동부채": "noncur_liab", "부채총계": "liab",
    "자본금": "capital", "이익잉여금": "retained", "자본총계": "equity",
    "매출액": "revenue", "영업이익": "op", "법인세차감전 순이익": "pretax", "당기순이익(손실)": "ni",
    "당기순이익": "ni",
}
REPORT_CODES = {"11013": "1Q", "11012": "반기", "11014": "3Q", "11011": "연간"}
