# ECM BOARD

메자닌(CB·BW·EB) 발행결정, 유상증자 결정, 일반 회사채(공모·사모) 발행을 모아 보여주는 사이트.
서버 없이 **GitHub Actions(수집) + Cloudflare Pages(호스팅)** 로 돌아가며 비용은 0원.

| 데이터 | 출처 | 갱신 |
|---|---|---|
| CB·BW·EB 발행결정 | OpenDART 주요사항보고서 | 평일 07~20시 5분 간격 감시 + 20:15 정기 갱신 |
| 유상증자 결정 | OpenDART 주요사항보고서 | 〃 |
| 회사채 공모 | OpenDART 증권신고서(채무증권) | 〃 |
| 회사채 사모 | 공공데이터포털 금융위원회_채권기본정보 (모집방법 = 사모) | 매일 20:15 (발행 다음 영업일 반영) |

금융채(은행·카드·캐피탈 등), ABS, 단기사채, 주식연계채권은 회사채 목록에서 제외한다.

---

## 처음 설정하기 (약 20분)

### 1. 공공데이터포털 키 받기 (사모 회사채용)
1. https://www.data.go.kr/data/15059592/openapi.do (금융위원회_채권기본정보) 접속 → **활용신청**
2. 개발계정은 자동 승인. 마이페이지 → 데이터활용 → Open API → 개발계정에서 **일반 인증키(Decoding)** 복사

### 2. Cloudflare 토큰 만들기 (사이트 배포용)
1. Cloudflare 대시보드 → 오른쪽 위 프로필 → **My Profile → API Tokens → Create Token**
2. **Create Custom Token** → 권한: `Account` / `Cloudflare Pages` / `Edit` → 토큰 복사
3. 대시보드 첫 화면(Workers & Pages) 오른쪽의 **Account ID** 복사

### 3. GitHub 저장소에 비밀값 넣기
저장소 → **Settings → Secrets and variables → Actions → New repository secret** 로 4개 추가

| 이름 | 값 |
|---|---|
| `DART_API_KEY` | OpenDART 인증키 (40자리) |
| `DATA_GO_KR_KEY` | 공공데이터포털 일반 인증키 (Decoding) |
| `CLOUDFLARE_API_TOKEN` | 2단계에서 만든 토큰 |
| `CLOUDFLARE_ACCOUNT_ID` | Cloudflare Account ID |

(선택) 사이트 주소를 바꾸고 싶으면 같은 화면의 **Variables** 탭에서 `CF_PAGES_PROJECT` = 원하는 이름 (기본값 `ecm-board`)

### 4. 과거 데이터 채우기
저장소 → **Actions → collect → Run workflow** → mode = `backfill`, days = `90` → 실행

끝나면 `https://ecm-board.pages.dev` (또는 정한 이름)에서 확인. 이후로는 자동으로 돈다.

---

## 확인할 것

- **Actions 실행 결과**: 각 실행의 Summary에 신규 건수와 호출 횟수가 나온다.
- **`data/diag.json`**: 사모채 API가 실제로 주는 `종목종류 | 특이채권종류` 값과 건수. 일반 회사채인데 빠졌거나,
  금융채인데 들어온 게 있으면 `collector/config.py` 의 `PRIVATE_BOND_EXCLUDE` 를 고친다.
- 사이트 이름: `web/build.py` 의 `SITE_NAME`

## 구조

```
collector/   수집기 (config.py 에서 키워드·기간 조정)
  run.py       watch / daily / backfill 실행
  dart.py      OpenDART 클라이언트
  datagokr.py  공공데이터포털 채권기본정보 클라이언트
  normalize.py 원본 응답 → 공통 레코드
  store.py     data/*.json 저장 (바뀐 경우에만 기록)
web/         template.html + build.py → public/index.html
data/        수집 결과 (Actions가 커밋)
public/      배포되는 사이트
tests/       가짜 응답으로 전체 흐름 점검: python -m tests.test_pipeline
```

## 로컬에서 돌려보기

```bash
pip install -r requirements.txt
export DART_API_KEY=...        # Windows PowerShell: $env:DART_API_KEY="..."
export DATA_GO_KR_KEY=...
python -m collector.run --mode daily --days 7
python -m web.build            # public/index.html 생성 → 브라우저로 열기
```

## 알려진 한계 (다음 단계)

- 콜옵션 비율, 리픽싱 조건 세부는 API에 없어 공시 원문 파싱이 필요하다.
- 유상증자 금액은 API에 금액 필드가 없어 '자금조달 목적' 합계로 표시한다.
- GitHub 정기 실행은 붐비는 시간에 몇 분씩 늦어질 수 있다.
- 공공데이터포털 개발계정은 일일 호출 한도가 작다. 한도가 부족하면 활용사례를 등록해 운영계정으로 올린다.
