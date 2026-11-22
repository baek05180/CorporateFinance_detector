# 공모사채 발행 제안서 — 사이트 연결 설정 (결제 없는 구성)

신용등급이 있는 회사의 **조달 제안 보고서** 화면에 `공모사채 제안서 작성 (PPT)` 버튼이 생긴다.
버튼 → 팝업에서 GitHub 로 로그인(허용된 아이디만) → 일정·금액·만기 입력 → 비공개 저장소의 GitHub Actions 가 PPT 생성 → 팝업에서 내려받기.

```
팝업 /proposal ──POST /api/proposal──▶ Pages Function ──workflow_dispatch──▶ 비공개 저장소 proposal-vault
  │ GitHub 로그인 (/api/auth/*)            │ (입력값은 비공개 실행 입력으로만 전달)     │ 공개 저장소 최신 main 의 코드·데이터·명세
  └──GET /api/proposal/{id}────────────────┴──────── 실행 결과물(7일 보관) 조회·내려받기 ◀──┘ + 비공개 템플릿으로 생성
```

쓰는 서비스는 GitHub(무료 계정)와 지금 사이트가 쓰는 Cloudflare Pages(무료)뿐이다. 결제수단 등록이 필요 없다.
- 비공개 저장소의 Actions: 무료 계정 기준 월 2,000분 · 저장 500MB 포함. 제안서 1건 5~6분 → 월 300건 안팎. 결제수단이 없으면 한도를 넘어도 과금되지 않고 실행만 멈춘다.
- 데이터 수집(collect)은 공개 저장소라 이 한도와 무관.

입력 항목·기본 문구·한미 이벤트·휴일·장표 켜기/끄기는 **`docs/proposal_spec.md` 의 `## 설정` 블록**에서 읽는다.
main 에 고치면 `site.yml` 이 사이트를 다시 배포하고(팝업 항목), 다음 생성부터 바로 적용된다(생성할 때마다 최신 main 을 받음).

## 1회 설정

### 1. 비공개 저장소 `proposal-vault`
1. GitHub → New repository → 이름 `proposal-vault`, **Private**
2. 내용: 이 저장소의 `vault/` 폴더 그대로 (`.github/workflows/proposal.yml`, `README.md`) + `template/proposal_template.pptx`(원본 템플릿)
3. (선택) Settings → Secrets and variables → Actions → `ANTHROPIC_API_KEY` — 있으면 크레딧 스토리·예상 Q&A 초안을 Claude 가 작성, 없으면 ##

### 2. GitHub 토큰 (사이트가 비공개 저장소 실행·결과 조회에 사용)
GitHub → Settings → Developer settings → Personal access tokens → **Fine-grained tokens** → Generate
- Repository access: **Only select repositories → `proposal-vault`**
- Permissions → Repository → **Actions: Read and write** (그 외 없음)
- 만료일: 원하는 기간 (만료되면 새로 만들어 Pages 의 `GH_TOKEN` 만 바꾸면 됨)

### 3. GitHub OAuth 앱 (팝업 로그인용)
GitHub → Settings → Developer settings → **OAuth Apps** → New OAuth App
- Application name: `ECM 제안서` (자유)
- Homepage URL: 사이트 주소 (예: `https://ecm-board.pages.dev`)
- Authorization callback URL: `https://사이트주소/api/auth/callback`
- 만든 뒤 **Client ID** 메모, **Generate a new client secret** → 메모

### 4. Cloudflare Pages 프로젝트 (`ecm-board`) → Settings → Variables and Secrets
| 이름 | 종류 | 값 |
|---|---|---|
| `GITHUB_CLIENT_ID` | 텍스트 | 3번 Client ID |
| `GITHUB_CLIENT_SECRET` | **Secret** | 3번 Client secret |
| `SESSION_SECRET` | **Secret** | 임의의 긴 문자열 32자 이상 (비밀번호 생성기로) |
| `ALLOWED_GITHUB_USERS` | 텍스트 | 쓸 수 있는 GitHub 아이디, 쉼표 구분 (예: `baek05180`) |
| `GH_TOKEN` | **Secret** | 2번 토큰 |
| `VAULT_REPO` | 텍스트 | `baek05180/proposal-vault` |

설정 후 재배포(공개 저장소 Actions → site 또는 collect 수동 실행, 또는 PR 머지).

## 보안 구조
| 무엇을 | 어떻게 막나 |
|---|---|
| API 키·토큰 | 코드에 없음. Cloudflare Secret / GitHub Secret 에만. GitHub 토큰은 `proposal-vault` 의 Actions 만 다룰 수 있음 |
| 제안서 기능 사용 | GitHub 로그인(2단계 인증 그대로) + `ALLOWED_GITHUB_USERS` 목록. 로그인 쿠키는 서버 비밀값으로 서명, 12시간 유효. 다른 사이트에서 보낸 요청은 차단 |
| 입력값(발행 계획·메모) | 비공개 저장소의 실행 입력으로만 전달 — 저장소 관리자만 볼 수 있음. 공개 저장소에는 남지 않음 |
| 템플릿·결과 PPT | 비공개 저장소. 결과는 요청한 본인만 상태 확인·내려받기, 7일 뒤 자동 삭제 |
| 코드·명세 md | 공개 저장소 — 누구나 볼 수 있음(키는 없으므로 남이 내 계정으로 돌릴 수 없음). md 에 고객사 정보·기밀 문구는 넣지 않는다 |

`ANTHROPIC_API_KEY` 를 넣으면 팝업 메모·공시 정보가 Anthropic API 로 전송된다(초안 작성용, 유료 API). 사내 규정상 외부 전송이 어렵다면 넣지 않는다.

## 로컬에서 만들어 보기
```bash
pip install -r requirements.txt -r requirements-proposal.txt
python -m web.build            # public/data/co 생성
echo '{"corp":"00126566","issue":"2027-01-15","base":3000,"max":4000,"split":"2:600,3:1500,5:900","priority":3}' > in.json
python -m tools.proposal5 --template 템플릿.pptx --inputs in.json --out out.pptx
```
LibreOffice(`soffice`)와 `pdftoppm` 이 있으면 회색 요약 박스 여백 정리(tighten)까지 한다.

## 손으로 채우는 값 (빨간 ##)
개별민평·예상 발행금리, 수수료(미입력 시), 향후 전망(엑셀 연동 전), FedWatch 확률(CME 약관상 자동 수집 안 함 — CME FedWatch → Probabilities 에서 확인해 입력), 발행 시사점 한 줄.
