# 공모사채 발행 제안서 — 사이트 연결 설정

신용등급이 있는 회사의 **조달 제안 보고서** 화면에 `공모사채 제안서 작성 (PPT)` 버튼이 생긴다.
버튼 → 팝업(`/proposal/?corp=…`)에서 일정·금액·만기 등을 입력 → GitHub Actions 가 PPT 를 만들어 → 팝업에서 내려받는다.

```
팝업 /proposal ──POST──▶ Pages Function /api/proposal ──workflow_dispatch──▶ Actions proposal.yml
     ▲                         │ R2: proposals/{id}/status.json                │ R2 에서 템플릿 받기
     └──GET /api/proposal/{id}─┘◀──────────── R2: proposal.pptx, status.json ◀──┘ tools.proposal_job
```

입력 항목·기본 문구·한미 이벤트·휴일·장표 켜기/끄기는 **`docs/proposal_spec.md` 의 `## 설정` 블록**에서 읽는다.
이 파일을 main 에 고치면 `site.yml` 이 사이트를 다시 배포하고(팝업 항목 반영), 다음 제안서부터 바로 적용된다.

## 1회 설정

### 1. R2 버킷 (비공개)
1. Cloudflare 대시보드 → R2 → 버킷 만들기: `ecm-proposals` (다른 이름이면 GitHub 변수 `R2_BUCKET` 에 입력)
2. 원본 템플릿 업로드: `templates/proposal_template.pptx` — 저장소가 공개이므로 **템플릿은 절대 커밋하지 않는다**
3. 버킷 → 설정 → Object lifecycle rules: 접두사 `proposals/`, 7일 후 삭제
4. R2 → API 토큰 관리 → 이 버킷 한정 **Object Read & Write** 토큰 생성 → Access Key ID / Secret 메모

### 2. GitHub (저장소 Settings → Secrets and variables → Actions)
| 종류 | 이름 | 값 |
|---|---|---|
| Secret | `R2_ACCESS_KEY_ID` | 위 R2 토큰의 Access Key ID |
| Secret | `R2_SECRET_ACCESS_KEY` | 위 R2 토큰의 Secret |
| Secret | `CLOUDFLARE_ACCOUNT_ID` | (이미 있음) |
| Secret (선택) | `ANTHROPIC_API_KEY` | 있으면 크레딧 스토리·예상 Q&A 초안을 Claude 가 작성, 없으면 ## |
| Variable (선택) | `R2_BUCKET` | 버킷 이름이 `ecm-proposals` 가 아닐 때 |

### 3. GitHub 토큰 (Pages Function → Actions 실행용)
GitHub → Settings → Developer settings → Fine-grained token
- Repository access: 이 저장소만
- Permissions: **Actions: Read and write** (그 외 없음)

### 4. Cloudflare Access (제안서 기능만 잠금)
Zero Trust → Access → Applications → Self-hosted
- 도메인: 사이트 도메인, 경로 2개 — `/proposal*`, `/api/proposal*`
  - `ecm-board.pages.dev` 주소로도 같은 경로 2개를 추가 (기본 주소로 우회하지 못하게)
- Policy: Allow → Emails → 본인(과 지정한 사람) 이메일
- 만든 뒤 애플리케이션의 **Application Audience (AUD) Tag** 와, Zero Trust → Settings 의 **팀 도메인**(`xxx.cloudflareaccess.com`)을 메모

### 5. Cloudflare Pages 프로젝트 (`ecm-board`) → Settings
- Variables and Secrets
  | 이름 | 값 |
  |---|---|
  | `GH_TOKEN` (Secret) | 3번 토큰 |
  | `GH_REPO` | `baek05180/CorporateFinance_detector` |
  | `GH_REF` | `main` |
  | `ACCESS_TEAM_DOMAIN` | 예: `myteam.cloudflareaccess.com` |
  | `ACCESS_AUD` | 4번 AUD Tag |
  | `ALLOWED_EMAILS` | 쓸 수 있는 이메일 (쉼표 구분) — Access 정책과 같은 목록 |
- Bindings → R2 bucket: 변수명 `PROPOSALS`, 버킷 `ecm-proposals`
- 설정 후 한 번 재배포 (Actions → site 또는 collect 수동 실행)

## 보안 구조
| 무엇을 | 어떻게 막나 |
|---|---|
| API 키·토큰 | 코드에 없음. GitHub Secret / Cloudflare Secret 에만. GitHub 토큰은 이 저장소 Actions 실행만, R2 토큰은 버킷 하나만 |
| 제안서 기능 사용 | Cloudflare Access 로그인 + API 가 Access 토큰(JWT) 서명·대상·만료를 직접 검증 + `ALLOWED_EMAILS` 목록. `*.pages.dev`·미리보기 주소로 와도 같은 검증 |
| 입력값(발행 계획·메모) | 비공개 R2 에만 저장. GitHub 에는 무작위 요청 ID 만 넘기고, 실행 기록에는 성공/실패만 찍음 (오류 상세는 R2 에만) |
| 템플릿·결과 PPT | 비공개 R2. 요청한 본인만 상태 확인·내려받기, 7일 뒤 삭제 |
| 코드·명세 md | **저장소가 공개면 누구나 볼 수 있음** (키는 없으므로 남이 내 계정으로 돌릴 수는 없음). 숨기려면 저장소를 비공개로 |

`ANTHROPIC_API_KEY` 를 넣으면 팝업 메모·공시 정보가 Anthropic API 로 전송된다(초안 작성용). 사내 규정상 외부 전송이 어렵다면 넣지 않는다.

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
