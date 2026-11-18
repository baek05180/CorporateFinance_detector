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

### 4. Cloudflare Pages 프로젝트 (`ecm-board`) → Settings
- Variables and Secrets: `GH_TOKEN`(Secret, 위 토큰), `GH_REPO` = `baek05180/CorporateFinance_detector`, `GH_REF` = `main`
- Bindings → R2 bucket: 변수명 `PROPOSALS`, 버킷 `ecm-proposals`
- 설정 후 한 번 재배포 (Actions → collect 수동 실행 또는 site 워크플로)

### 5. Cloudflare Access (제안서 기능만 잠금)
Zero Trust → Access → Applications → Self-hosted
- 도메인: 사이트 도메인, 경로 2개 — `/proposal*`, `/api/proposal*`
- Policy: Allow → Emails (또는 Emails ending in `@회사도메인`)
- 나머지 사이트는 그대로 공개

API 는 Access 헤더가 없으면 401 로 거절한다(설정 누락 대비). 로컬 개발 때만 `REQUIRE_ACCESS=0`.

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
