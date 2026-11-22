# proposal-vault (비공개)

공모사채 발행 제안서 생성 전용 **비공개** 저장소. 사이트의 `공모사채 제안서 작성 (PPT)` 버튼이 이 저장소의 `proposal` 워크플로를 실행한다.

- `template/proposal_template.pptx` — 원본 신한 템플릿 (공개 저장소에 올리지 않음)
- `.github/workflows/proposal.yml` — 공개 저장소(CorporateFinance_detector) 최신 main 의 코드·데이터·명세로 PPT 생성, 결과는 7일 보관

워크플로 원본은 공개 저장소의 `vault/` 폴더에 있다. 그쪽을 고치면 이 저장소에도 같은 내용을 반영한다.
