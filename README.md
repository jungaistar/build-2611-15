# build-2611-15 · OEM 원소재 연동제 자동화

생성형 AI 업무자동화 전문가과정 **2차수** 수강생 프로젝트 15 — 기획 양재현.
(12번 「도면 기반 원소재 단가 영향 분석」에서 도면을 쓸 수 없게 되어 바꾼 과제입니다. 12번은 그대로 둡니다.)

**바로 쓰기: https://build-2611-15.jobability.co.kr/**

업체별 계약 수식 · 조정요건과 품번별 기준 파트단가 · 원소재 사용량을 나눠 등록해 두고,
분기 · 반기마다 원소재 Index 엑셀을 올리면 품번별 연동액과 변경 단가를 계산해 업체 요청단가와 비교하고 팀장 보고용 Word 문서를 만듭니다.
자료는 브라우저(IndexedDB)에만 저장되고 서버로 올라가지 않습니다. 외부 CDN · AI 호출이 없습니다.

| 파일 | 역할 |
|---|---|
| `js/linkage.js` | 계산 엔진 — 품번 엑셀 읽기 · 계약서 글 읽기 · 수식 계산기 · 가로형 Index · Index 맞추기 · 연동액 계산 · 검증 · 보고서 값. **계산은 전부 여기** |
| `js/costidx.js` | 12번에서 가져온 Index 통합문서 읽기(머리 블록 · 세로형 · 원화 환산 · 재질 계열 · 변동 상관). 고칠 때는 12번과 함께 본다 |
| `js/app.js` | 화면 6개(대시보드 · 업체 · 계약 · 품번 · 사용량 · Index · 연동제 계산 · 보고서 · 이력)와 IndexedDB 저장 |
| `js/template.js` | Word 보고서 기본 틀(base64) — `samples/make-template.py` 가 만든다 |
| `lib/` | SheetJS · Mammoth · PizZip · docxtemplater · pdf.js (동봉) |
| `samples/` | 가짜 예제 엑셀 3종 · 보고서 기본 틀 · 생성 스크립트 |
| `proto/` | 기획자가 만든 시제품 V2(참고용 원본) |
| `test/logic.test.js` | 계산 규칙 테스트 `node test/logic.test.js` |
| `test/smoke.browser.js` | 화면 연기 테스트(playwright) |
| `CLAUDE.md` | 작업 지시서 + 기획 프롬프트 원문 |

예제 데이터는 실제 사내 데이터가 아닙니다(가짜 업체 2곳 · 품번 6개 · 가짜 Index).

## 기획서 · 버전

- 사이트 「기획서」 페이지: https://build-2611-15.jobability.co.kr/plan.html — 워드 판 내려받기
- 워드 판: `docs/기획서/build-2611-15_양재현_<제목>_기획서_v<버전>.docx` (버전마다 남김)
- 원본: `docs/plan.json` → 생성기 `build-2611-00/tools/plandoc/make_plan.py`
