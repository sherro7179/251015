# FastAPI 기반 결재 서류 사전 검토 서버

FastAPI + Swagger 구성으로 결재 문서를 빠르게 검증할 수 있는 경량 서버입니다.
v2 규정 패키지에서 요약한 규칙을 로드하여 **양식 구조**와 **필수 규정 문구**를 체크합니다.

## 1. 준비
`ash
cd eapproval_fastapi
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
`

## 2. 실행
`ash
uvicorn app.main:app --reload
`

브라우저에서 [http://localhost:8000](http://localhost:8000) 을 열면
AI 기반 결재 서류 사전 검토 대시보드를 사용할 수 있습니다.
API 문서는 [http://localhost:8000/docs](http://localhost:8000/docs) 에서 확인하세요.

## 3. 주요 API
| Method | Path                        | 설명 |
| ------ | --------------------------- | ---- |
| GET    | /health                   | 서버 상태 확인 |
| GET    | /api/v1/rules             | 적용 중인 규칙 메타데이터 |
| POST   | /api/v1/rules/reload      | ules_bundle_v2.json 다시 로드 |
| POST   | /api/v1/validate          | JSON 기반 규칙 검증(레거시) |
| POST   | /api/v1/documents/inspect | DOCX 업로드 후 양식/규정 검사 |

### DOCX 검사 예시
`
POST /api/v1/documents/inspect
Content-Type: multipart/form-data
  doc_type = EXR
  document = (파일) 결재서_샘플_지출품의_v2.docx
`

### 응답 예시
`json
{
  "passed": true,
  "doc_type": "EXR",
  "filename": "결재서_샘플_지출품의_v2.docx",
  "structure": {
    "ok": true,
    "coverage": 1.0,
    "checked": ["결재서_샘플_지출품의_v2", "문서번호", "결재선(예시)", ...],
    "missing": []
  },
  "regulation": {
    "ok": true,
    "checked": ["견적서 2부 이상", "행사 계획서"],
    "missing": []
  }
}
`

## 4. 규칙 수정
1. data/rules_bundle_v2.json을 수정하여 규정을 업데이트합니다.
2. 서버 실행 중이라면 /api/v1/rules/reload (POST)로 즉시 반영할 수 있습니다.

## 5. 향후 확장 아이디어
- **문서 자동 파싱**: DOCX 내용을 분석해 필수 필드와 위험 플래그를 자동 추출하도록 확장합니다.
- **검증 로그 저장**: SQLite/PostgreSQL에 Pass/Fail 이력을 저장해 추적성을 확보합니다.
- **배치·연동**: 메시지 큐나 파일 드롭 폴더를 붙여 다량의 결재서를 일괄 검증하거나 전자결재 시스템과 API/Webhook으로 연동할 수 있습니다.
