# SMB Precheck VBA Toolkit · v0.1

이 폴더는 **중소기업 결재 서류 사전 검토 자동화**를 위한 VBA 코드와 로그를 보관합니다.  
v0.1에서는 시트 기반 버튼 UI를 기본으로 제공하며, 동일한 모듈을 활용해 리본/Ribbon 또는 사용자 폼으로 확장할 수 있습니다.

---

## 폴더 구조

| 경로 | 설명 |
| --- | --- |
| `modules/modGlobals.bas` | 공통 상수, 시트/열 위치, 구조체 선언 |
| `modules/modUtils.bas` | 경로/백업/진행률/로그/필터/선택 유틸리티 |
| `modules/modTasks.bas` | 실질 작업 로직(케이스 ID 재번호, IO 치환, 값 찾기·변경 등) |
| `modules/modDashboard.bas` | UI 버튼에서 호출할 엔트리 매크로(`Command_*`) |
| `log/` | 작업 중 오류가 발생했을 때만 생성되는 로그 파일 보관 폴더 |

> 아직 모듈은 `.bas` 형태로만 제공되며, 버튼이나 리본은 직접 추가해야 합니다.

---

## 사전 준비

1. **시트 이름 통일**  
   컨트롤 통합문서에 아래 시트를 반드시 포함시키세요.
   `파일`, `IO_name`, `data_update`, `script_move`

2. **`파일` 시트 구성**  
   - B2: 대상 폴더 경로  
   - B4: 포함 필터(세미콜론 `;` 구분), B5: 제외 필터  
   - 7행 헤더: `File Name | Original Path | Include? | Status | Message`  
   - 데이터는 8행부터 채워집니다.

3. **보조 시트**  
   - `IO_name`: [원본, 치환값] 쌍 입력  
   - `data_update`: 값 찾기 결과 및 수정 값 입력 테이블(A1:F1 자동 헤더)  
   - `script_move`: 하위 폴더 목록 출력용

4. **폴더 구조**  
   - 작업 대상 폴더 하위에 `_backup`, `_processed`가 자동 생성됩니다.  
   - 오류가 발생하면 `vba/log/SMB_yyyymmdd_hhnnss.log`가 생성됩니다.

---

## 주요 매크로 (시트 버튼용)

| 매크로 | 설명 |
| --- | --- |
| `Command_SelectFolderPath` | 폴더 선택 후 `파일` 시트에 Excel 파일 목록 로드 |
| `Command_UpdateFiles` | 선택 파일의 Test Case ID 재번호 (`Task_UpdateFiles`) |
| `Command_IOChange` | `IO_name` 시트의 매핑으로 문자열 치환 (`Task_IOChange`) |
| `Command_ValueFind` | 조건(B10/B12)에 맞는 값을 찾아 `data_update`에 기록 |
| `Command_ChangeValue` | `data_update` 정보를 기반으로 값 일괄 변경 |
| `Command_ScriptMoveFolderPath` | 하위 폴더 목록을 `script_move` 시트에 작성 |

모듈을 VBA 프로젝트에 추가한 뒤, 시트에 도형 버튼을 삽입하거나 Alt+F8에서 해당 매크로를 실행하면 됩니다.

---

## 실행 흐름 예시

1. `Command_SelectFolderPath`로 폴더 선택 → `_backup`, `_processed` 자동 생성
2. `Include?` 열과 필터(B4/B5)를 조정하여 처리 대상 결정
3. `Command_UpdateFiles`, `Command_IOChange`, `Command_ValueFind` 등을 필요 순서대로 실행
4. `Command_ChangeValue`로 `data_update` 기반 일괄 수정 (`_processed` 사본에만 적용)
5. 각 작업 후 `Status`/`Message` 열에 결과 기록, 오류 발생 시 즉시 중단 및 로그 생성

---

## 향후 확장

- **Ribbon**: `modDashboard`의 `Command_*` 매크로를 리본 버튼에 연결하면 동일 기능을 리본에서 사용할 수 있습니다.
- **UserForm**: 파일 리스트/필터링 UI를 폼으로 구현하려면 `SelectedFileEntry` 구조체 배열을 활용해 `RunSelectedFilesTask` 방식으로 호출하면 됩니다.

---

## 버전 이력

| 버전 | 내용 |
| --- | --- |
| v0.1 | 시트 기반 버튼 UI 및 공통 모듈 정의, 백업·로그·진행률 기능 포함 |

요구사항이나 개선점이 있으면 `modDashboard.bas` 또는 `modTasks.bas`에 주석으로 남겨 주세요. 다음 버전에서 Ribbon/UserForm 템플릿과 자동 버튼 생성 매크로를 제공할 예정입니다.
