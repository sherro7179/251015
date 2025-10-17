# VBA Macros → Python UI

이 폴더는 기존 VBA 매크로 6개를 Python 모듈과 Tkinter UI로 옮긴 버전입니다.  
`python -m vba_to_python.ui` 또는 `python vba_to_python/ui.py`를 실행하면
버튼이 있는 간단한 창이 표시되고, 각 버튼은 VBA 매크로와 동일한 작업을 수행합니다.

## 준비 사항

1. Python 3.10 이상
2. 필수 패키지 설치
   ```bash
   pip install openpyxl
   ```
   (Windows Excel COM을 사용하지 않으므로 `.xls` 파일은 지원하지 않습니다.)
3. VBA에서 사용하던 컨트롤 통합문서(`파일`, `IO_name`, `data_update` 시트를 포함)를 준비합니다.
   - 모든 시트 이름을 **한글 "파일"** 로 통일했습니다.
   - `파일!B2` 에는 대상 폴더 경로가 저장됩니다.
   - `data_update` 시트는 `Value Find` 실행 시 자동으로 초기화 후 채워집니다.

## 실행 방법

```bash
cd C:\Users\user\Desktop\251015   # 저장소 루트(부모 폴더)에서 실행
python -m vba_to_python.ui
```

> 폴더를 옮겨서 실행하는 경우에도 **항상 `vba_to_python` 패키지를 포함하는 상위 경로**를 현재 작업 폴더로 잡아야 합니다.  
> 만약 `vba_to_python` 폴더 내부에서 실행하려면 `python ui.py`와 같이 직접 스크립트를 호출하세요.

### UI 동작

1. **컨트롤 통합문서 선택** 버튼으로 기준 Excel을 지정합니다.
2. 각 작업 버튼은 다음 VBA 매크로를 그대로 구현합니다.
   | 버튼 | VBA 매크로 | 설명 |
   | --- | --- | --- |
   | Select Folder Path | `SelectFolderPath` | 폴더 선택 후 `파일!B2` 입력 및 파일 목록 갱신 |
   | List Excel Files | `ListExcelFilesInFolder` | `파일!B2` 기준으로 Excel 파일 재스캔 |
   | Update Files | `UpdateFiles` | Test Case ID 갱신 |
   | IO Change | `IOCHANGE` | `IO_name` 시트 매핑대로 문자열 치환 |
   | Value Find | `Value_find` | 조건 검색 결과를 `data_update` 시트에 기록 |
   | Change Value | `Change_Value` | `data_update` 정보를 이용해 셀 값 일괄 변경 |

3. 작업 완료 시 메시지가 팝업되고, 오류가 있으면 상세 정보를 보여줍니다.  
   `Change Value`는 실패한 항목을 최대 5건까지 팝업에 표시합니다.

## 모듈 구성

```
vba_to_python/
├── actions/              # 각 버튼별 비즈니스 로직
├── constants.py          # 시트명 및 셀 주소 상수
├── utils.py              # 공통 함수 (경로, 워크북 로딩 등)
└── ui.py                 # Tkinter UI 진입점
```

필요하다면 `vba_to_python/__init__.py` 를 통해 개별 함수를 직접 import 하여
스크립트나 테스트 코드에서 사용할 수 있습니다.
