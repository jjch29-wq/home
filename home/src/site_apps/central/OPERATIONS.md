# 중앙지사 운영 데이터 분리

운영 실행 환경의 `PMI_CENTRAL_DATA_ROOT`에 코드 체크아웃 밖에 있는
중앙지사 전용 폴더의 절대 경로를 지정한다. 운영 앱에서 이 설정은
기존 `src/sync_config.json`보다 우선한다. 지정한 드라이브와 경로를
그대로 사용하므로 PC마다 드라이브 문자가 다르면 각 PC에서 설정한다.
개발 모드에서는 운영 경로 설정을 무시한다.

전환 전에 앱과 자동 동기화를 종료한다. 사용 중인 운영 원본을 확인하고
대상 폴더에 복사한다. 기존 원본을 보관한 상태에서 앱을 재실행해
동일한 날짜의 작업 이력, 자재 재고와 사진을 확인한다.

- 루트에는 `daily_work_history.json`, 기존 `report_issue_history.json`과
  중앙지사 제출자료 폴더를 둔다.
- 기성 내보내기 자료를 사용하는 경우 `data/billing_export.json`도 복사한다.
- `data` 하위에는 `Material_Inventory.xlsx`와 `process_photos`를 둔다.
- 모바일 동기화를 사용하면 `data` 하위에 운영용
  `google_sync_config.json`과 기존 `.mobile_sync_state.json`도 복사한다.
  `.mobile_pending_deletes.json`이 있으면 재고와 동일한 시점의 파일을 복사한다.
  외부 운영 경로가 설정된 경우 모바일 동기화는 이 재고 사본만 사용한다.

복원 전에도 앱과 동기화를 종료한다. 현재 상태를 별도 폴더에 보관하고,
검증한 백업에서 같은 시점의 재고, 작업 이력, 사진 및 동기화 상태를
복원한다. 앱에서 재고 합계와 최근 작업을 확인한 뒤 동기화를 재개한다.

## 전환 후 점검

회사 PC에서는 `F:`, 집 PC에서는 `G:`의 중앙지사 공유 폴더를
지정한다. 두 PC가 같은 공유 폴더를 사용하면 모바일 설정은 한 번만
복사한다. 새 프로세스로 운영 앱을 실행해 상태 영역의 마지막 모바일
동기화 성공 시각과 콘솔의 시작 동기화 소요 시간을 확인한다.

## 작업 ID와 동시 작업 확인

- 모바일 작업에 `ClientRequestID`가 있으면 같은 초에 입력한 다른
  작업과 구분한다. ID가 없는 기존 작업은 입력 시각으로 대조한다.
- 기존 ID를 이관할 때는 PC와 모바일에서 모두 입력 시각이 유일하고
  날짜, 현장, 작업자, 검사방법, 수량, 관경, 단가가 일치하는 행을
  먼저 확인한다. 장비명 등 각 원본의 업무 값은 보존한다. 짝이
  모호하거나 한쪽에만 있는 행은 개별 확인 후 처리한다.
- 동시 작업은 운영본과 분리한 테스트 재고로 확인한다. 회사와 집에서
  같은 테스트 재고를 열고, 한쪽에서 저장한 다음 다른 쪽의 이전
  화면에서 저장을 시도한다. 최신 기록이 덮어써지지 않고, 저장을
  시도한 PC의 입력이 보존되는지 확인한다.
- 모바일 삭제 재시도는 테스트 시트에서 확인한다. 같은 초의 두 작업을
  만든 뒤 한쪽에서 한 작업만 삭제한다. 동기화 실패 후 실행을 다시
  시작해 대기 요청이 남아 있는지, 재시도 후 대상 ID만 사라지는지 확인한다.
  앱 내부 잠금은 한 프로세스의 동기화 요청에 적용된다.

## FastAPI 모바일 서버

직접 실행한 FastAPI 서버는 `127.0.0.1:8000`에서만 접속할 수 있고,
모바일 페이지와 API는 같은 주소에서 제공한다. 휴대폰 등 다른
기기에서 운영에 사용하려면 HTTPS와 사용자 인증이 있는 게이트웨이
뒤에 두고 모바일 페이지와 API 모두에 접근 제어를 적용한다.
FastAPI 서버 자체에는 사용자 인증이 없다. 이 설정은 Google Apps
Script 웹 앱의 배포 및 접근 설정과 별개다.

## Windows 회귀 테스트

Tk 8.6이 설치된 프로젝트 가상환경에서, 전체 실행 중 Tk 초기화
오류가 발생하면 테스트 프로세스에 라이브러리 경로를 명시한다.

```powershell
$pythonBase = .venv/Scripts/python.exe -B -c "import sys; print(sys.base_prefix)"
$env:TCL_LIBRARY = Join-Path $pythonBase 'tcl/tcl8.6'
$env:TK_LIBRARY = Join-Path $pythonBase 'tcl/tk8.6'
$tests = @(Get-ChildItem home/tests -Filter 'test_central_*.py' | Select-Object -ExpandProperty FullName)
.venv/Scripts/python.exe -B -m pytest -q -p no:cacheprovider @tests home/tests/test_site_isolation.py
```
