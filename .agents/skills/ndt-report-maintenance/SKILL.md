---
name: ndt-report-maintenance
description: 비파괴검사보고서(RT, PAUT, MT, PT, UT, PMI) 앱의 코드·화면·성적서 서식·발행 이력·출력 기능을 유지보수할 때 사용한다. 보고서 양식 훼손 방지, 발행 이력 보존, 로고/서명 설정 및 회귀 테스트 절차를 준수한다.
---

# 비파괴검사보고서 시스템 유지보수

## 작업 시작

- 현재 작업 디렉터리와 저장소 루트를 확인한다 (`C:\Users\jjch2\Desktop\PMI`).
- 활성 브랜치가 `비파괴검사보고서-개발` (또는 작업 목적에 맞는 브랜치)인지 확인한다 (`git branch --show-current`).
- 기존 변경 사항(`git status --short`)을 확인하고, 다른 사업장 전용 데이터나 임시 파일을 섞지 않는다.
- 이 스킬은 비파괴검사 성적서·보고서 발행 시스템 유지보수 절차이며 일반적인 비파괴 기술 질의에는 적용하지 않는다.

## 코드 위치와 경계

비파괴검사보고서 앱의 핵심 코드는 `home/src/report_apps/ndt_report/` 아래에 독립 보관된다.

- **실행기**:
  - `home/src/비파괴검사보고서.py` (외부 바로가기 및 호환용 런처)
  - `home/src/report_apps/ndt_report/src/비파괴검사보고서.py` (보고서 앱 메인 구현체)
- **검사방법별 레이아웃 및 엑셀 생성 모듈**:
  - RT (방사선투과검사 / 가스공사 KOGAS): `src/rt_report_layout.py`
  - PAUT (위상배열초음파검사): `src/paut_report_layout.py`, `src/paut_fields.py`
  - MT (자분탐상검사): `src/mt_report_layout.py`, `src/mt_footer.py`
  - PT / UT / PMI 및 공통 NDT 처리: `src/ndt.py`, `src/worksheet_xml.py`
- **발행 이력 및 서명/인장**:
  - 보고서 발행 이력 관리: `src/report_issue_history.py`
  - 서명/인장 이미지 처리: `src/report_signatures.py`
  - 도장/로고 통합 설정: `home/config/logo_settings_unified.json`
- **외부 현장 일보 가져오기 (Read-Only)**:
  - `src/ndt_history_import.py`: 중앙지사/경인사무소 일일작업기록을 읽기 전용으로 검색·불러오기
- **데이터베이스 및 설정**:
  - 발행 이력 DB: `home/src/report_apps/ndt_report/data/report_issue_history.json`
  - 구글 드라이브 동기화 설정: `home/src/report_apps/ndt_report/sync_config.json`
- **회귀 검증 테스트**:
  - `home/src/report_apps/ndt_report/tests/test_*.py`

## 핵심 비즈니스 규칙 및 데이터 보존

1. **보고서 서식 및 엑셀 템플릿 보존**:
   - 엑셀 시트(`Cover`, `001`, `Detail` 등)의 인쇄 영역, 행 높이, 열 너비, 폰트, 테두리 스타일, 결재란 규격을 훼손하지 않는다.
   - XML 직접 조작(`worksheet_xml.py`) 시 파일 손상(Corruption)이 발생하지 않도록 네임스페이스와 닫는 태그를 엄격히 검증한다.
2. **발행 이력(`report_issue_history.json`) 무결성**:
   - 보고서 번호, 검사일자, 라인/도면, Joint, 규격, 판정 결과를 스냅샷으로 보존한다.
   - 보고서 수정 시 이전 값은 `revisions`에 기록하고, 삭제 시에는 물리적 삭제가 아닌 `deleted_at`을 기록하는 소프트 삭제 원칙을 준수한다.
   - 원본 일보 데이터나 이미 생성된 엑셀 출력물을 임의로 덮어쓰거나 삭제하지 않는다.
3. **사업장 데이터 보호 (Read-Only)**:
   - 중앙지사(`site_apps/central`)나 경인사무소(`site_apps/gyeongin`)의 작업일보를 불러올 때 원본 DB(`Material_Inventory.xlsx` 등)에 절대 쓰기 작업을 수행하지 않는다.
4. **도장 및 로고 위치 정밀도 (공식 표준 규격)**:
   - **직인(회사 도장) 표준 날인 기준**:
     * **상호명 겹침**: 상호(`서 울 檢 査 株 式 會 社`)의 우측 끝 한자(`株 式 會 社`) 위에 자연스럽게 겹치도록 날인한다 (공식 성적서 법적·미적 표준).
     * **경계선 준수**: 우측 `Customer / 발주처` 테두리 선이나 보고서 제목 세로 경계선을 절대 침범하지 않고 회사 정보 박스 안쪽에 정밀 안착한다.
     * **높이 및 비율**: 대략 4~7행(높이 약 64~72px) 범위 내에서 정사각형(1:1) 비율과 투명 PNG 배경을 엄격히 유지한다.
   - **검사방법별 적용 지침**:
     * **MT**: `mt_report_layout.py`의 `apply_header_stamp` (`col=15, row=3` 기반 OneCellAnchor) 기준 적용.
     * **RT**: `rt_report_layout.py`의 `header_stamp_anchor` (회사 정보와 제목 경계 기준 동적 정렬) 준수.
     * **PT / UT / PMI**: MT/RT와 동일한 헤더 레이아웃 규칙을 공통 적용하여 일관된 위치를 보장한다.
     * **PAUT**: 갑지 결재란 및 두 번째 시트 헤더 직인의 투명화 및 앵커 좌표를 보존한다.
   - 관련 기준 문서: `RT_STAMP_POSITION.md`, `home/config/logo_settings_unified.json`

## 검증 및 테스트 절차

수정 후에는 반드시 다음 검증 절차를 거친다:

1. **Python 가상환경 확인**:
   - `.venv/Scripts/python.exe`를 사용하여 실행한다.
2. **비파괴검사보고서 전용 회귀 테스트 실행**:
   ```powershell
   .\.venv\Scripts\python.exe -m unittest discover -s home/src/report_apps/ndt_report/tests -p "test_*.py"
   ```
   - 총 72개 이상의 테스트가 전부 `OK` 통과해야 한다.
3. **사업장 격리 검증 실행**:
   ```powershell
   .\.venv\Scripts\python.exe home/tests/test_site_isolation.py
   ```
4. **GUI 및 엑셀 파일 검증 시 주의**:
   - 테스트용 임시 엑셀 파일은 `.sync-backups/`나 `outputs/` 등 임시 경로를 사용하고, 운영 데이터 파일과 섞이지 않도록 한다.

## Git 형상 관리

- 수정 작업 및 사전 검증은 `비파괴검사보고서-개발` 브랜치에서 진행한다.
- 커밋 시 테스트로 생성된 임시 엑셀(`test_*.xlsx`, 임시 출력물)이나 로컬 로그는 제외하고 순수 소스 코드 및 설정 파일만 선별 스테이징한다.
- 안정화 및 사용자 승인 후 `비파괴검사보고서-안정`으로 안전하게 머지/푸시한다.

