---
name: gyeongin-app-maintenance
description: 경인사무소 자재 관리 및 모바일 현장 입력 시스템의 코드·화면·저장·조회·출력 기능을 유지보수할 때 사용한다. 경인사무소 사업장 경계(site_apps/gyeongin)와 전용 데이터베이스를 확인하고 안전하게 검증한다. 다른 사업장만의 작업이나 일반적인 비파괴검사 질문에는 적용하지 않는다.
---

# 경인사무소 앱 및 모바일 시스템 유지보수

## 작업 시작
- 현재 작업 디렉터리와 저장소 루트를 확인한다 (`c:\Users\-\PMI`).
- 적용 가능한 규칙인 `home/src/site_apps/AGENTS.md`의 사업장 격리 규칙을 준수한다.
- 활성 브랜치가 `경인사무소-개발`인지 확인한다.
- 경인사무소 작업 시 중앙지사(`site_apps/central`)나 타 사업장 코드를 혼합하거나 직접 수정하지 않는다.

## 코드 위치와 경계
경인사무소 앱의 모든 파일은 `home/src/site_apps/gyeongin/`에 독립 보관된다.
- PC 실행기: `home/src/경인사무소.py`, `home/src/site_apps/gyeongin/main.py`
- PC 코어 및 동기화:
  - `home/src/site_apps/gyeongin/src/app.py`
  - `home/src/site_apps/gyeongin/src/sync_google_sheet.py` (구글 시트 양방향 동기화)
  - `home/src/site_apps/gyeongin/src/sync_gas_assets.py` (PC 마스터 → GAS 상수 단일 진실 원천 동기화)
- **주력 모바일 (24시간 Google Apps Script 웹 앱)**:
  - `home/src/site_apps/gyeongin/google_apps_script/Code.gs`
  - `home/src/site_apps/gyeongin/google_apps_script/index.html`
  - 신규 모바일 기능·UI 수정은 GAS를 우선 기준으로 반영하고, 수정 완료 시 사용자에게 Apps Script 편집기 재배포가 필요함을 안내한다.
- **로컬 백업 모바일 (FastAPI + Cloudflare, 포트 8001)**:
  - `home/src/site_apps/gyeongin/mobile/server.py`, `mobile/static/index.html`, `run_mobile.py`
  - 사내 PC 긴급 오프라인/터널 백업 용도로 유지한다 (별도 요청 시에만 동기 변경).
- 설정 및 마스터 파일:
  - `home/src/site_apps/gyeongin/src/Material_Manager_Config.json`
  - `home/src/site_apps/gyeongin/resources/company_ndt_presets.json`
  - `home/src/site_apps/gyeongin/src/aux_equipment.py`, `company_master.py`
- 데이터베이스 (다중 경로 자동 탐색):
  - 로컬 DB: `home/src/site_apps/gyeongin/data/Material_Inventory.xlsx`
  - Google Drive DB: `[A-Z]:\내 드라이브\07_Antigravity\PMI_Data\gyeongin\data\Material_Inventory.xlsx` (예: `F:\`, `G:\`)
  - 동기화 설정: `home/src/site_apps/gyeongin/data/google_sync_config.json`
- 양식 및 리소스: `home/src/site_apps/gyeongin/resources/`

## 업무 및 비즈니스 규칙
- 기본 현장명(Site) / 업체명(Company): `경인사무소`
- OT 수당 단가 규칙:
  - 주간 연장 (18:00~22:00): 4,000원/h
  - 야간 연장 (22:00~24:00): 5,000원/h
  - 주말 및 휴일/금요일 자정 이후 익일 심야: 7,500원/h
- 검사방법별 기본 장비/보조장비 매핑 (`aux_equipment.py` 및 `Material_Manager_Config.json` 기준):
  - RT: 주 장비 `Se-75(880D D10692)` 등, 보조장비 `ADR, SURVEY, 콜리메타`
  - PAUT: 주 장비 `PAUT MX2(OMNI2-103085)` 등, 보조장비 `PAUT SCANNER (MANUAL)`
  - MT: 주 장비 `Yoke AC(MY-2 NW0453)` 등, 보조장비 `자분살포기` 등
  - PT: 주 장비 `PT 약품`, 보조장비 `조도계 LUX Meter 1330A (180604411)` 등
  - 주 장비·보조장비·자재명(`; ` 구분 결합) 복수 선택 및 직접 입력 지원
- 업체별·검사방법별 검사품명 프리셋(`company_ndt_presets.json`):
  - 해당 업체에 등록된 품목만 우선 표시해 타 업체 품목 혼입을 방지한다.
- NDT 상세 조건 및 물량 산출 규칙:
  - 지역난방 전용 명칭('열수송관', '핫태핑') 및 제경비(110%)/기술료(20%) 단가 산출식은 숨김 처리하고 기본 구분은 `경인사무소`로 고정한다.
  - 관경에 따른 조인트/물량 산출(`[ORI] 조인트/물량`, `[REP] 조인트/물량`) 및 당일 불량(`REJ`) 매수 계산 기능은 유지한다.

## 마스터 동기화 및 검증 절차
- 업체·사용자·차량·장비·자재·프리셋을 수정한 경우 반드시 아래 명령으로 GAS 마스터를 동기화한다:
  - `PYTHONPATH=home/src .\.venv\Scripts\python.exe -m site_apps.gyeongin.src.sync_gas_assets`
- 작업 완료 후 반드시 아래 검증을 실행한다:
  - 격리 검증: `.\.venv\Scripts\python.exe home/tests/test_site_isolation.py`
  - GAS 마스터 일치 검증: `PYTHONPATH=home/src .\.venv\Scripts\python.exe -m site_apps.gyeongin.src.sync_gas_assets --check`
  - 경인 단위 테스트: `.\.venv\Scripts\python.exe -m unittest discover -s home/tests -p "test_gyeongin_*.py"`
- 주의: 테스트 작성 시 실제 Google Apps Script 웹 앱으로 더미 레코드가 Push되거나 로컬/Drive DB가 오염되지 않도록 반드시 임시 파일(`tempfile`)과 mock을 사용한다.

## Git 및 버전 관리
- 경인사무소 관련 모든 작업 커밋 및 푸시는 `경인사무소-개발` 브랜치를 사용한다.
- 실데이터/로그(`Material_Inventory.xlsx`, `google_sync_config.json`, `daily_work_history.json`, `debug_save.txt`)는 별도 요청이 없는 한 커밋에 포함하지 않는다.
- 중앙지사나 타 브랜치와 격리 상태를 유지한다.

