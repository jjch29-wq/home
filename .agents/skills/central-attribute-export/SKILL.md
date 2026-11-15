---
name: central-attribute-export
description: 한국지역난방 중앙지사 제출용 속성정보 및 검사내역 엑셀(중앙지사_속성정보.xlsx) 자동 생성 시스템의 코드·화면·매칭 로직·출력 기능을 유지보수할 때 사용한다. 과업내용서 별표3 기준 결함종류/판독코드 보존, 보고서·필름영상화·PAUT데이터 자동 연결, 원본 일보 읽기 전용 보호, 전 셀 가운데 정렬 규격 및 회귀 테스트를 준수한다.
---

# 중앙지사 속성정보 생성 시스템 유지보수

## 작업 시작

- 현재 작업 디렉터리와 저장소 루트를 확인한다 (`C:\Users\jjch2\Desktop\PMI`).
- 활성 브랜치가 `중앙지사_속성정보_생성` (또는 관련 개발 브랜치)인지 확인한다 (`git branch --show-current`).
- 원본 작업일보나 운영 엑셀 데이터를 절대 직접 수정하지 않는다 (Read-Only 원칙).
- 타 사업장(경인, 롯데 등) 코드를 혼합하거나 변경하지 않는다.

## 코드 위치와 경계

속성정보 및 검사내역 생성 앱의 핵심 코드는 중앙지사 독립 모듈로 관리된다.

- **속성정보 생성 메인 GUI 및 통합 컨트롤러**:
  - `home/src/site_apps/central/src/attribute_export.py`
- **과업내용서 별표3 계약 연동 및 파일 자동 연결 엔진**:
  - `home/src/site_apps/central/src/contract_links.py`
- **원본 데이터 소스 (Read-Only)**:
  - 중앙지사 일일작업기록: `home/src/site_apps/central/data/` (작업일보 히스토리 JSON 및 엑셀)
- **로컬 캐시 및 설정**:
  - 연결 정보 캐시: `.central_attribute_connections.json`
- **회귀 검증 테스트**:
  - `home/tests/test_central_attribute_export.py` (속성정보 내보내기/정렬/서식 검증)
  - `home/tests/test_contract_links.py` (결함종류 매핑, 보고서/영상/데이터 자동 연결 검증)
  - `home/tests/test_site_isolation.py` (사업장 격리 검증)

## 핵심 비즈니스 규칙 및 제출 데이터 표준

1. **원본 데이터 절대 보존 (Read-Only)**:
   - 중앙지사 일일작업기록 및 `Material_Inventory.xlsx`는 읽기 전용으로만 참조하며 파일 내용을 절대 변경하지 않는다.
2. **과업내용서 별표3 기준 결함종류 표기 규격**:
   - **RT 무결함**: 공란으로 두지 않고 반드시 **`N/1`**로 명시 표기한다.
   - **RT 판독 등급 보존**: `P/1`, `P/2`, `P/3`, `CL/4`, `UC/1` 등 원본 판독 등급 코드를 임의 변환 없이 그대로 보존한다.
   - **용접부특성 구분**:
     * 용접부번호(Joint No)가 `R`로 끝나면(예: `15R`, `02R`): **`REPAIR`**
     * 원본 검사이면: **`O`**
3. **제출 자료 3종(보고서·필름영상화·PAUT데이터) 자동 연결 무결성**:
   - **보고서 파일**: PDF/엑셀/워드/한글 보고서와 자동 매칭 및 실제 보고서 매수 자동 감지.
   - **RT 필름영상화**: 도면/라인 및 촬영포인트(`1-1`, `1-2` 등) 기준 매칭.
   - **PAUT 데이터**: 라인명 및 검사데이터 폴더 기준 매칭.
4. **엑셀 출력 서식 표준 (전 셀 중앙 정렬)**:
   - 생성되는 `속성정보` 시트와 `검사내역` 시트의 모든 데이터 셀은 **가로/세로 가운데 정렬**(`Alignment(horizontal='center', vertical='center')`)을 필수로 적용한다.
   - 필수 컬럼 규격:
     `공사명`, `시공사`, `하도사`, `용역사`, `도면번호`, `라인번호`, `용접부번호`, `용접부특성`, `검사일자`, `결함종류`, `보고서종류`, `보고서 매수`

## 검증 및 테스트 절차

수정 후에는 반드시 다음 검증 절차를 거친다:

1. **Python 가상환경 확인**:
   - `.venv/Scripts/python.exe`를 사용하여 실행한다.
2. **속성정보 및 계약 연동 단위 테스트 실행**:
   ```powershell
   .\.venv\Scripts\python.exe -m unittest home.tests.test_central_attribute_export home.tests.test_contract_links
   ```
   - 21개 이상의 테스트 항목이 전부 `OK` 통과해야 한다.
3. **사업장 격리 검증 실행**:
   ```powershell
   .\.venv\Scripts\python.exe home/tests/test_site_isolation.py
   ```

## Git 형상 관리

- 모든 작업은 `중앙지사_속성정보_생성` 브랜치에서 진행한다.
- 커밋 시 테스트로 생성된 임시 엑셀(`중앙지사_속성정보.xlsx`, 임시 파일)이나 캐시 파일은 제외하고 순수 코드 및 스킬 설정만 선별 스테이징한다.
- 커밋 메시지는 규칙에 맞게 작성한다 (예: `feat(central): ...`, `fix(central): ...`, `style(central): ...`).
