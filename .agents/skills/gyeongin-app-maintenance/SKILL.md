---
name: gyeongin-app-maintenance
description: 경인사무소 자재 관리 및 모바일 현장 입력 시스템의 코드·화면·저장·조회·출력 기능을 유지보수할 때 사용한다. 경인사무소 사업장 경계(site_apps/gyeongin)와 전용 데이터베이스를 확인하고 안전하게 검증한다. 다른 사업장만의 작업이나 일반적인 비파괴검사 질문에는 적용하지 않는다.
---

# 경인사무소 앱 및 모바일 시스템 유지보수

## 작업 시작
- 현재 작업 디렉터리와 저장소 루트를 확인한다 (`C:/Users/jjch2/Desktop/PMI`).
- 적용 가능한 규칙인 `home/src/site_apps/AGENTS.md`의 사업장 격리 규칙을 준수한다.
- 활성 브랜치가 `경인사무소-개발`인지 확인한다.
- 경인사무소 작업 시 중앙지사(`site_apps/central`)나 타 사업장 코드를 혼합하거나 직접 수정하지 않는다.

## 코드 위치와 경계
경인사무소 앱의 모든 파일은 `home/src/site_apps/gyeongin/`에 독립 보관된다.
- PC 실행기: `home/src/경인사무소.py`
- 사업장 메인: `home/src/site_apps/gyeongin/main.py`
- 모바일 서버: `home/src/site_apps/gyeongin/mobile/server.py`
- 모바일 화면: `home/src/site_apps/gyeongin/mobile/static/index.html`
- 모바일 및 터널 실행기: `home/src/site_apps/gyeongin/run_mobile.py` (포트 8001)
- 설정 파일: `home/src/site_apps/gyeongin/src/Material_Manager_Config.json`
- 로컬 DB: `home/src/site_apps/gyeongin/data/Material_Inventory.xlsx`
- 구글 드라이브 클라우드 DB: `G:\내 드라이브\07_Antigravity\PMI_Data\gyeongin\data\Material_Inventory.xlsx`
- 양식 및 리소스: `home/src/site_apps/gyeongin/resources/`

## 업무 및 비즈니스 규칙
- 현장명(Site): `경인사무소`
- 업체명(Company): `경인사무소`
- OT 수당 단가 규칙:
  - 주간 연장 (18:00~22:00): 4,000원/h
  - 야간 연장 (22:00~24:00): 5,000원/h
  - 주말 및 휴일/금요일 자정 이후 익일 심야: 7,500원/h
- 검사방법별 기본 장비/보조장비 매핑:
  - RT: 주 장비 `Se-75`, 보조장비 `ADR, SURVEY, 콜리메타`
  - PAUT: 주 장비 `PAUT MX2(OMNI2-103085)`, 보조장비 `PAUT SCANNER (MANUAL)`
  - MT: 주 장비 `Yoke`, 보조장비 `자분살포기` (옵션: `조도계`, `자계강도계`, `휴대용 UV메타`)
  - PT: 주 장비 `PT 약품`, 보조장비 `조도계` (옵션: `온습도계`)
  - 보조장비 직접 입력: `✏️ 직접 입력...` 지원

## 네트워크 및 외부 접속 (LTE/5G)
- 로컬 포트: `8001` (중앙지사 8000과 충돌 없이 동시 실행 가능)
- Cloudflare Tunnel(`tools/cloudflared.exe`)을 통해 외부 스마트폰(LTE/5G)용 전용 HTTPS URL 자동 발행.

## Git 및 버전 관리
- 경인사무소 관련 모든 작업 커밋 및 푸시는 `경인사무소-개발` 브랜치를 사용한다.
- 중앙지사나 타 브랜치와 격리 상태를 유지한다.
