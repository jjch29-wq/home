# -*- coding: utf-8 -*-
"""
================================================================================
가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역 과업내용서 요약 프로그램
================================================================================
이 프로그램은 가산~가평 천연가스 공급시설 비파괴검사 기술용역 과업내용서의
핵심 사항을 요약하고, 터미널 환경에서 직관적으로 조회할 수 있는 CLI 뷰어입니다.
"""

import sys
import os

# Windows 터미널에서 ANSI escape 코드 및 UTF-8 출력을 지원하도록 설정
if sys.platform.startswith('win'):
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        # ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
        # ENABLE_PROCESSED_OUTPUT = 0x0001
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
    except Exception:
        pass
    
    try:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except Exception:
        pass

class Colors:
    HEADER = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def print_title(title):
    print(f"\n{Colors.BOLD}{Colors.HEADER}■ {title}{Colors.ENDC}")
    print(f"{Colors.HEADER}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.ENDC}")

def print_subtitle(sub):
    print(f"\n{Colors.BOLD}{Colors.BLUE}▶ {sub}{Colors.ENDC}")

def show_overview():
    print_title("1. 과업 개요 (Project Overview)")
    print(f"  * {Colors.BOLD}용역명{Colors.ENDC} : 가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역")
    print(f"  * {Colors.BOLD}발주처{Colors.ENDC} : 한국가스공사 건설사업단")
    print(f"  * {Colors.BOLD}목  적{Colors.ENDC} : 용접부 비파괴검사를 통한 결함 제거 및 설비 품질·안전성 확보")
    
    print_subtitle("검사 방법 및 범위 (100% 검사 원칙)")
    print("  ┌───────┬──────────────────────────────────────────┬────────────────────────┐")
    print("  │  구분 │                 검사 범위                │        시행 범위       │")
    print("  ├───────┼──────────────────────────────────────────┼────────────────────────┤")
    print("  │  RT   │ 맞대기 용접부(가배관 포함), RT가능 개소   │ 100%                   │")
    print("  │  (방사│ 소구경(2\" 이하) 맞대기 및 2\" 이상 필렛   │ (현장 여건 따라 PT 가능)│")
    print("  ├───────┼──────────────────────────────────────────┼────────────────────────┤")
    print("  │  UT   │ 내압·기밀시험 불가 맞대기, 특수구간(하천,│ 100%                   │")
    print("  │  (초음│ 철도, 교량 등) 맞대기, 필렛 용접부       │ 발주자 지정개소        │")
    print("  ├───────┼──────────────────────────────────────────┼────────────────────────┤")
    print("  │  PT   │ 내압·기밀시험 불가 맞대기, 특수구간      │ 100% / 지정개소        │")
    print("  └───────┴──────────────────────────────────────────┴────────────────────────┘")

    print_subtitle("용역 예상 물량 (실검사 물량 기준)")
    print(f"  * {Colors.BOLD}방사선투과검사 (RT) 총 24,536매{Colors.ENDC}")
    print("    - 3 ⅓\" × 17\" : 20,368매 (배관 20\" × 34.65㎞, 3개 관리소 포함)")
    print("      ※ 관리소: 직두BV, 원흥VS, 가평GS")
    print("    - 3 ⅓\" × 12\" : 2,464매")
    print("    - 3 ⅓\" × 6\"  : 1,704매")
    print(f"  * {Colors.BOLD}초음파탐상검사 (UT){Colors.ENDC} : 실검사 길이 319.02 M")
    print(f"  * {Colors.BOLD}액체침투탐상검사 (PT){Colors.ENDC} : 실검사 길이 338.63 M")

def show_personnel():
    print_title("2. 검사원 자격 및 관리 기준")
    print_subtitle("투입 인력 자격 요건")
    print("  ┌──────────┬────────────────────────────────────────────────────────┐")
    print("  │   구분   │                        자격 요건                       │")
    print("  ├──────────┼────────────────────────────────────────────────────────┤")
    print("  │현장대리인│ - 비파괴검사 기술사                                    │")
    print("  │          │ - 기사(또는 ASNT L3) 취득 후 경력 5년 이상             │")
    print("  │          │ - 산업기사 취득 후 경력 8년 이상                       │")
    print("  ├──────────┼────────────────────────────────────────────────────────┤")
    print("  │ 판 독 자 │ - 분야별 기사(ASNT L3 포함) 취득 후 현장 경력 3년 이상 │")
    print("  ├──────────┼────────────────────────────────────────────────────────┤")
    print("  │ 검 사 자 │ - 분야별 기능사(ASNT L3 포함) 취득 후 현장 경력 1년이상│")
    print("  │          │   (RT의 경우 RT면허 또는 작업조장 교육 이수 필수)      │")
    print("  ├──────────┼────────────────────────────────────────────────────────┤")
    print("  │ 검사보조 │ - 분야별 기능사(ASNT L2 포함) 또는 동등 이상 자격증    │")
    print("  └──────────┴────────────────────────────────────────────────────────┘")
    
    print_subtitle("인력 관리 핵심 수칙")
    print(f"  1. {Colors.WARNING}현장 상주 의무{Colors.ENDC}: 현장대리인, 안전관리자, 판독자는 공사 중 상주 필수.")
    print(f"     ※ {Colors.FAIL}겸임 금지{Colors.ENDC}: 현장대리인과 방사선안전관리자는 겸임할 수 없음.")
    print("  2. 인력 변경 승인: 인원 교체 및 철수 시 발주자 용역감독원 승인 필수.")
    print("     - 제출 서류: 이력서(사진부착), 경력증명서, 자격증 사본.")
    print("  3. 업무 정지 권한: 감독원은 서류 이상이나 수행능력 부족 판정 시 투입 정지 가능.")
    print("     - 계약상대자는 즉시 대체인원을 투입해야 함.")

def show_technical_standards():
    print_title("3. 기술 기준 및 촬영 규정")
    print_subtitle("관련 규격 및 기준")
    print("  - KS B 0888 (배관 용접부 비파괴시험 방법)")
    print("  - KS B 0845 (강 용접부 RT 시험 및 등급 분류)")
    print("  - KS B 0896 (강 용접부 UT 시험 및 등급 분류)")
    print("  - KS B 0816 (PT 시험방법 및 결함 분류)")
    print("  - 도시가스사업법 KGS Code GC 205 (가스시설 용접 및 비파괴)")
    print("  - KOGAS 표준: KOGAS-GSD-2130 (용접 기술표준), KOGAS-GSD-0102 (선정기준)")

    print_subtitle("RT 합격 기준 (방사선 투과검사)")
    print("  ┌──────────────────────────────┬─────────────┬─────────────────────┐")
    print("  │             구분             │  합격 등급  │         비고        │")
    print("  ├──────────────────────────────┼─────────────┼─────────────────────┤")
    print("  │ 내압시험 실시 일반/Tie-in    │     2급     │ KGS CODE 기준       │")
    print("  │ 내압시험 불가 Tie-in         │     1급     │ 기체 내압 임시배관  │")
    print("  ├──────────────────────────────┼─────────────┼─────────────────────┤")
    print("  │ 내압시험 생략 일반/Tie-in    │     2급     │ 공급관리소 해당     │")
    print("  │ 내압시험 생략 불가 Tie-in    │     1급     │ 기밀시험 불가 개소  │")
    print("  └──────────────────────────────┴─────────────┴─────────────────────┘")
    print("  ※ 승압/임시 공급 설비: 기본 2급 적용 (기밀시험 불가 시 1급 적용)")

    print_subtitle("촬영 방법 및 기준")
    print("  1. 필름 요건: ASTM E94 TYPE I 이상 사용, 필름 농도 1.5 ~ 3.5 유지.")
    print("  2. DWSI(이중벽 단상촬영): 1회 1장 촬영 원칙 (26\" 이상 배관은 1회 2장 촬영 가능).")
    print("  3. 내부선원법: Crawler 등 장비 투입 가능 시 우선 투입하여 촬영.")
    print("  4. 배관 구경별 촬영 매수 기준:")
    print("     - 30\"~32\": 7매 (이중벽/단벽 단상)    - 10\"~12\": 4매 (이중벽 단상)")
    print("     - 26\": 6매 (이중벽/단벽 단상)        - 6\"~8\": 3매 (이중벽 단상)")
    print("     - 20\"~24\": 5매 (이중벽/단벽 단상)    - 3\"~4\": 3매 (이중벽 단상)")
    print("     - 18\": 6매 (이중벽 단상)             - 2½\" 이하: DWDI 2매 또는 DWSI 3매")
    print("     - 14\"~16\": 5매 (이중벽 단상)")

def show_safety():
    print_title("4. 방사선 안전관리 및 작업 규정")
    print_subtitle("원자력 안전 및 방호 대책")
    print(f"  1. {Colors.WARNING}인허가 이행{Colors.ENDC}: 방사성동위원소 이동사용·운반·저장에 대해 원자력안전법 허가 필득.")
    print("     - KINS(한국원자력안전기술원) 개설/변경/폐지 신고 승인 문서 제출 필수.")
    print("  2. 차폐 기준: 차폐체(공사 제공 납일체형 외 보조 차폐체는 수급사 조달) 사용 의무.")
    print(f"  3. 선원 제한: {Colors.BOLD}Ir-192 20Curi 이하{Colors.ENDC}, {Colors.BOLD}Se-75 60Curi 이하{Colors.ENDC} 동위원소만 사용.")
    print(f"  4. 방사선량 제한: 작업자 외부 선량 {Colors.WARNING}시간당 10μSv 이하{Colors.ENDC} 유지.")
    print("     - 시간당 1μSv 초과 구역은 일반인 접근 감시 및 통제선 설치 필수.")
    print(f"  5. {Colors.FAIL}배관 내부 진입 절대 금지{Colors.ENDC}: 내부선원 Crawler 촬영 시 인원 진입 금지 (적발 시 퇴출).")
    print("     - 부득이한 진입 시 밀폐공간작업 프로그램(KOGAS-SI-102) 수립 및 감독원 승인 필요.")
    print("  6. 폐기물 처리: 현상액, 정착액 등 액상 폐기물은 관련법에 따라 적법 처리.")

    print_subtitle("휴일 및 야간 작업 규정 (22:00 ~ 익일 06:00)")
    print("  * 시행 사유: 공기단축, 휴일 Tie-in, 하천/압입 집중 작업, 도로오픈 횡단, KGS 시공감리 등.")
    print("  * 절차: 시공자 신청 -> 시공감독원 검토 -> 용역감독원 승인 및 요청 -> 계약상대자 시행 -> 결과 보고.")

def show_health_safety_buddy():
    print_title("5. 현장 안전보건, 교육 및 [2인 1조] 의무 작업")
    print_subtitle("안전보건 의무 사항")
    print("  - KOGAS 안전관리계획서(KOGAS-SI-110 기준) 및 매월 안전활동보고서 제출.")
    print("  - 환경보전관리교육 및 안전·보건교육 실시 의무.")
    print("  - 근로자 위험성평가(최초, 정기, 수시) 교육 및 이행.")
    print("  - 안전보건협의체 운영(매월 1회) 및 합동 안전점검(분기 1회 이상).")
    print("  - TBM 및 10대 기본안전수칙 위반 시 작업 중지 권한 발동.")
    print(f"  - {Colors.BOLD}근로자 작업중지권 보장{Colors.ENDC} 및 이를 이유로 한 불이익 처분 금지.")

    print_subtitle("2인 1조 의무 수행 작업 (사고 예방)")
    print(f"  1. {Colors.WARNING}관리소 내 화기 작업{Colors.ENDC} (용접, 절단, 천공, Tie-in 등)")
    print(f"  2. {Colors.WARNING}지상 기준 2m 이상{Colors.ENDC} 고소 작업")
    print(f"  3. {Colors.WARNING}배관 내부 등 밀폐공간{Colors.ENDC} 작업")
    print("  4. 화학물질 취급 (염소, 가성소다, 차아염소산 나트륨, 산/염기 등)")
    print("  5. 중량물 이동 작업")
    print("  6. 분진·비산, 화재·폭발 위험 작업")
    print(f"  7. {Colors.BOLD}근속 기간 6개월 미만{Colors.ENDC} 근로자 단독 작업 금지")
    print("  8. 기타 발주자, 계약상대자, 근로자가 필요하다고 판단하는 작업")

    print_subtitle("환경보전관리 및 안전·보건 교육 기준")
    print(f"  {Colors.BOLD}[1] 환경보전관리교육 (년 1회 / 수시){Colors.ENDC}")
    print("    - 교육명 : 환경관리교육")
    print("    - 내  용 : 1. 생활쓰레기 및 일반 폐기물 처리방법")
    print("              2. 현상액·정착액의 수질관리 및 보관")
    print("              3. 폐기물(방사성동위원소) 관리법")
    print("    - 근  거 : 환경보전관리에 관한 법령")
    print(f"  {Colors.BOLD}[2] 안전·보건교육{Colors.ENDC}")
    print("    - 일반 산업안전 및 보건 교육 : 매월 1시간 이상 (신규 8시간)")
    print("      * 내  용 : 산업안전보건 관계법령, 작업 환경 특성 따른 위험성, 표준 작업방법,")
    print("                보호구/안전장구 사용법, 안전사고사례 및 예방 대책 (밀폐공간, 내부선원법 작업)")
    print("      * 근  거 : 산업안전보건법 시행규칙 별표8.2")
    print("    - 방사선 안전교육 : 매월 1시간 이상 (신규 18시간)")
    print("      * 내  용 : 방사선이 인체에 미치는 영향, 방사성동위원소/발생장치 안전 취급, 방사선장해방지/규정")
    print("      * 근  거 : 원자력안전법 시행령 제148조 및 시행규칙 제138조")
    print("    - 방화관리 교육 : 수시 (신규 1시간)")
    print("      * 내  용 : 소방관계법령, 소방시설 종류/작동원리, 소화기 사용법, 발화/인화성 물질 관리")
    print("      * 근  거 : 소방법시행규칙 제82조")
    print("    - 안전재해 예방교육 : 작업 시작 전 10분")
    print("      * 내  용 : 낙하물/추락주의, 장비 이동시 주의, 방사선 개인 안전 장구류 착용 등")
    print("    - 전입직원교육 (신입포함) OJT : 최초 8시간")
    print("      * 내  용 : 산업안전보건교육, 안전수칙, 현장작업사항 및 현장교육")
    print("      * 근  거 : EHSQ 지침서, 용역수행계획서")

    print_subtitle("기타 기후 및 상벌 수칙")
    print(f"  * {Colors.WARNING}폭염 대책{Colors.ENDC}: 체감온도 35℃ 이상 시 무더위 시간대(14~17시) 옥외작업 중지.")
    print("    - 매시간 15분 이상 그늘(휴식공간) 제공 의무화.")
    print("  * 안전경고장 누적 조치:")
    print("    - 1회/년 발부 : 근로자 특별안전교육 시행 및 결과 보고")
    print("    - 2회/월 또는 2회/년 발부 : 현장 내 출입정지")

def show_submittals_finance():
    print_title("6. 제출 서류 및 정산 기준")
    print_subtitle("용역 제출 서류 (11종)")
    print("  1. 계약서 (계약 후 10일 이내, 3부)       2. 착수계 (계약 후 20일 이내, 1부)")
    print("  3. 검사용역수행계획서 (20일 이내, 3부)    4. 작업장 개설/변경 신고서 (승인 후 즉시, 1부)")
    print("  5. 월간 용역진도보고서 (매월 7일, 1부)    6. 종합용역진도보고서 (준공시, 2부)")
    print("  7. 비파괴검사 보고서 (검사 후 즉시, 2부)  8. 비파괴검사 일보 (매일, 1부)")
    print("  9. 작업장 폐지 신고서 (준공시, 1부)      10. 사진첩 (기성/준공시, 1~3부)")
    print("  11. 비파괴검사 기록 외장하드/CD (준공시, 3식)")

    print_subtitle("기록 및 비용 정산 기준")
    print("  1. 실적 정산 원칙: 검사물량(RT, 일반, 야간, 휴일) 및 용역기간은 실제 수행 기준 정산.")
    print("  2. 안전관리비 실적 정산:")
    print("     - 대상: 각종 안전표지, 경고등, 저장시설, 차폐벽 제작, TLD 판독료, 특수검진비 등.")
    print("     - 제외: 개인보호구(안전모, 안전화 등), 업무용 기기, 일반 피복 등.")
    print("  3. 기타 정산 항목: 주재비/출장비, 기계경비(장비투입기간), 가설사무실, 손해배상보험료.")
    print(f"  4. {Colors.FAIL}보수 및 재촬영 비용 제한 (3% 룰){Colors.ENDC}")
    print(f"     - 결함 보수 및 시공사 과실로 인한 재촬영 방사선투과검사비는 {Colors.BOLD}최대 3%(필름 매수 기준){Colors.ENDC}만 인정.")
    print("     - 3% 초과분에 대해서는 시공사 부담으로 처리하며, 기성 청구 시 용접불량률 제출 필수.")

def main():
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print(f"{Colors.BOLD}{Colors.GREEN}===================================================================={Colors.ENDC}")
        print(f"{Colors.BOLD}{Colors.GREEN}    가산~가평 천연가스 공급시설 NDT 기술용역 과업내용서 요약 뷰어   {Colors.ENDC}")
        print(f"{Colors.BOLD}{Colors.GREEN}===================================================================={Colors.ENDC}")
        print("  [1] 과업 개요 (용역 목적, 검사 범위 및 예상 물량)")
        print("  [2] 검사원 자격 및 인력 관리 기준 (현장 상주, 겸임 금지 등)")
        print("  [3] 기술 기준 및 촬영 규정 (RT 합격등급, 구경별 필름 매수)")
        print("  [4] 방사선 안전관리 및 야간/휴일작업 규정")
        print("  [5] 현장 안전보건, 교육 및 [2인 1조] 의무 작업 수칙")
        print("  [6] 제출 서류 및 정산 기준 (3% 재촬영 제한 등)")
        print("  [7] 전체 요약 보기")
        print("  [0] 종료")
        print(f"{Colors.GREEN}────────────────────────────────────────────────────────────────────{Colors.ENDC}")
        
        try:
            choice = input("  조회할 메뉴 번호를 입력하세요: ").strip()
        except KeyboardInterrupt:
            print("\n  프로그램을 종료합니다.")
            break
            
        if choice == '1':
            show_overview()
        elif choice == '2':
            show_personnel()
        elif choice == '3':
            show_technical_standards()
        elif choice == '4':
            show_safety()
        elif choice == '5':
            show_health_safety_buddy()
        elif choice == '6':
            show_submittals_finance()
        elif choice == '7':
            show_overview()
            show_personnel()
            show_technical_standards()
            show_safety()
            show_health_safety_buddy()
            show_submittals_finance()
        elif choice == '0':
            print("  프로그램을 종료합니다.")
            break
        else:
            print(f"\n{Colors.FAIL}  올바르지 않은 입력입니다. 다시 선택해주세요.{Colors.ENDC}")
            
        input(f"\n{Colors.CYAN}  [Enter] 키를 누르면 주 메뉴로 돌아갑니다...{Colors.ENDC}")

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--all':
        show_overview()
        show_personnel()
        show_technical_standards()
        show_safety()
        show_health_safety_buddy()
        show_submittals_finance()
    else:
        main()
