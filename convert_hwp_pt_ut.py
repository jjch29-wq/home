import win32com.client as win32
import os
import time

def replace_text(hwp, old_text, new_text):
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = old_text
    hwp.HParameterSet.HFindReplace.ReplaceString = new_text
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

def main():
    filepath = r"C:\Users\-\OneDrive\바탕 화면\3. 유해 위험요인 조사표(사업장 순회점검, 청취조사, 안전보건자료, 체크리스트).hwp"
    
    if not os.path.exists(filepath):
        print(f"파일을 찾을 수 없습니다: {filepath}")
        return

    print("한글(HWP) 프로그램을 실행합니다...")
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    except Exception as e:
        print("한글 프로그램 실행 실패:", e)
        return

    # 보안 모듈 승인 창이 뜰 수 있으므로 화면에 보이게 함
    hwp.XHwpWindows.Item(0).Visible = True
    
    try:
        # 보안 모듈 등록 시도 (실패해도 무시)
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    except:
        pass

    print(f"파일을 엽니다: {filepath}")
    print("⚠️ 주의: 한글 프로그램에서 '접근 허용' 관련 보안 팝업이 뜨면 [허용]을 눌러주세요!")
    time.sleep(2) # 사용자가 볼 수 있도록 잠시 대기
    
    hwp.Open(filepath)
    
    print("방사선(RT) 위험요인을 PT/UT 위험요인으로 변환 중...")

    # 변환할 텍스트 매핑 (핵심 키워드로 검색하여 전체 문장이나 해당 키워드 교체)
    # 기존 방사선 관련 문구가 문서에 어떻게 적혀있을지 모르니, 핵심 단어로 변경합니다.
    replacements = {
        # 3. 가이딩 찍힘
        "가이딩 찍힘, 꺽임 등으로 방사선원의 이탈 및 회수가 되지 않는 사고 위험(최대)": "밀폐된 공간이나 환기가 불충분한 곳에서 세척액 및 침투액 사용 시 유기용제 증기 흡입에 의한 중독 위험(최대)",
        # 4. Pig Tail
        "방사선조사기 Pig Tail 유격으로 인한 방사선원 이탈 사고(최대)": "인화성 에어로졸(세척액, 현상액) 취급 중 주변 용접 불꽃 등 화기 접촉으로 인한 화재 및 폭발 위험(최대)",
        # 5. 방사선관리구역 미설정
        "방사선관리구역 미설정 시 일반인 미통제로 피폭사고 위험(최대)": "침투액 등 화학물질 취급 시 보호구 미착용으로 인한 피부 접촉 및 피부염 발생 위험(중)",
        # 6. 개인안전장구 미착용
        "방사선 개인안전장구 미착용으로 방사선 노출시 방사선 피폭을 인지하지 못하여 과피폭 위험(최대)": "고소작업(비계, 사다리 등) 중 초음파탐상검사 시 안전대 미체결 및 부주의에 의한 추락 위험(최대)",
        # 7. 서베이메타
        "교정되지 않은 서베이메타 사용시 실제 방사선량 측정을 할 수 없을시 피폭사고 위험(최대)": "좁고 불편한 자세로 장시간 탐촉자를 문지르는 반복작업에 의한 근골격계 질환 위험(중)",
        # 8. 콜리메타
        "콜리메타 미사용 또는 부적절한 콜리메타 사용 시 피폭사고 위험(최대)": "접촉매질(Couplant, 겔 등)이 바닥에 흘러 작업자가 밟고 미끄러짐(넘어짐) 위험(중)",
        # 9. 혼재작업 피폭
        "통제되지 않은 타공정과 혼재작업 시 일반인 피폭사고 위험(최대)": "탐상장비 전원 케이블 피복 손상 또는 습윤한 환경에서 작업 시 감전 위험(최대)",
        # 10. 낙하로 장비파손
        "무거운 방사선장비 이동 및 사용시 방사선장비의 낙하로 장비 파손으로 인한 방사선사고 위험(최대)": "통제되지 않은 타공정과 혼재작업 시 중량물 낙하 등에 의한 맞음 사고 위험(최대)",
        # 12. 어두운 환경
        "야간작업 시, 어두운 환경에서 이동 및 작업할 경우 넘어짐 또는 방사선사고 발생 위험(최대)": "야간작업 시, 어두운 환경에서 이동 및 작업할 경우 시야 미확보에 의한 넘어짐 및 부딪힘 위험(최대)",
        # 13. 필름현상
        "필름현상 시, MSDS물질을 취급함에 따라 흡입 위험(중)": "PT 폐기물(사용한 걸레, 빈 캔)의 무단 방치 및 화기 근접으로 인한 화재 위험(중)"
    }

    for old_t, new_t in replacements.items():
        replace_text(hwp, old_t, new_t)
        
    # 사고 유형의 '방사선피폭' 텍스트 변경
    replace_text(hwp, "기타(방사선피폭)", "기타(PT/UT위험)")

    # 다른 이름으로 저장
    save_path = filepath.replace(".hwp", "_PT_UT_수정본.hwp")
    hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
    hwp.HParameterSet.HFileOpenSave.filename = save_path
    hwp.HParameterSet.HFileOpenSave.Format = "HWP"
    hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

    print(f"\n변환이 완료되었습니다! 파일이 저장되었습니다:\n{save_path}")
    
    # 종료 여부 묻기 (사용자가 직접 확인 후 종료할 수 있도록)
    # hwp.Quit()

if __name__ == "__main__":
    main()
