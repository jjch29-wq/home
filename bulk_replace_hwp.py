import win32com.client as win32
import os
import glob

# 치환할 딕셔너리 정의 (기존 텍스트 : 바꿀 태그)
replacements = {
    "롯데건설(주)": "{{발주처}}",
    "롯데바이오로직스 신축공사": "{{공사명}}",
    "고준호(010-5090-8871)": "{{발주처담당자}}",
    "김춘호(010-7763-3436)": "{{안전관리자}}",
    "Se-75 60Ci x 1EA": "{{사용선원}}",
    "인천광역시 연수구 송도동 418": "{{현장주소}}",
    "인천 연수구 송도동 418": "{{현장주소}}",
    "032-549-8457": "{{현장전화번호}}",
    "3인 1조 (1개조)": "{{작업조}}",
    "2025.04.03": "{{시작일}}",
    "2025.05.07": "{{종료일}}",
    "25.04.03": "{{시작일_단축}}",
    "25.05.07": "{{종료일_단축}}",
    "신 매립지": "{{주변특징}}",
    "800m": "{{민가거리}}",
    "야간 작업": "{{주야간작업}}",
    "Project PROVIEDENCE(K1) 건설공사현장": "{{공사명}}",
    "10km": "{{관공서거리}}",
    "인천 송도 경찰서": "{{경찰서이름}}",
    "인천 송도 소방서": "{{소방서이름}}",
    "롯데바이오로직스 Project PROVIDENCE 송도1공장": "{{공사명_상세}}",
    "114-81-16377": "{{사업자번호}}",
    "고준호": "{{담당자이름}}",
    "롯데건설 품질팀": "{{담당자부서}}",
    "010-5090-8871": "{{담당자휴대폰}}",
    "032-822-0988": "{{담당자사무실전화}}",
    "junhoko@lotte.net": "{{담당자이메일}}",
    "소래로 → 아암대로 → 인천신항대로": "{{운반경로}}",
    "소래로 -> 아암대로 -> 인천신항대로": "{{운반경로}}",
    "15km": "{{총이동거리}}",
    "30분": "{{소요시간}}"
}

# 대상 파일 지정 (단일 파일)
target_file = r"C:\Users\jjch2\Desktop\신고서_마스터템플릿\템플릿 신고서류.hwp"

try:
    print("HWP 프로그램을 실행합니다...")
    # gencache.EnsureDispatch를 사용하여 Early Binding 강제 적용
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    hwp.XHwpWindows.Item(0).Visible = True # 화면에 보이게 설정
    
    # 단일 파일 리스트
    hwp_files = [target_file] if os.path.exists(target_file) else []
    
    if not hwp_files:
        print("경고: 폴더 내에 HWP 파일이 없습니다!")
    
    for file_path in hwp_files:
        print(f"처리 중: {os.path.basename(file_path)}")
        hwp.Open(file_path, "HWP", "forceopen:true")
        
        for old_text, new_text in replacements.items():
            hwp.HAction.Run("MoveDocBegin") # 문서 맨 앞으로 커서 이동
            hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
            hwp.HParameterSet.HFindReplace.FindString = old_text
            hwp.HParameterSet.HFindReplace.ReplaceString = new_text
            hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
            hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
            
        hwp.Save()
        hwp.Clear(1) # 문서 닫기
    
    hwp.Quit()
    print("모든 파일의 치환이 완료되었습니다!")
except Exception as e:
    print(f"에러 발생: {e}")
