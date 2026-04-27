import pandas as pd
import sys

try:
    file_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xls'
    df = pd.read_excel(file_path, header=None)
    
    keywords = ["업체명", "공사명", "적용규격", "사용장비명", "성적서번호", "검사품명", "검사자", "차량번호", 
                "검사방법", "단위", "검사량", "단가", "출장비", "합계",
                "센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타", "총계",
                "O/T시간", "O/T금액",
                "반출", "이월사용량", "임고", "재고"]
    
    print("KEYWORD LOCATIONS:")
    for r_idx, row in df.iterrows():
        for c_idx, val in enumerate(row):
            s_val = str(val)
            if any(k in s_val for k in keywords):
                print(f"Row {r_idx}, Col {c_idx} (Cell {chr(65+c_idx)}{r_idx+1}): {val}")
except Exception as e:
    print(f"Error: {e}")
