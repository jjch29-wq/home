import pandas as pd
import json

file_path = r"C:\Users\jjch2\Desktop\산출내역서(가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역)_감독경유.xlsx"
out_path = r"C:\Users\jjch2\Desktop\PMI\_debug_scripts\excel_analysis.txt"

try:
    xl = pd.ExcelFile(file_path)
    sheet_names = xl.sheet_names
    
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(f"총 시트 개수: {len(sheet_names)}\n")
        f.write(f"시트 목록: {', '.join(sheet_names)}\n\n")
        
        for sheet in sheet_names:
            f.write(f"=== 시트: {sheet} ===\n")
            try:
                df = xl.parse(sheet, nrows=20)
                # print columns
                f.write(f"컬럼: {list(df.columns)}\n")
                # print first few rows as text table
                f.write(df.head(10).to_string() + "\n\n")
            except Exception as e:
                f.write(f"Error reading sheet: {e}\n\n")
            
    print("Done")
except Exception as e:
    print(f"Error: {e}")
