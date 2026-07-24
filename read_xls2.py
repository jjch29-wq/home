import pandas as pd

try:
    xl = pd.ExcelFile(r'C:\Users\-\OneDrive\바탕 화면\04.인원 및 장비투입계획서.xls')
    out = ""
    for s in xl.sheet_names:
        out += f"\n\n{'='*40}\n# Sheet: {s}\n{'='*40}\n"
        df = xl.parse(s)
        out += df.to_string(index=False)
    
    with open("xls_report2.txt", "w", encoding="utf-8") as f:
        f.write(out)
    print("Success")
except Exception as e:
    print(f"Error: {e}")
