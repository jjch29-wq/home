import pandas as pd
import json

file_path = r"C:\Users\jjch2\Desktop\산출내역서(가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역)_감독경유.xlsx"
xl = pd.ExcelFile(file_path)

sheet_name = None
for name in xl.sheet_names:
    if "6-1." in name or "RT" in name:
        sheet_name = name
        break

if sheet_name:
    df = xl.parse(sheet_name)
    df.to_csv("dump_rt_sheet.csv", index=False, encoding="utf-8-sig")
    print(f"Dumped sheet: {sheet_name} to dump_rt_sheet.csv")
else:
    print("Could not find sheet for RT")
