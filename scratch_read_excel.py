import pandas as pd
import glob
import os

folder = r'C:\Users\jjch2\Desktop\2.PAUT 검사 설정 및 빔 방향'
excel_files = glob.glob(os.path.join(folder, '*.xlsx'))

for file in excel_files:
    print(f"\n=========================================")
    print(f"File: {os.path.basename(file)}")
    print(f"=========================================")
    try:
        xls = pd.ExcelFile(file)
        print(f"Sheets: {xls.sheet_names}\n")
        
        for sheet in xls.sheet_names:
            print(f"--- Sheet: {sheet} ---")
            df = pd.read_excel(xls, sheet_name=sheet, nrows=15)
            # Remove entirely empty columns/rows for cleaner print
            df.dropna(how='all', axis=1, inplace=True)
            df.dropna(how='all', axis=0, inplace=True)
            print(df.head(15).to_string())
            print("\n")
    except Exception as e:
        print(f"Error reading {file}: {e}")
