import pandas as pd
import glob
import os

folder = r'C:\Users\jjch2\Desktop\2.PAUT 검사 설정 및 빔 방향'
excel_files = glob.glob(os.path.join(folder, '*.xlsx'))

with open(r'c:\Users\jjch2\Desktop\PMI\excel_review.txt', 'w', encoding='utf-8') as f:
    for file in excel_files:
        f.write(f"\n=========================================\n")
        f.write(f"File: {os.path.basename(file)}\n")
        f.write(f"=========================================\n")
        try:
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                f.write(f"--- Sheet: {sheet} ---\n")
                df = pd.read_excel(xls, sheet_name=sheet, nrows=15)
                df.dropna(how='all', axis=1, inplace=True)
                df.dropna(how='all', axis=0, inplace=True)
                f.write(df.to_string())
                f.write("\n\n")
        except Exception as e:
            f.write(f"Error reading {file}: {e}\n")
