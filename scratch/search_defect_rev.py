import os
import pandas as pd

def search_defect():
    print("Searching Excel files for 'Defect' or 'Rev'...")
    for root, dirs, files in os.walk('Na-aba'):
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.xlsm'):
                fp = os.path.join(root, file)
                try:
                    xls = pd.ExcelFile(fp)
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(fp, sheet_name=sheet_name, header=None)
                        for r_idx, row in df.iterrows():
                            if r_idx > 50: break
                            for c_idx, val in enumerate(row.values):
                                if pd.isna(val): continue
                                s = str(val).lower()
                                if 'defect' in s or 'rev' in s:
                                    print(f"Found in File: {fp} | Sheet: {sheet_name} | Row {r_idx}, Col {c_idx} | Value: {val}")
                except Exception as e:
                    pass

search_defect()
