import os
import pandas as pd

def search_defect_anywhere():
    print("Searching Excel files for 'Defect' or 'Rev' anywhere in columns...")
    for root, dirs, files in os.walk('.'):
        if '.venv' in root or '.git' in root:
            continue
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.xlsm'):
                fp = os.path.join(root, file)
                try:
                    xls = pd.ExcelFile(fp)
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(fp, sheet_name=sheet_name, header=None)
                        for r_idx, row in df.iterrows():
                            # check the whole sheet
                            for c_idx, val in enumerate(row.values):
                                if pd.isna(val): continue
                                s = str(val).lower()
                                if 'defect' in s and 'rev' in s:
                                    print(f"MATCH 'defect' and 'rev' in File: {fp} | Sheet: {sheet_name} | Row {r_idx}, Col {c_idx} | Value: {val}")
                                elif 'defect' in s:
                                    # Print to see context
                                    if any(term in s for term in ['rev', '수량', 'qty', 'num']):
                                        print(f"MATCH 'defect' context in File: {fp} | Sheet: {sheet_name} | Row {r_idx}, Col {c_idx} | Value: {val}")
                except Exception as e:
                    pass

search_defect_anywhere()
