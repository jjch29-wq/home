import os
import pandas as pd

def find_val():
    print("Searching Excel files for cells containing '-RT-' or 'RT-00'...")
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
                            for c_idx, val in enumerate(row.values):
                                if pd.isna(val): continue
                                s = str(val).strip()
                                if '-rt-' in s.lower() or 'rt-00' in s.lower():
                                    print(f"MATCH: File: {fp} | Sheet: {sheet_name} | Row {r_idx}, Col {c_idx} | Value: {val}")
                except Exception as e:
                    pass

find_val()
