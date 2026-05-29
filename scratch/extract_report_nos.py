import os
import re
import pandas as pd

def extract_details():
    folder = "Na-aba"
    for file in os.listdir(folder):
        if not file.endswith('.xlsm'): continue
        fp = os.path.join(folder, file)
        try:
            xls = pd.ExcelFile(fp)
            for sheet in xls.sheet_names:
                df = pd.read_excel(fp, sheet_name=sheet, header=None)
                # Find cell containing REPORT NO or 성적서 번호
                for r_idx, row in df.iterrows():
                    if r_idx > 30: break
                    for c_idx, val in enumerate(row.values):
                        if pd.isna(val): continue
                        s = str(val)
                        if 'report' in s.lower() or '성적서' in s:
                            print(f"File: {file} | Sheet: {sheet} | Cell ({r_idx}, {c_idx}): {repr(s)}")
        except Exception as e:
            print(f"Error {file}: {e}")

extract_details()
