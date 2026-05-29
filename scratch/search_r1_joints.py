import os
import pandas as pd

def search_r1_joints():
    print("Searching for R1, R2, Rev, or defect joints...")
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
                                # Look for typical revision/repair identifiers
                                if s in ['R1', 'R2', 'R3', 'Rev.1', 'Rev.2'] or any(term in s.upper() for term in ['-R1', '-R2', '-R3', '/R1', '/R2', 'REPAIR', 'DEFECT']):
                                    print(f"File: {fp} | Sheet: {sheet_name} | Row {r_idx}, Col {c_idx} | Value: {val}")
                except Exception as e:
                    pass

search_r1_joints()
