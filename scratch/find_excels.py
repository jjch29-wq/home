import os
import glob
import pandas as pd

def scan_dir(path):
    print(f"Scanning for '수량' or 'qty' in Excel headers...")
    for root, dirs, files in os.walk(path):
        if '.venv' in root or '.git' in root or 'dist' in root or 'build' in root or 'archive' in root:
            continue
        for file in files:
            if (file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.xlsm')) and not file.startswith('~$'):
                fp = os.path.join(root, file)
                try:
                    if fp.endswith('.xls'):
                        xls = pd.ExcelFile(fp, engine='xlrd')
                    else:
                        xls = pd.ExcelFile(fp)
                    for s in xls.sheet_names:
                        df = pd.read_excel(fp, sheet_name=s, nrows=25, header=None)
                        for idx, row in df.iterrows():
                            vals = [str(v).strip() for v in row.values if pd.notna(v)]
                            # Search for keywords
                            matches = [v for v in vals if any(kw in v.lower() for kw in ['수량', 'qty', 'quantity', 'film', '필름'])]
                            if matches:
                                print(f"File: {fp} | Sheet: {s} | Row {idx}: {vals} (Matched: {matches})")
                except Exception as e:
                    pass

scan_dir('.')
