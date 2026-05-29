import pandas as pd

fp = r"Na-aba/home/data/가스공사 의뢰서.xlsx"
try:
    xls = pd.ExcelFile(fp)
    print("Sheets in 가스공사 의뢰서.xlsx:", xls.sheet_names)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(fp, sheet_name=sheet_name, header=None)
        # Search for Defect, Rev, R1, Repair, etc.
        for r_idx, row in df.iterrows():
            for c_idx, val in enumerate(row.values):
                if pd.isna(val): continue
                s = str(val).lower()
                if 'defect' in s or 'rev' in s or 'repair' in s or 'r1' in s or 'r2' in s:
                    print(f"Sheet: {sheet_name} | Row {r_idx}, Col {c_idx} | Value: {val}")
except Exception as e:
    print("Error:", e)
