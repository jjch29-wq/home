import pandas as pd
import glob

def find_sit():
    print("Searching RT_Report files for SIT, RT, or 0001...")
    files = glob.glob("RT_Report_*.xlsx")
    for f in files:
        xls = pd.ExcelFile(f)
        for s in xls.sheet_names:
            df = pd.read_excel(f, sheet_name=s, header=None)
            for r_idx, row in df.iterrows():
                for c_idx, val in enumerate(row.values):
                    if pd.isna(val): continue
                    s_val = str(val)
                    if 'sit' in s_val.lower() or 'rt-000' in s_val.lower() or '0001' in s_val:
                        print(f"File: {f} | Sheet: {s} | Row {r_idx}, Col {c_idx} | Value: {s_val}")

find_sit()
