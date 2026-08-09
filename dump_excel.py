import pandas as pd
import traceback

file_path = r"C:\Users\jjch2\Documents\카카오톡 받은 파일\월간진도보고서 (3)\제1호월간진도보고서(2026년5월)\월간진도보고서01호(26년05월).xls"
out_path = r"C:\Users\jjch2\Desktop\PMI\excel_summary.txt"

try:
    xls = pd.ExcelFile(file_path)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(f"Sheets: {xls.sheet_names}\n\n")
        for sheet_name in xls.sheet_names:
            f.write(f"\n{'='*20} Sheet: {sheet_name} {'='*20}\n")
            df = pd.read_excel(xls, sheet_name=sheet_name)
            for r_idx, row in df.iterrows():
                row_vals = [str(x).strip() for x in row if pd.notnull(x) and str(x).strip() != '']
                if row_vals:
                    f.write(f"Row {r_idx}: " + " | ".join(row_vals) + "\n")
except Exception as e:
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(traceback.format_exc())
