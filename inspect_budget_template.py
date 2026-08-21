import pandas as pd

file_path = r"C:\Users\-\OneDrive\바탕 화면\공사실행예산서(중앙지사)수정 (2).xlsx"
out_path = r"C:\Users\-\PMI\budget_template_inspect.txt"

try:
    with open(out_path, "w", encoding="utf-8") as f:
        # read with openpyxl to get exact cell positions or pandas without header
        df = pd.read_excel(file_path, sheet_name="사전원가", header=None, nrows=40, usecols="A:H")
        for idx, row in df.iterrows():
            row_vals = [f"{col_idx}: {val}" for col_idx, val in enumerate(row) if pd.notna(val)]
            if row_vals:
                f.write(f"Row {idx+1}: {row_vals}\n")
except Exception as e:
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(f"Error: {e}")
