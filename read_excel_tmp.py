import pandas as pd
import sys

file_path = r"C:\Users\-\OneDrive\바탕 화면\산출내역서(2026년 중앙지사 열수송관 비파괴검사용역 단가계약).xlsx"
out_path = r"C:\Users\-\PMI\excel_output2.txt"

try:
    with open(out_path, "w", encoding="utf-8") as f:
        xl = pd.ExcelFile(file_path)
        f.write(f"Sheets: {xl.sheet_names}\n")
        for sheet in xl.sheet_names:
            f.write(f"\n--- Sheet: {sheet} ---\n")
            df = xl.parse(sheet)
            f.write(df.head(40).to_string())
            f.write(f"\nShape: {df.shape}\n")
except Exception as e:
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(f"Error: {e}")
