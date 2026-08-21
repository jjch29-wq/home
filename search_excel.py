import pandas as pd
import sys

file_path = r"C:\Users\-\PMI\home\data\Material_Inventory.xlsx"
out_path = r"C:\Users\-\PMI\excel_search_out.txt"

try:
    with open(out_path, "w", encoding="utf-8") as f:
        xl = pd.ExcelFile(file_path)
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            f.write(f"\n--- Sheet: {sheet} ---\n")
            for idx, row in df.iterrows():
                row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
                if "중앙지사" in row_str or "한국지역난방" in row_str:
                    f.write(f"Row {idx}: {row_str}\n")
except Exception as e:
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(f"Error: {e}")
