import os
import pandas as pd

def inspect_rt_file(file_path):
    print(f"\n=========================================\nFILE: {file_path}")
    try:
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            print(f"--- Sheet: {sheet_name} ---")
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            print(f"Shape: {df.shape}")
            for idx in range(min(15, len(df))):
                row = df.iloc[idx]
                row_vals = [f"{c_idx}: {val}" for c_idx, val in enumerate(row.values) if pd.notna(val)]
                print(f"Row {idx}: {row_vals}")
    except Exception as e:
        print(f"Error reading {file_path}: {e}")

inspect_rt_file("RT_Report_20260524_204924.xlsx")
