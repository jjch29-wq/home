import pandas as pd
import os

filepath = r"F:\내 드라이브\Office\착공\Project PROVIDENCE 정 비파괴검사\준공\박스라벨.xls"
# Wait, the path might be slightly different.
filepath = r"F:\내 드라이브\Office\착공\Project PROVIDENCE 중 비파괴검사\준공\박스라벨.xls"

try:
    print(f"Reading {filepath}...")
    xls = pd.ExcelFile(filepath)
    print(f"Sheets: {xls.sheet_names}")
    for sheet in xls.sheet_names:
        print(f"\n--- Sheet: {sheet} ---")
        df = pd.read_excel(filepath, sheet_name=sheet, header=None)
        print(f"Shape: {df.shape}")
        if not df.empty:
            print(df.head(10).to_string())
except Exception as e:
    print(f"Error: {e}")
