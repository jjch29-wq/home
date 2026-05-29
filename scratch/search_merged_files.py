import os
import glob
import pandas as pd

def search_merged():
    print("Searching for Merged files...")
    for root, dirs, files in os.walk('.'):
        if '.venv' in root or '.git' in root:
            continue
        for file in files:
            if "merged" in file.lower() and file.endswith('.xlsx'):
                fp = os.path.join(root, file)
                print(f"\nFound Merged File: {fp}")
                try:
                    xls = pd.ExcelFile(fp)
                    print(f"  Sheets: {xls.sheet_names}")
                    for s in xls.sheet_names:
                        df = pd.read_excel(fp, sheet_name=s)
                        print(f"  Sheet '{s}' Columns: {list(df.columns)}")
                        print("  Last 5 rows:")
                        print(df.tail(5).to_string())
                except Exception as e:
                    print(f"  Error: {e}")

search_merged()
