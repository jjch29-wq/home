import os
import pandas as pd

def print_root_excels():
    print("Checking Excel files in the root folder...")
    for file in os.listdir('.'):
        if file.endswith('.xlsx') or file.endswith('.xls'):
            print(f"\n--- File: {file} ---")
            try:
                xls = pd.ExcelFile(file)
                print("Sheets:", xls.sheet_names)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(file, sheet_name=sheet)
                    print(f"  Sheet: {sheet} | Shape: {df.shape} | Columns: {list(df.columns[:8])}")
                    # Print first few rows to see if SIT-K1-JHC-PIP-RT-0001 is present
                    mask = df.astype(str).apply(lambda row: row.str.contains('SIT-K1-JHC|RT-0001', case=False).any(), axis=1)
                    matching_rows = df[mask]
                    if len(matching_rows) > 0:
                        print(f"  FOUND matching rows in sheet '{sheet}':")
                        print(matching_rows.to_string())
            except Exception as e:
                print("Error reading:", e)

print_root_excels()
