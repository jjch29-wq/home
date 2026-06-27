import pandas as pd
import os

db_path = 'Material_Inventory.xlsx'
if os.path.exists(db_path):
    print(f"File: {db_path} (Size: {os.path.getsize(db_path)} bytes)")
    try:
        xl = pd.ExcelFile(db_path)
        print(f"Sheets: {xl.sheet_names}")
        for sheet in xl.sheet_names:
            df = pd.read_excel(db_path, sheet_name=sheet)
            print(f"--- Sheet: {sheet} ---")
            print(f"Rows: {len(df)}")
            print(f"Columns: {df.columns.tolist()}")
            if len(df) > 0:
                print(f"First row: {df.iloc[0].to_dict()}")
    except Exception as e:
        print(f"Error reading DB: {e}")
else:
    print("DB file not found!")
