
import pandas as pd
import os
import re

# Correct filename identified from MaterialManager-10.py
db_path = "Material_Inventory.xlsx"

if not os.path.exists(db_path):
    print(f"Error: {db_path} not found.")
else:
    try:
        with pd.ExcelFile(db_path) as xls:
            sheets = xls.sheet_names
            print(f"Sheets found: {sheets}")
            
            for sheet in sheets:
                df = pd.read_excel(xls, sheet)
                print(f"\n--- Sheet: {sheet} ---")
                print(f"Columns: {df.columns.tolist()}")
                # print(f"Types:\n{df.dtypes}")
                print(f"Shape: {df.shape}")
                
                # Check for standard columns in DailyUsage
                if sheet == 'DailyUsage':
                    norm_cols = [re.sub(r'\s+', '', str(c)) for c in df.columns]
                    print(f"Normalized Columns: {norm_cols}")
                    
                    # Check for missing critical columns
                    critical = ['Date', 'Site', 'MaterialID', 'Usage', 'EntryTime', 'FilmCount']
                    missing = [c for c in critical if c not in norm_cols]
                    if missing:
                        print(f"WARNING: Missing critical columns: {missing}")
                    else:
                        print("All critical columns present (normalized).")
                
                # print(f"Sample data (first 2 rows):\n{df.head(2)}")
    except Exception as e:
        print(f"Failed to read Excel file: {e}")
