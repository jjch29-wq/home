import pandas as pd
import os
import sys

# Set path to the database
db_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Material_Inventory.xlsx'

if not os.path.exists(db_path):
    print(f"Database file not found at {db_path}")
    sys.exit(1)

print(f"Loading database from {db_path}...")

try:
    with pd.ExcelFile(db_path) as xls:
        print("Sheet names:", xls.sheet_names)
        
        # Check DailyUsage
        if 'DailyUsage' in xls.sheet_names:
            print("\n--- Inspecting DailyUsage ---")
            df = pd.read_excel(xls, 'DailyUsage')
            print("Columns:", df.columns.tolist())
            print("Dtypes:\n", df.dtypes)
            
            # Check for empty strings in float columns
            for col in df.columns:
                if pd.api.types.is_float_dtype(df[col]):
                    print(f"Checking float column '{col}'...")
                    # Check if any value is exactly '' (empty string)
                    mask = df[col].apply(lambda x: x == '')
                    if mask.any():
                        print(f"!!! Found empty strings in float column '{col}'. Count: {mask.sum()}")
                        print(df[mask])
                elif pd.api.types.is_object_dtype(df[col]):
                    # Maybe it should be float but contains strings
                    print(f"Checking object column '{col}' for mix of numbers and strings...")
                    unique_types = df[col].apply(type).unique()
                    print(f"Types in '{col}': {unique_types}")

        # Check Materials
        if 'Materials' in xls.sheet_names:
            print("\n--- Inspecting Materials ---")
            df = pd.read_excel(xls, 'Materials')
            print("Columns:", df.columns.tolist())
            print("Dtypes:\n", df.dtypes)
            for col in ['수량', '재고하한', '가격']:
                if col in df.columns:
                    print(f"Checking '{col}'...")
                    unique_types = df[col].apply(type).unique()
                    print(f"Types in '{col}': {unique_types}")
                    if df[col].apply(lambda x: x == '').any():
                         print(f"!!! Found empty string in '{col}'")

except Exception as e:
    print(f"Error loading Excel: {e}")
    import traceback
    traceback.print_exc()
