import os
import pandas as pd
import glob

def inspect_file(file_path):
    print(f"\n=========================================\nFILE: {file_path}")
    try:
        if file_path.endswith('.xls'):
            xls = pd.ExcelFile(file_path, engine='xlrd')
        else:
            xls = pd.ExcelFile(file_path)
            
        for sheet_name in xls.sheet_names:
            print(f"--- Sheet: {sheet_name} ---")
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            print(f"Shape: {df.shape}")
            # Print first 10 rows (non-empty columns only)
            print(df.dropna(how='all').head(15))
    except Exception as e:
        print(f"Error reading {file_path}: {e}")

# Check current directory
excel_files = glob.glob("*.xls*")
for f in excel_files:
    if not os.path.basename(f).startswith('~$'):
        inspect_file(f)

# Also check data folder
if os.path.exists("data"):
    data_files = glob.glob("data/*.xls*")
    for f in data_files:
        if not os.path.basename(f).startswith('~$'):
            inspect_file(f)
