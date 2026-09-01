import sys, os
sys.path.insert(0, r'c:\Users\-\PMI\home\src')

# Load the app module to access daily_usage_df  
# Instead, let's check the Excel file directly
import glob, pandas as pd

xlsx_files = glob.glob(r'c:\Users\-\PMI\home\**\*.xlsx', recursive=True)
print('XLSX files:', xlsx_files)

for f in xlsx_files:
    if 'Material_Inventory' in f or 'Lotte' in f:
        print(f'\n=== {f} ===')
        try:
            xl = pd.ExcelFile(f)
            print('Sheets:', xl.sheet_names)
            for sheet in xl.sheet_names:
                df = xl.parse(sheet, nrows=3)
                print(f'\n  Sheet: {sheet}')
                print('  Cols:', list(df.columns))
                # Check for daily usage related columns
                if any(k in str(list(df.columns)) for k in ['User', '인건비', '업체명', '회사코드', '검사비']):
                    print('  ** DAILY USAGE SHEET FOUND **')
                    # Show relevant columns
                    rel_cols = [c for c in df.columns if c in ['Date', 'Site', '업체명', '회사코드', 'User', 'User2', '인건비', '검사비', '재료비', '제경비', '기술료']]
                    print('  Relevant cols:', rel_cols)
                    print(df[rel_cols].to_string())
        except Exception as e:
            print(f'  ERROR: {e}')
