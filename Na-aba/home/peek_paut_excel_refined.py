import pandas as pd
import os

file_path = 'data/Material_Inventory.xlsx'
if os.path.exists(file_path):
    df = pd.read_excel(file_path, sheet_name='Materials')
    print("--- Materials Sheet (PAUT) ---")
    paut_rows = df[df['품목명'].astype(str).str.contains('PAUT', na=False)]
    if not paut_rows.empty:
        print(paut_rows[['MaterialID', '품목명', '모델명', '규격', 'Active']])
    else:
        print("No PAUT found in Materials sheet.")
    
    df_usage = pd.read_excel(file_path, sheet_name='DailyUsage')
    print("\n--- DailyUsage Sheet (PAUT) ---")
    paut_ids = paut_rows['MaterialID'].tolist()
    usage_rows = df_usage[df_usage['MaterialID'].isin(paut_ids)]
    if not usage_rows.empty:
        print(usage_rows[['Date', 'Site', 'MaterialID', 'Usage', 'Note']])
    else:
        print("No usage records for PAUT.")
else:
    print(f"File not found: {file_path}")
