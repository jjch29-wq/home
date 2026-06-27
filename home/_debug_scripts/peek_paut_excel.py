import pandas as pd
import os

file_path = 'data/Material_Inventory.xlsx'
if os.path.exists(file_path):
    df = pd.read_excel(file_path, sheet_name='Materials')
    print("--- Materials Sheet ---")
    paut_rows = df[df['품목명'].astype(str).str.contains('PAUT', na=False)]
    print(paut_rows)
    
    df_usage = pd.read_excel(file_path, sheet_name='DailyUsage')
    print("\n--- DailyUsage Sheet (PAUT) ---")
    # Need to find the MaterialID for PAUT
    paut_ids = paut_rows['MaterialID'].tolist()
    usage_rows = df_usage[df_usage['MaterialID'].isin(paut_ids)]
    print(usage_rows)
else:
    print(f"File not found: {file_path}")
