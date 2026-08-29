import pandas as pd

def check_sheets(file_path):
    print(f"\n--- {file_path} ---")
    try:
        df_d = pd.read_excel(file_path, sheet_name='DailyUsage')
        print(f"DailyUsage rows: {len(df_d)}")
        
        df_t = pd.read_excel(file_path, sheet_name='Transactions')
        print(f"Transactions rows: {len(df_t)}")
        
        df_m = pd.read_excel(file_path, sheet_name='Materials')
        print(f"Materials rows: {len(df_m)}")
        
    except Exception as e:
        print(f"Error: {e}")

check_sheets('home/data/Material_Inventory.xlsx')
check_sheets('home/data/Kogas_Material_Inventory.xlsx')
