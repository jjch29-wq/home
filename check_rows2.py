import pandas as pd

def check_sheets(file_path):
    print(f"\n--- {file_path} ---")
    try:
        sheets = ['MonthlyUsage', 'Budget', 'Settings']
        for s in sheets:
            df = pd.read_excel(file_path, sheet_name=s)
            print(f"{s} rows: {len(df)}")
    except Exception as e:
        print(f"Error: {e}")

check_sheets('home/data/Material_Inventory.xlsx')
check_sheets('home/data/Kogas_Material_Inventory.xlsx')
