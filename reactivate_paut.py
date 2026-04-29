import pandas as pd
import os

file_path = 'data/Material_Inventory.xlsx'
if os.path.exists(file_path):
    excel_file = pd.ExcelFile(file_path)
    sheets = {name: excel_file.parse(name) for name in excel_file.sheet_names}
    
    if 'Materials' in sheets:
        df = sheets['Materials']
        paut_mask = df['품목명'].astype(str).str.contains('PAUT', na=False, case=False)
        df.loc[paut_mask, 'Active'] = 1
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for name, sheet_df in sheets.items():
                sheet_df.to_excel(writer, sheet_name=name, index=False)
        print("Re-activated PAUT items (Active=1).")
    else:
        print("Materials sheet not found.")
else:
    print(f"File not found: {file_path}")
