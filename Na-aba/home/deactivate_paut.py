import pandas as pd
import os

file_path = 'data/Material_Inventory.xlsx'
if os.path.exists(file_path):
    # Load all sheets
    excel_file = pd.ExcelFile(file_path)
    sheets = {name: excel_file.parse(name) for name in excel_file.sheet_names}
    
    if 'Materials' in sheets:
        df = sheets['Materials']
        # Find rows with PAUT
        paut_mask = df['품목명'].astype(str).str.contains('PAUT', na=False, case=False)
        print(f"Found {paut_mask.sum()} PAUT-related materials.")
        
        # Set Active to 0
        df.loc[paut_mask, 'Active'] = 0
        
        # Save back to Excel
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for name, sheet_df in sheets.items():
                sheet_df.to_excel(writer, sheet_name=name, index=False)
        print("Updated Materials sheet: Set PAUT items to Active=0.")
    else:
        print("Materials sheet not found.")
else:
    print(f"File not found: {file_path}")
