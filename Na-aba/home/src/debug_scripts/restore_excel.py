import pandas as pd
import os
import re

db_path = r"c:\Users\-\OneDrive\바탕 화면\home\data\Material_Inventory.xlsx"

try:
    xl = pd.ExcelFile(db_path)
    all_data = {}
    for sheet in xl.sheet_names:
        df = pd.read_excel(db_path, sheet_name=sheet)
        print(f"Processing sheet: {sheet}")
        
        if sheet == 'Materials':
            if '관리단위' in df.columns and '수량' in df.columns:
                # [USER REQUEST] Move '관리단위' content to '수량'
                def move_unit_to_qty(row):
                    unit_val = str(row['관리단위']).strip()
                    qty_val = str(row['수량']).strip()
                    
                    # If '관리단위' looks like a number, move it to '수량'
                    if unit_val.isdigit() or (unit_val.replace('.','',1).isdigit()):
                        row['수량'] = unit_val
                        row['관리단위'] = 'EA' # Default back to EA or empty?
                    return row
                
                df = df.apply(move_unit_to_qty, axis=1)
                
                # Convert 수량 to numeric
                df['수량'] = pd.to_numeric(df['수량'], errors='coerce').fillna(0)
                print("Moved '관리단위' numbers to '수량' and reset units to EA.")

        all_data[sheet] = df

    with pd.ExcelWriter(db_path, engine='openpyxl') as writer:
        for sheet, df in all_data.items():
            df.to_excel(writer, sheet_name=sheet, index=False)
            print(f"Saved sheet: {sheet}")
    print("RESTORATION V6 COMPLETE.")
except Exception as e:
    print(f"ERROR: {e}")
