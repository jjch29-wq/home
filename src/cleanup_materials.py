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
            # 1. Move '창고' to '모델명' as requested
            # If '모델명' is empty or NaN, use '창고'
            if '창고' in df.columns and '모델명' in df.columns:
                # Prioritize '창고' content for '모델명' as user said 'move warehouse content to model'
                mask = df['창고'].notna() & (df['창고'].astype(str).str.strip() != '')
                df.loc[mask, '모델명'] = df.loc[mask, '창고']
                df.loc[mask, '창고'] = '' # Clear moved content
                print("Moved '창고' content to '모델명'.")

            # 2. Extract SN from '품목명' or '모델명' if 'SN' is empty
            if 'SN' in df.columns:
                for col in ['품목명', '모델명']:
                    if col in df.columns:
                        # Find patterns like SN:XXXX, S/N:XXXX, (XXXX)
                        def extract_sn(row):
                            current_sn = str(row['SN']).strip()
                            if current_sn and current_sn.lower() != 'nan' and current_sn != '':
                                return row # Already has SN
                            
                            text = str(row[col])
                            # Common SN patterns in Korean industry data
                            match = re.search(r'(?:SN|S/N)[:\s-]*([A-Z0-9-]+)', text, re.I)
                            if match:
                                row['SN'] = match.group(1)
                                # Optionally remove from original text? User said "move", so let's try to remove if it's a clear prefix
                                # row[col] = text.replace(match.group(0), '').strip()
                                return row
                            return row
                        
                        df = df.apply(extract_sn, axis=1)
                print("Extracted SNs from item names/models where possible.")

            # 3. Final force Active=1 just in case
            if 'Active' in df.columns:
                df['Active'] = 1

        all_data[sheet] = df

    with pd.ExcelWriter(db_path, engine='openpyxl') as writer:
        for sheet, df in all_data.items():
            df.to_excel(writer, sheet_name=sheet, index=False)
            print(f"Saved sheet: {sheet}")
    print("CLEANUP COMPLETE.")
except Exception as e:
    print(f"ERROR: {e}")
