import os
import pandas as pd

def run_debug():
    f = "c:/Users/jjch2/Desktop/보고서/Project PROVIDENCE/Request/PMI/Na-aba/home/data/가스공사 의뢰서.xlsx"
    print(f"\n==========================================")
    print(f"Debugging File: {os.path.basename(f)}")
    try:
        with pd.ExcelFile(f) as xls:
            for sheet_name in xls.sheet_names[:1]:
                print(f"Sheet: {sheet_name}")
                temp_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                
                # Extract sheet level dwg
                sheet_level_dwg = ""
                for r_pos in [4, 5]:
                    for c_pos in [2, 1, 3]:
                        if r_pos < len(temp_df) and c_pos < len(temp_df.columns):
                            val = temp_df.iloc[r_pos, c_pos]
                            val_str = str(val).strip()
                            if pd.notna(val):
                                if val_str and val_str.lower() != 'nan' and len(val_str) > 2:
                                    if val_str in ["도면번호", "도면", "DWG", "DWG NO", "DWG.NO", "도면 NO", "도면 NO."]:
                                        continue
                                    val_upper = val_str.upper()
                                    if any(k in val_upper for k in ["CO.", "LTD", "INSPECTION", "TESTING", "CORP", "INC", "SEOUL", "주식회사"]):
                                        continue
                                    if val_str.count(" ") > 2:
                                        continue
                                    sheet_level_dwg = val_str
                                    break
                    if sheet_level_dwg: break
                print(f"Extracted sheet_level_dwg: '{sheet_level_dwg}'")
                
                # Test the pairing row loop
                col_dwg = 2 # Column C
                start_row = 13 # index 13 (Row 14)
                print("\n--- Simulating first 3 joint pairings ---")
                for r_idx in range(start_row, start_row + 6, 2):
                    row_top = temp_df.iloc[r_idx]
                    row_bot = temp_df.iloc[r_idx + 1] if r_idx + 1 < len(temp_df) else None
                    
                    # Top Row Dwg
                    curr_dwg = sheet_level_dwg if sheet_level_dwg else (str(row_top[col_dwg]).strip() if col_dwg is not None else '')
                    
                    # Bottom Row Dwg (Sub)
                    curr_dwg_sub = ''
                    if sheet_level_dwg:
                        curr_dwg_sub = sheet_level_dwg
                    elif row_bot is not None and col_dwg is not None:
                        curr_dwg_sub = str(row_bot[col_dwg]).strip()
                    if not curr_dwg_sub or curr_dwg_sub == 'nan':
                        curr_dwg_sub = curr_dwg
                        
                    print(f"Joint {int((r_idx-start_row)/2)+1} (Rows {r_idx+1}-{r_idx+2}): Dwg='{curr_dwg}', Dwg_Sub='{curr_dwg_sub}'")
                    
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    run_debug()
