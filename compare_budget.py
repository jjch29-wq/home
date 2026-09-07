import pandas as pd
import json
import numpy as np

# Convert everything to basic python types
def coerce_val(v):
    if pd.isna(v): return ""
    if isinstance(v, (np.integer, int)): return int(v)
    if isinstance(v, (np.floating, float)): return float(v)
    return str(v)

try:
    df_dl = pd.read_excel('C:/Users/-/Downloads/공사실행예산서(중앙지사) (4).xlsx', sheet_name='사전원가')
    # Print non-empty rows
    for i, row in df_dl.iterrows():
        vals = [coerce_val(v) for v in row.values]
        if any(v != "" for v in vals):
            print(f"Row {i:02d}: {vals[:10]}")
except Exception as e:
    print(f"Error reading Excel: {e}")

print("\n\n--- App Budget ---")
try:
    df_inv = pd.read_excel('c:/Users/-/PMI/home/src/site_apps/central/data/Material_Inventory.xlsx', sheet_name='Budget')
    central_row = df_inv[df_inv['Site'] == '중앙지사']
    for col in central_row.columns:
        val = central_row.iloc[0][col]
        print(f"[{col}]")
        if isinstance(val, str) and val.startswith('{') or isinstance(val, str) and val.startswith('['):
            try:
                print(json.dumps(json.loads(val), indent=2, ensure_ascii=False))
            except:
                print(val)
        else:
            print(val)
except Exception as e:
    print(f"Error reading Budget: {e}")
