import pandas as pd
import json
import numpy as np

def coerce_val(v):
    if pd.isna(v): return ""
    if isinstance(v, (np.integer, int)): return int(v)
    if isinstance(v, (np.floating, float)): return float(v)
    return str(v)

df_dl = pd.read_excel('C:/Users/-/Downloads/공사실행예산서(중앙지사) (4).xlsx', sheet_name='사전원가')
dl_data = []
for i, row in df_dl.iterrows():
    vals = [coerce_val(v) for v in row.values]
    if any(v != "" for v in vals):
        dl_data.append(vals[:10])

df_inv = pd.read_excel('c:/Users/-/PMI/home/src/site_apps/central/data/Material_Inventory.xlsx', sheet_name='Budget')
central_row = df_inv[df_inv['Site'] == '중앙지사']
app_data = {}
for col in central_row.columns:
    val = central_row.iloc[0][col]
    if isinstance(val, str) and (val.startswith('{') or val.startswith('[')):
        try:
            app_data[col] = json.loads(val)
        except:
            app_data[col] = val
    else:
        app_data[col] = coerce_val(val)

out = {
    'downloaded_excel': dl_data,
    'app_budget': app_data
}

with open('c:/Users/-/PMI/budget_comparison_results.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
