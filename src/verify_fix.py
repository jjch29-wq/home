import pandas as pd
import re
import os

# 1. Mock necessary constants and helpers
NAN_PATTERN = re.compile(r'^nan(\.0+)?$|^none$|^null$|^0\.0+|-0\.0+$', re.IGNORECASE)
DOT_ZERO_PATTERN = re.compile(r'\.0$')

def clean_nan(val):
    if pd.isna(val) or val is None: return ""
    s = str(val).strip()
    if not s or NAN_PATTERN.match(s):
        return ""
    s = DOT_ZERO_PATTERN.sub('', s)
    if NAN_PATTERN.match(s): return ""
    return s

def normalize_id(val):
    if pd.isna(val) or val == '' or str(val).lower() == 'nan': return ""
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def _is_consumable_material(name, method):
    if not name: return False
    n = str(name).strip().upper().replace(' ', '')
    m = str(method).strip().upper()
    ndt_materials_all = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
    ndt_keywords = [x.upper().replace(' ', '') for x in ndt_materials_all]
    if any(kw in n for kw in ndt_keywords): return True
    rt_keywords = ['FILM', 'CARESTREAM', 'MX125', 'T200', 'AA400', 'HS800', 'IX100', 'AGFA', 'FUJI']
    if any(kw in n for kw in rt_keywords): return True
    if m in ['MT', 'PT']:
        equip_keywords = ['YOKE', '장비', 'EQUIP', 'METER', 'GAUGE', 'UVLAMP']
        if any(kw in n for kw in equip_keywords): return False
        return True
    return False

def normalize_cols(df):
    if df is not None and not df.empty:
        df.columns = [re.sub(r'\s+', '', str(c)) for c in df.columns]
        new_cols = {}
        for col in df.columns:
            c = str(col)
            if 'ǰ' in c: new_cols[col] = '품목명'
            elif '𵨸' in c: new_cols[col] = '모델명'
            elif 'â' in c: new_cols[col] = '창고'
            elif '˻' in c: 
                if '緮' in c or '量' in c: new_cols[col] = '검사량'
                else: new_cols[col] = '검사방법'
            elif 'ܰ' in c: new_cols[col] = '단가'
            elif 'ȸڵ' in c: new_cols[col] = '회사코드'
        if new_cols:
            print(f"DEBUG: Fixing garbled columns: {new_cols}")
            df.rename(columns=new_cols, inplace=True)
    return df

# 2. Run logic on actual file
path = '../data/Material_Inventory.xlsx'
materials_df = pd.read_excel(path, sheet_name='Materials')
materials_df = normalize_cols(materials_df)

daily_usage_df = pd.read_excel(path, sheet_name='DailyUsage')
daily_usage_df = normalize_cols(daily_usage_df)

# Consumable IDs from master
consumable_ids = set()
for _, m in materials_df.iterrows():
    m_name = str(m.get('품목명', '')).strip()
    if _is_consumable_material(m_name, ''):
        c_id = normalize_id(m.get('MaterialID'))
        if c_id: consumable_ids.add(c_id)

print(f"DEBUG: Consumable IDs in Master: {len(consumable_ids)}")

# Test Daily Usage Lookup
temp_daily = daily_usage_df.copy()
temp_daily['NormID'] = temp_daily['MaterialID'].apply(normalize_id)

def _f(v):
    if pd.isna(v) or v is None: return 0.0
    try: return float(str(v).replace(',', '').strip()) if str(v).strip() else 0.0
    except: return 0.0

temp_daily['TotalUsage'] = temp_daily.apply(lambda r: 
    _f(r.get('Usage', r.get('검사량', r.get('수량', r.get('Quantity', 0))))) + 
    _f(r.get('FilmCount', r.get('매수', 0))), axis=1)

temp_daily['IsConsumable'] = temp_daily.apply(lambda r: 
    (r['NormID'] in consumable_ids) or 
    _is_consumable_material(
        str(r.get('품목명', r.get('장비명', ''))).strip(), 
        str(r.get('검사방법', '')).strip()
    ), axis=1)

temp_consumable = temp_daily[temp_daily['IsConsumable']]
daily_usage_lookup = temp_consumable.groupby('NormID')['TotalUsage'].sum().to_dict()
temp_consumable['NormName'] = temp_consumable.apply(lambda r: str(r.get('품목명', r.get('장비명', ''))).strip(), axis=1)
daily_name_lookup = temp_consumable.groupby('NormName')['TotalUsage'].sum().to_dict()

print(f"DEBUG: Daily Consumable Lookup built: {daily_name_lookup}")

# Test Stock Calc for Material 410
mat_rows = materials_df[materials_df['MaterialID'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True) == '410']
if not mat_rows.empty:
    mat = mat_rows.iloc[0]
    mat_id = mat['MaterialID']
    str_mat_id = normalize_id(mat_id)
    
    daily_qty = daily_usage_lookup.get(str_mat_id, 0.0)
    mat_name_str = str(mat.get('품목명', '')).strip()
    if daily_qty == 0:
        daily_qty = daily_name_lookup.get(mat_name_str, 0.0)
        
    val = mat.get('수량', 0)
    try: stored_qty = float(str(val).replace(',', '')) if pd.notna(val) else 0.0
    except: stored_qty = 0.0
    
    mat_name_raw = clean_nan(mat.get('품목명', ''))
    model_name_raw = clean_nan(mat.get('모델명', ''))
    
    print(f"DEBUG: Stock Calc for '{mat_name_raw} ({model_name_raw})': Master={stored_qty}, Daily={daily_qty}")
else:
    print("Material 410 not found in Master")
