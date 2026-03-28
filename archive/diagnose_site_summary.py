import pandas as pd
import os
import re

db_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Material_Inventory.xlsx'

def _parse_ot_hours(ot_value):
    if not ot_value or pd.isna(ot_value): return 0.0
    val = str(ot_value).strip()
    match = re.search(r'(\d+\.?\d*)\s*(시간|H|h)', val)
    if match: return float(match.group(1))
    match = re.search(r'^\s*(\d+\.?\d*)\s*$', val)
    if match: return float(match.group(1))
    return 0.0

def calculate_ot_amount(ot_value):
    if not ot_value or pd.isna(ot_value): return 0
    val = str(ot_value).strip()
    if '(' in val and '원)' in val:
        try:
            return int(float(val.split('(')[1].split('원')[0].replace(',', '').strip()))
        except: pass
    h = _parse_ot_hours(val)
    if '휴일' in val: return int(h * 6000)
    if '야간' in val: return int(h * 5000)
    return int(h * 4000)

if os.path.exists(db_path):
    df = pd.read_excel(db_path, sheet_name='DailyUsage')
    # Normalize columns
    df.columns = [re.sub(r'\s+', '', str(c)) for c in df.columns]
    
    # Filter for 가스공사
    gas_df = df[df['Site'].str.contains('가스공사', na=False, case=False)].copy()
    
    if gas_df.empty:
        print("가스공사 데이터를 찾을 수 없습니다.")
    else:
        print(f"가스공사 데이터 건수: {len(gas_df)}")
        
        # Calculate row-level aggregates same as MaterialManager
        def calc_row_values(row):
            h_sum = 0
            a_sum = 0
            w_count = 0
            for i in range(1, 11):
                u_c = 'User' if i == 1 else f'User{i}'
                if u_c in row and pd.notna(row[u_c]) and str(row[u_c]).strip() not in ['', 'nan', '0.0']:
                    w_count += 1
                
                ot_c = 'OT' if i == 1 else f'OT{i}'
                if ot_c in row and pd.notna(row[ot_c]):
                    v = str(row[ot_c]).strip()
                    if v:
                        h_sum += _parse_ot_hours(v)
                        a_sum += calculate_ot_amount(v)
            return pd.Series([h_sum, a_sum, w_count])

        gas_df[['OTH', 'OTA', 'WC']] = gas_df.apply(calc_row_values, axis=1)
        
        # Site Summary
        summary = {
            '총공수': gas_df['WC'].sum(),
            'OT시간': gas_df['OTH'].sum(),
            'OT금액': gas_df['OTA'].sum(),
            '검사비': gas_df['검사비'].sum() if '검사비' in gas_df.columns else 0,
            '출장비': gas_df['출장비'].sum() if '출장비' in gas_df.columns else 0,
            '일식': gas_df['일식'].sum() if '일식' in gas_df.columns else 0,
        }
        summary['총합계'] = summary['검사비'] + summary['출장비'] + summary['일식']
        
        print("\n--- 가스공사 누계 집계 결과 ---")
        for k, v in summary.items():
            if isinstance(v, float):
                print(f"{k}: {v:.1f}")
            else:
                print(f"{k}: {v:,}")
        
        print("\n--- 상세 데이터 (최근 10건) ---")
        cols_to_show = ['Date', 'Site', 'WC', 'OTA', '검사비', '출장비', '일식']
        existing_cols = [c for c in cols_to_show if c in gas_df.columns]
        print(gas_df[existing_cols].tail(10))
else:
    print(f"파일을 찾을 수 없습니다: {db_path}")
