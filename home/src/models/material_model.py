import pandas as pd
import numpy as np
from datetime import datetime
import re
from utils.helpers import *

def calculate_current_stock_impl(self, mat_id):
    """Calculate current stock for a material based on transactions (fully robust version)"""
    # 1. Normalize the target mat_id to clean string
    str_mat_id = re.sub(r'\.0$', '', str(mat_id).strip())
    if not str_mat_id or str_mat_id.lower() == 'nan':
        return 0.0

    # Helper to get normalized MaterialID series
    def get_norm_series(df):
        return df['MaterialID'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

    # 2. Get Transaction Sum
    net_trans_qty = 0.0
    if not self.transactions_df.empty:
        mat_trans = self.transactions_df[get_norm_series(self.transactions_df) == str_mat_id]
        net_trans_qty = float(mat_trans['Quantity'].sum()) if not mat_trans.empty else 0.0
    
    # 3. Get Base Quantity and Model Name from Materials Master
    stored_qty = 0.0
    is_chemical_group = False
    if not self.materials_df.empty:
        mat_rows = self.materials_df[get_norm_series(self.materials_df) == str_mat_id]
        if not mat_rows.empty:
            mat = mat_rows.iloc[0]
            model_name = str(mat.get('모델명', '')).strip().upper()
            if model_name == 'MT약품' and str(mat.get('관리단위', '')).strip().upper() != 'EA':
                # Only exclude non-consumables (usually EA units are for equipment)
                # If it's a chemical but mislabeled as MT약품, we might still want to track it if it has a non-EA unit.
                # However, safer for now is to just fix normalization first. 
                # Let's keep the user's specific request but fix the ID matching below.
                return 0.0
            
            val = mat.get('수량', 0)
            try: stored_qty = float(str(val).replace(',', '')) if pd.notna(val) else 0.0
            except: stored_qty = 0.0
            
    return stored_qty + net_trans_qty


def _sync_dataframe_schema_impl(self, df, sheet_name):
    """Ensure the given DataFrame has all required columns for its sheet_name."""
    if df is None:
        return None
        
    schemas = {
        'Materials': [
            'MaterialID', '회사코드', '관리품번', '품목명', 'SN', '창고',
            '모델명', '규격', '품목군코드', '공급업체', '제조사', '제조국', 
            '가격', '원가', '관리단위', '수량', '재고하한', 'Active'
        ],
        'Transactions': [
            'Date', 'MaterialID', 'Site', 'Type', 'Quantity', 'Note', 'User', 
            '차량번호', '주행거리', '차량점검', '차량비고'
        ],
        'MonthlyUsage': [
            'MaterialID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'EntryDate'
        ],
        'DailyUsage': [
            'Date', 'Site', '업체명', 'MaterialID', 'Usage', 'Note', 'EntryTime',
            'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크',
            'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타', '장비명', '검사방법', '검사량',
            '단가', '출장비', '일식', '검사비', '회사코드', 'FilmCount',
            'User', 'WorkTime', 'OT',
            'User2', 'WorkTime2', 'OT2',
            'User3', 'WorkTime3', 'OT3',
            'User4', 'WorkTime4', 'OT4',
            'User5', 'WorkTime5', 'OT5',
            'User6', 'WorkTime6', 'OT6',
            'User7', 'WorkTime7', 'OT7',
            'User8', 'WorkTime8', 'OT8',
            'User9', 'WorkTime9', 'OT9',
            'User10', 'WorkTime10', 'OT10',
            '차량번호', '주행거리', '차량점검', '차량비고', '검사구분', '조인트수', '불량수', '관경(Inch)'
        ],
        'Budget': [
            'Site', 'Revenue', 'UnitPrice', 'LaborCost', 'MaterialCost', 
            'Expense', 'OutsourceCost', 'Profit', 'Note', 'LaborDetail', 'MaterialDetail'
        ]
    }
    
    required_cols = schemas.get(sheet_name, [])
    any_added = False
    for col in required_cols:
        if col not in df.columns:
            # Default values based on col type heuristics
            if col in ['Quantity', 'Usage', '검사량', '단가', '출장비', '일식', '검사비', 'FilmCount', '수량', '원가', '가격', '재고하한']:
                df[col] = 0.0
            elif col == 'Active':
                df[col] = 1
            elif col in ['Date', 'EntryTime', 'EntryDate', 'Year', 'Month']:
                if col in ['Year', 'Month']: df[col] = 0
                else: df[col] = pd.NaT
            else:
                df[col] = ""
            any_added = True
    
    if any_added:
        print(f"DEBUG: Synchronized schema for {sheet_name}. Missing columns added.")
    return df


def _is_consumable_material_impl(self, name, method):
    """
    Determines if a material should be automatically registered and tracked as stock.
    Consumables: RT Films, MT/PT drugs, chemicals.
    Equipment (exclude): Scanners, Crawlers, Sources, Meters, Yokes, etc.
    """
    if not name: return False
    n = str(name).strip().upper().replace(' ', '')
    m = str(method).strip().upper()

    # 1. MT/PT consumables (NDT drugs) - Enhanced keywords
    ndt_keywords = [x.upper().replace(' ', '') for x in self.ndt_materials_all]
    ndt_keywords += ['PT약품', 'MT약품', 'NDT약품', '침투액', '세척액', '현상액', '자분액']
    
    # Stricter equipment check for MT/PT
    equip_keywords = ['YOKE', '장비', 'EQUIP', 'METER', 'GAUGE', 'UVLAMP', '전등', '라이트', '자화', 'SCANNER', '스캐너', 'CRAWLER', '크롤러']
    
    if any(kw in n for kw in ndt_keywords):
        # Check if it also contains equipment keywords (e.g. "MT 장비")
        if any(kw in n for kw in equip_keywords):
            return False
        return True
        
    # 2. RT consumables (Films)
    rt_keywords = ['FILM', 'CARESTREAM', 'MX125', 'T200', 'AA400', 'HS800', 'IX100', 'AGFA', 'FUJI']
    if any(kw in n for kw in rt_keywords):
        if any(kw in n for kw in equip_keywords): # e.g. "RT 장비"
            return False
        return True

    # 3. Method-based logic
    if m in ['MT', 'PT']:
        if any(kw in n for kw in equip_keywords):
            return False
        return True # Default to True for chemicals in MT/PT

    # 4. If name contains MT/PT but no equip keywords, and method is empty
    # This helps in Inventory Status loop where method is empty
    if not m and (n.startswith('PT') or n.startswith('MT')) and len(n) <= 10:
         if not any(kw in n for kw in equip_keywords):
             return True

    # 5. Default: If method is PAUT/UT/RT/PMI, most items are equipment
    return False


def get_material_defaults_impl(self):
    """Extract material defaults from settings_df"""
    if not hasattr(self, 'settings_df') or self.settings_df.empty:
        return [
            ("PT 약품", "세척제", "CAN", 1500), ("PT 약품", "침투제", "CAN", 2300),
            ("PT 약품", "현상제", "CAN", 2000), ("MT 약품", "백색페인트", "CAN", 2350),
            ("MT 약품", "흑색자분", "CAN", 1800), ("방사선투과검사 필름", "MX125", "매", 990),
            ("글리세린", "20L", "통", 100000), ("필름 현상액", "3L", "통", 16500),
            ("필름 정착액", "3L", "통", 16500), ("수적방지액", "200mL", "통", 2500)
        ]
    df = self.settings_df[self.settings_df['Category'] == 'Material']
    if df.empty:
         return [
            ("PT 약품", "세척제", "CAN", 1500), ("PT 약품", "침투제", "CAN", 2300),
            ("PT 약품", "현상제", "CAN", 2000), ("MT 약품", "백색페인트", "CAN", 2350),
            ("MT 약품", "흑색자분", "CAN", 1800), ("방사선투과검사 필름", "MX125", "매", 990),
            ("글리세린", "20L", "통", 100000), ("필름 현상액", "3L", "통", 16500),
            ("필름 정착액", "3L", "통", 16500), ("수적방지액", "200mL", "통", 2500)
        ]
    return [tuple(x) for x in df[['Name', 'Spec', 'Unit', 'Rate']].values]


