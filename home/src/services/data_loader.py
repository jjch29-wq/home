import pandas as pd
import json
import os
import traceback
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

def load_data_impl(self):
    import re
    def normalize_cols(df):
        if df is not None and not df.empty:
            # 1. Standardize whitespace and handle numeric column names
            df.columns = [re.sub(r'\s+', '', str(c)) for c in df.columns]
            
            # 2. [ROBUST] Fallback mapping for common garbled column names
            # Instead of a complex map, we focus on the most critical ones using partial matches
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
                df.rename(columns=new_cols, inplace=True)
        return df

    try:
        print(f"DEBUG: Loading data from {self.db_path}...")
        
        # Check if database exists in app_dir. If not, try to restore from bundle_dir
        if not os.path.exists(self.db_path):
            bundled_db = os.path.join(self.bundle_dir, 'Material_Inventory.xlsx')
            print(f"DEBUG: Main DB not found. Trying to restore from bundle: {bundled_db}")
            if os.path.exists(bundled_db):
                import shutil
                try:
                    shutil.copy2(bundled_db, self.db_path)
                    print("DEBUG: Restored DB from bundle.")
                    # Also try to copy config if it exists in bundle but not in app_dir
                    bundled_config = os.path.join(self.bundle_dir, 'Material_Manager_Config.json')
                    if os.path.exists(bundled_config) and not os.path.exists(self.config_path):
                        shutil.copy2(bundled_config, self.config_path)
                        print("DEBUG: Restored Config from bundle.")
                except Exception as e:
                    print(f"Failed to restore data from bundle: {e}")

        if not os.path.exists(self.db_path):
            print("DEBUG: DB still not found. Initializing new DataFrames.")
            # Initialize with new schema if still not found
            self.materials_df = pd.DataFrame(columns=[
                'MaterialID', '회사코드', '관리품번', '품목명', 'SN', '창고',
                '모델명', '규격', '품목군코드', '공급업체', '제조사', '제조국', 
                '가격', '원가', '관리단위', '수량', '재고하한', 'Active'
            ])
            self.transactions_df = pd.DataFrame(columns=['Date', 'MaterialID', 'Site', 'Type', 'Quantity', 'Note', 'User', '차량번호', '주행거리', '차량점검', '차량비고'])
            self.monthly_usage_df = pd.DataFrame(columns=['MaterialID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'Entry Date'])
            self.daily_usage_df = pd.DataFrame(columns=['Date', 'Site', '업체명', 'MaterialID', 'Usage', 'Note', 'EntryTime',
                                            'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크',
                                            'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타', '장비명', '검사방법', '검사량',
                                            '단가', '출장비', '일식', '검사비', '회사코드',
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
                                            '차량번호', '주행거리', '차량점검', '차량비고'])
            self.budget_df = pd.DataFrame(columns=['Site', 'Revenue', 'UnitPrice', 'LaborCost', 'MaterialCost', 'Expense', 'OutsourceCost', 'Profit', 'Note', 'LaborDetail', 'MaterialDetail'])
        else:
            print("DEBUG: DB found. Reading Excel...")
            # 1. Materials
            self.materials_df = pd.read_excel(self.db_path, sheet_name='Materials')
            self.materials_df = normalize_cols(self.materials_df)
            self.materials_df = self._sync_dataframe_schema(self.materials_df, 'Materials')
            
            # Handle column rename from '품명' to '품목명'
            if '품명' in self.materials_df.columns and '품목명' not in self.materials_df.columns:
                self.materials_df.rename(columns={'품명': '품목명'}, inplace=True)
            
            # [NEW] Handle various quantity column aliases to prevent 0.0 readings
            qty_aliases = {'기초재고': '수량', '현재고': '수량', 'Stock': '수량', 'Quantity': '수량', 'Qty': '수량', '재고': '수량'}
            for alias, target in qty_aliases.items():
                if alias in self.materials_df.columns and target not in self.materials_df.columns:
                    print(f"DEBUG: Mapping {alias} to {target}")
                    self.materials_df.rename(columns={alias: target}, inplace=True)
            
            print(f"DEBUG: Materials columns loaded: {self.materials_df.columns.tolist()}")
            
            # Ensure Active is numeric and handle NaNs (treat as Active=1)
            self.materials_df['Active'] = pd.to_numeric(self.materials_df['Active'], errors='coerce').fillna(1)
            
            # [STABILITY] Ensure MaterialID is consistently numeric
            if 'MaterialID' in self.materials_df.columns:
                self.materials_df['MaterialID'] = pd.to_numeric(self.materials_df['MaterialID'], errors='coerce')
            
            # Force specific columns to string type and clean numeric artifacts (.0, -0.0)
            str_cols = ['회사코드', '관리품번', '품목명', 'SN', '창고', '모델명', '규격', 
                        '품목군코드', '공급업체', '제조사', '제조국', '관리단위']
            for col in str_cols:
                if col in self.materials_df.columns:
                    # Ensure all strings are stripped and nan-free
                    self.materials_df[col] = self.materials_df[col].astype(str).str.strip().replace(['nan', 'None', 'NULL', '-0.0', '0.0', 'NaN', 'NaN.0'], '')
                    self.materials_df[col] = self.materials_df[col].str.replace(r'\.0$', '', regex=True)
            
            # [NEW] One-time Data Migration: Strip "MT " and "PT " prefixes from NDT medicine models
            # This ensures existing records are correctly summarized in the Inventory Status tab.
            if not self.materials_df.empty:
                ndt_parent_cats = ["PT약품", "MT약품", "NDT약품"]
                # Normalize category names for matching
                temp_cats = self.materials_df['품목명'].str.replace(' ', '').str.upper()
                mask = temp_cats.isin(ndt_parent_cats)
                if mask.any():
                    def strip_ndt_prefix(model):
                        s = str(model).strip()
                        # Check for PT or MT prefix followed by a space
                        if s.upper().startswith("MT "): return s[3:].strip()
                        if s.upper().startswith("PT "): return s[3:].strip()
                        return s
                    
                    self.materials_df.loc[mask, '모델명'] = self.materials_df.loc[mask, '모델명'].apply(strip_ndt_prefix)
                    print(f"DEBUG: Migrated {mask.sum()} NDT items by stripping prefixes.")
            
            # 2. Transactions
            self.transactions_df = pd.read_excel(self.db_path, sheet_name='Transactions')
            self.transactions_df = normalize_cols(self.transactions_df)
            self.transactions_df = self._sync_dataframe_schema(self.transactions_df, 'Transactions')
            
            # Ensure Date column is datetime and MaterialID is numeric
            if not self.transactions_df.empty:
                self.transactions_df['Date'] = pd.to_datetime(self.transactions_df['Date'], errors='coerce')
                self.transactions_df['MaterialID'] = pd.to_numeric(self.transactions_df['MaterialID'], errors='coerce')
                
                # Force string columns and clean numeric artifacts
                for col in ['Type', 'Note', 'User', 'Site', '차량번호', '주행거리', '차량점검', '차량비고']:
                    self.transactions_df[col] = self.transactions_df[col].astype(str).replace(['nan', 'None', 'NULL', '-0.0', '0.0', 'NaN'], '')
                    self.transactions_df[col] = self.transactions_df[col].str.replace(r'\.0$', '', regex=True)
                
                # One-time cleanup: Remove '현장사용' and redundant model names from historical notes
                self.transactions_df['Note'] = self.transactions_df['Note'].astype(str).str.replace('현장사용', '', regex=False).str.strip()
                
                # Clean up notes that are identical to model names
                if not self.transactions_df.empty and not self.materials_df.empty:
                    # Create a map for MaterialID -> Model Name
                    id_to_model = self.materials_df.set_index('MaterialID')['모델명'].astype(str).to_dict()
                    
                    def clean_redundant_note(row):
                        note = str(row['Note']).strip()
                        mat_id = row['MaterialID']
                        model = str(id_to_model.get(mat_id, '')).strip()
                        if note and model and note == model:
                            return ''
                        return note
                        
                    self.transactions_df['Note'] = self.transactions_df.apply(clean_redundant_note, axis=1)
                
                # Normalize OUT transaction quantities to negative and strip Site names
                if not self.transactions_df.empty:
                    if 'Site' in self.transactions_df.columns:
                        self.transactions_df['Site'] = self.transactions_df['Site'].astype(str).str.strip().replace(['nan', 'None'], '')
                    
                    out_mask = (self.transactions_df['Type'] == 'OUT') & (self.transactions_df['Quantity'] > 0)
                    if out_mask.any():
                        self.transactions_df.loc[out_mask, 'Quantity'] = -self.transactions_df.loc[out_mask, 'Quantity']
                        print(f"DEBUG: Normalized {out_mask.sum()} OUT transactions to negative.")
            
            # 3. Monthly Usage
            try:
                self.monthly_usage_df = pd.read_excel(self.db_path, sheet_name='MonthlyUsage', dtype={'Site': str, 'Note': str})
                self.monthly_usage_df = normalize_cols(self.monthly_usage_df)
                self.monthly_usage_df = self._sync_dataframe_schema(self.monthly_usage_df, 'MonthlyUsage')
                
                if not self.monthly_usage_df.empty:
                    self.monthly_usage_df['MaterialID'] = pd.to_numeric(self.monthly_usage_df['MaterialID'], errors='coerce')
                    self.monthly_usage_df['EntryDate'] = pd.to_datetime(self.monthly_usage_df['EntryDate'])
                    
                    self.monthly_usage_df['Site'] = self.monthly_usage_df['Site'].astype(str).str.strip().replace(['nan', 'None'], '')
            except Exception as e:
                print(f"DEBUG: Failed to load MonthlyUsage: {e}")
                self.monthly_usage_df = pd.DataFrame(columns=['MaterialID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'EntryDate'])
            
            # 4. Daily Usage
            try:
                # Explicitly set dtypes for vehicle and note columns to avoid float inference for empty cells
                self.daily_usage_df = pd.read_excel(self.db_path, sheet_name='DailyUsage', 
                                                    dtype={'Site': str, 'Note': str, 'User': str,
                                                           '차량번호': str, '주행거리': str, '차량점검': str, '차량비고': str})
                print(f"DEBUG: Loaded {len(self.daily_usage_df)} records from DailyUsage sheet.")
                self.daily_usage_df = normalize_cols(self.daily_usage_df)
                
                # [NEW] Column Name Migration for Daily Usage
                if '날짜' in self.daily_usage_df.columns:
                    if 'Date' not in self.daily_usage_df.columns:
                        self.daily_usage_df.rename(columns={'날짜': 'Date'}, inplace=True)
                    else:
                        # Merge if both exist
                        self.daily_usage_df['Date'] = self.daily_usage_df['Date'].fillna(self.daily_usage_df['날짜'])
                        self.daily_usage_df.drop(columns=['날짜'], inplace=True)

                if '현장' in self.daily_usage_df.columns:
                    if 'Site' not in self.daily_usage_df.columns:
                        self.daily_usage_df.rename(columns={'현장': 'Site'}, inplace=True)
                    else:
                        self.daily_usage_df['Site'] = self.daily_usage_df['Site'].fillna(self.daily_usage_df['현장'])
                        self.daily_usage_df.drop(columns=['현장'], inplace=True)

                if '수량' in self.daily_usage_df.columns:
                    if 'Usage' not in self.daily_usage_df.columns:
                        self.daily_usage_df.rename(columns={'수량': 'Usage'}, inplace=True)
                    else:
                        self.daily_usage_df['Usage'] = self.daily_usage_df['Usage'].fillna(self.daily_usage_df['수량'])
                        self.daily_usage_df.drop(columns=['수량'], inplace=True)

                self.daily_usage_df = self._sync_dataframe_schema(self.daily_usage_df, 'DailyUsage')

                # [NEW] Migrate FilmCount to Usage/수량 if needed (for legacy data)
                if 'FilmCount' in self.daily_usage_df.columns or '필름매수' in self.daily_usage_df.columns:
                    f_col = 'FilmCount' if 'FilmCount' in self.daily_usage_df.columns else '필름매수'
                    u_col = 'Usage' if 'Usage' in self.daily_usage_df.columns else ('수량' if '수량' in self.daily_usage_df.columns else None)
                    
                    if u_col and f_col:
                        # If usage is 0/empty, take FilmCount as fallback for legacy records
                        self.daily_usage_df[u_col] = pd.to_numeric(self.daily_usage_df[u_col], errors='coerce').fillna(0)
                        self.daily_usage_df[f_col] = pd.to_numeric(self.daily_usage_df[f_col], errors='coerce').fillna(0)
                        mask = (self.daily_usage_df[u_col] == 0) & (self.daily_usage_df[f_col] > 0)
                        # self.daily_usage_df.loc[mask, u_col] = self.daily_usage_df.loc[mask, f_col]
                        # [DEFENSIVE] Do NOT migrate if we want to keep them separate now. 
                        # We just keep both columns.
                        pass 

                                                    
                if not self.daily_usage_df.empty:
                    self.daily_usage_df['Date'] = pd.to_datetime(self.daily_usage_df['Date'])
                    self.daily_usage_df['EntryTime'] = pd.to_datetime(self.daily_usage_df['EntryTime'])
                    
                    # Fill NaNs and clean numeric artifacts in string columns
                    string_columns = ['Site', 'Note', '장비명', '검사방법', '업체명', '차량번호', '주행거리', '차량점검', '차량비고']
                    # Add all users, worktimes and OT columns
                    for i in range(1, 11):
                        u_col = 'User' if i == 1 else f'User{i}'
                        w_col = 'WorkTime' if i == 1 else f'WorkTime{i}'
                        o_col = 'OT' if i == 1 else f'OT{i}'
                        string_columns.extend([u_col, w_col, o_col])
                        
                    for col in string_columns:
                        if col not in self.daily_usage_df.columns:
                            self.daily_usage_df[col] = ''
                        
                        # [SAFE] Ensure NO numeric-locked columns receive an empty string by accident
                        # Only apply string replacement if the column is NOT designated as numeric later
                        numeric_intended = ['Usage', '검사량', '단가', '출장비', '일식', '검사비', '수량', 'OT', 'OT금액']
                        for i in range(1, 11): numeric_intended.append(f'OT{i}')
                        
                        if col not in numeric_intended:
                            # Ensure all strings are stripped and nan-free
                            self.daily_usage_df[col] = self.daily_usage_df[col].astype(str).str.strip().replace(['nan', 'None', 'NULL', '-0.0', '0.0', 'NaN', 'NAN', 'nan.0'], '')
                    
                    # Add/Fix columns for auto-calculation
                    if '검사방법' not in self.daily_usage_df.columns and '검사량' in self.daily_usage_df.columns:
                        # Migrate old string '검사량' (PAUT, UT...) to '검사방법'
                        self.daily_usage_df['검사방법'] = self.daily_usage_df['검사량'].astype(str)
                        self.daily_usage_df['검사량'] = 0.0
                    
                    for col in ['검사량', '단가', '출장비', '일식', '검사비', '수량']:
                        if col not in self.daily_usage_df.columns:
                            self.daily_usage_df[col] = 0.0
                        else:
                            # [ROBUST] Handle strings with commas before conversion
                            def clean_num(s):
                                if pd.isna(s): return 0.0
                                try:
                                    # Remove commas and handle currency/unit markers if any
                                    clean_s = str(s).replace(',', '').replace('원', '').strip()
                                    return float(clean_s) if clean_s else 0.0
                                except: return 0.0
                                
                            self.daily_usage_df[col] = self.daily_usage_df[col].apply(clean_num)
                        
                    # [SELF-HEALING] Recover historical Inspection Fees that were saved as 0 due to previous bugs
                    # Inspection Fee = (Amount * Unit Price) + Travel Expense + Meal Cost
                    if not self.daily_usage_df.empty:
                        zero_fee_mask = (self.daily_usage_df['검사비'] == 0) | (self.daily_usage_df['검사비'].isna())
                        if zero_fee_mask.any():
                            # Only recover if we have the inputs
                            can_recover_mask = zero_fee_mask & (self.daily_usage_df['검사량'] > 0) & (self.daily_usage_df['단가'] > 0)
                            if can_recover_mask.any():
                                recovered_fees = (
                                    (self.daily_usage_df.loc[can_recover_mask, '검사량'] * self.daily_usage_df.loc[can_recover_mask, '단가']) + 
                                    self.daily_usage_df.loc[can_recover_mask, '출장비'].fillna(0)
                                )
                                self.daily_usage_df.loc[can_recover_mask, '검사비'] = recovered_fees
                                print(f"DEBUG: Auto-recovered {can_recover_mask.sum()} historical inspection fees.")
                    
                    # [NEW] Compatibility: If '수량' exists but '검사량' is empty/missing, map '수량' to '검사량'
                    if '수량' in self.daily_usage_df.columns:
                         # If both exist, we prefer '수량' if '검사량' is all zeros or NaNs
                         if '검사량' in self.daily_usage_df.columns:
                             mask = (self.daily_usage_df['검사량'] == 0) | (self.daily_usage_df['검사량'].isna())
                             self.daily_usage_df.loc[mask, '검사량'] = self.daily_usage_df.loc[mask, '수량']
                         else:
                             self.daily_usage_df['검사량'] = self.daily_usage_df['수량']
                    
                    # Ensure MaterialID is numeric
                    if 'MaterialID' in self.daily_usage_df.columns:
                        self.daily_usage_df['MaterialID'] = pd.to_numeric(self.daily_usage_df['MaterialID'], errors='coerce')
                    rtk_columns = ['RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 
                                  'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타']
                    for col in rtk_columns:
                        if col not in self.daily_usage_df.columns:
                            self.daily_usage_df[col] = 0.0
                        else:
                            # [CRITICAL] Ensure numeric type for smart hiding aggregation
                            self.daily_usage_df[col] = pd.to_numeric(self.daily_usage_df[col], errors='coerce').fillna(0.0)
                    
                    # Also ensure RTK Category column is removed
                    if 'RTK Category' in self.daily_usage_df.columns:
                        self.daily_usage_df = self.daily_usage_df.drop('RTK Category', axis=1)
                    
            except Exception as e:
                print(f"DEBUG: Failed to load DailyUsage: {e}")
                self.daily_usage_df = pd.DataFrame(columns=['Date', 'Site', 'MaterialID', 'Usage', 'Note', 'EntryTime',
                                                    'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크',
                                                    'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타', 'User', '장비명', '검사량',
                                                    '단가', '출장비', '일식', '검사비',
                                                    '차량번호', '주행거리', '차량점검', '차량비고', '검사구분', '조인트수', '불량수', '관경(Inch)'])
            
            # 5. Migrate vehicle data from daily_usage_df to transactions_df
            if hasattr(self, 'daily_usage_df') and hasattr(self, 'transactions_df'):
                if not self.daily_usage_df.empty and not self.transactions_df.empty:
                    # Group daily usage by date, site, material to match transactions
                    for _, daily_row in self.daily_usage_df.iterrows():
                        # Handle NaT values safely
                        daily_date = pd.to_datetime(daily_row['Date'], errors='coerce')
                        if pd.isna(daily_date):
                            continue  # Skip rows with invalid dates
                        daily_date = daily_date.normalize()
                        daily_site = str(daily_row['Site']).strip()
                        daily_mat = daily_row['MaterialID']
                        
                        # Find matching transactions
                        mask = (
                            (pd.to_datetime(self.transactions_df['Date'], errors='coerce').dt.normalize() == daily_date) &
                            (self.transactions_df['Site'].str.strip() == daily_site) &
                            (self.transactions_df['MaterialID'] == daily_mat)
                        )
                        
                        # Update vehicle info for matching transactions - ensured string conversion to match Transactions_df dtype
                        if mask.any():
                            self.transactions_df.loc[mask, '차량번호'] = str(daily_row.get('차량번호', '')).strip()
                            self.transactions_df.loc[mask, '주행거리'] = str(daily_row.get('주행거리', '')).strip()
                            self.transactions_df.loc[mask, '차량점검'] = str(daily_row.get('차량점검', '')).strip()
                            self.transactions_df.loc[mask, '차량비고'] = str(daily_row.get('차량비고', '')).strip()


            
            # Add SN column if it doesn't exist (for backward compatibility)
            if 'SN' not in self.materials_df.columns:
                self.materials_df['SN'] = ''
            
            # Migrate old schema if needed
            if 'Equipment Code' in self.materials_df.columns and '회사코드' not in self.materials_df.columns:
                self.migrate_old_schema()
            
            # Apply SN extraction from Model Name to existing data
            if not self.materials_df.empty:
                updated = False
                for idx, row in self.materials_df.iterrows():
                    model = row.get('모델명', '')
                    sn = row.get('SN', '')
                    new_model, new_sn = self.extract_sn_from_model(model, sn)
                    
                    if str(model) != str(new_model) or str(sn) != str(new_sn):
                        self.materials_df.at[idx, '모델명'] = new_model
                        self.materials_df.at[idx, 'SN'] = new_sn
                        updated = True
                
                if updated:
                    self.save_data()

            # 6. Budget
            try:
                self.budget_df = pd.read_excel(self.db_path, sheet_name='Budget',
                                               dtype={'Site': str, 'Note': str, 'LaborDetail': str, 'MaterialDetail': str, 'ExpenseDetail': str})
                self.budget_df = normalize_cols(self.budget_df)
                # Add missing detail columns for backward compatibility
                # [FIX] Use 0.0 instead of '' for numeric columns to prevent TypeError: Invalid value '' for dtype 'float64'
                if 'UnitPrice' not in self.budget_df.columns:
                    self.budget_df['UnitPrice'] = 0.0
                if 'LaborDetail' not in self.budget_df.columns:
                    self.budget_df['LaborDetail'] = '{}'
                if 'MaterialDetail' not in self.budget_df.columns:
                    self.budget_df['MaterialDetail'] = '{}'
                if 'ExpenseDetail' not in self.budget_df.columns:
                    self.budget_df['ExpenseDetail'] = '{}'

            except Exception as e:
                print(f"DEBUG: Budget sheet not found or load failed: {e}")
                self.budget_df = pd.DataFrame(columns=['Site', 'Revenue', 'UnitPrice', 'LaborCost', 'MaterialCost', 'Expense', 'OutsourceCost', 'Profit', 'Note', 'LaborDetail', 'MaterialDetail', 'ExpenseDetail'])
            
            # [NEW] 7. Settings (Excel-based Rate Management)
            try:
                self.settings_df = pd.read_excel(self.db_path, sheet_name='Settings')
                self.settings_df = normalize_cols(self.settings_df)
                print("DEBUG: Loaded Settings sheet.")
            except Exception as e:
                print(f"DEBUG: Settings sheet missing or load failed: {e}. Initializing defaults.")
                # Initialize with hardcoded defaults
                labor_defaults = [
                    ['Labor', r, '', '', s] for r, s in {
                        "이사": 55250000, "부장": 55250000, "차장": 47670000, "과장": 41170000,
                        "대리": 37920000, "계장": 34670000, "주임": 31420000, "기사": 29250000
                    }.items()
                ]
                material_defaults = [
                    ['Material', item, spec, unit, price] for item, spec, unit, price in [
                        ("PT 약품", "세척제", "CAN", 1500), ("PT 약품", "침투제", "CAN", 2300),
                        ("PT 약품", "현상제", "CAN", 2000), ("MT 약품", "백색페인트", "CAN", 2350),
                        ("MT 약품", "흑색자분", "CAN", 1800), ("방사선투과검사 필름", "MX125", "매", 990),
                        ("글리세린", "20L", "통", 100000), ("필름 현상액", "3L", "통", 16500),
                        ("필름 정착액", "3L", "통", 16500), ("수적방지액", "200mL", "통", 2500)
                    ]
                ]
                expense_defaults = [
                    ['Expense', '차량유지비', '주유, 수리, 통행, 주차 등', '일', 150000 // 30],
                    ['Expense', '소모품비', '장갑,일회용 작업복외', '일', 15000 // 30],
                    ['Expense', '복리후생비', '생수, 음료 외 기타', '일', 50000 // 30],
                    ['Expense', 'Se-175', '방사성동위원소 구매', 'EA', 10000000 // 280]
                ]
                outsource_defaults = [
                    ['Outsource', '케이엔디이', '방사선투과검사', '공수', 15000]
                ]
                self.settings_df = pd.DataFrame(labor_defaults + material_defaults + expense_defaults + outsource_defaults, 
                                               columns=['Category', 'Name', 'Spec', 'Unit', 'Rate'])
                self.save_data() # Save the newly created sheet

            # [NEW] Ensure budget columns are numeric to prevent NaN assignment errors in some pandas versions
            numeric_cols = ['Revenue', 'UnitPrice', 'LaborCost', 'MaterialCost', 'Expense', 'OutsourceCost', 'Profit']
            for col in numeric_cols:
                if col in self.budget_df.columns:
                    self.budget_df[col] = pd.to_numeric(self.budget_df[col], errors='coerce').fillna(0).astype(float)
    except Exception as e:
        import traceback
        traceback.print_exc()
        self.show_error_dialog("Error", f"데이터를 불러오는데 실패했습니다: {e}")
        self.materials_df = pd.DataFrame(columns=[
            'MaterialID', '회사코드', '관리품번', '품목명', 'SN', '창고',
            '모델명', '규격', '품목군코드', '공급업체', '제조사', '제조국', 
            '가격', '원가', '관리단위', '수량', '재고하한', 'Active'
        ])
        self.transactions_df = pd.DataFrame(columns=['Date', 'MaterialID', 'Site', 'Type', 'Quantity', 'Note', 'User', '차량번호', '주행거리', '차량점검', '차량비고'])
        self.monthly_usage_df = pd.DataFrame(columns=['MaterialID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'Entry Date'])
        self.daily_usage_df = pd.DataFrame(columns=['Date', 'Site', 'MaterialID', 'Usage', 'Note', 'EntryTime',
                                            'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크',
                                            'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타', '장비명', '검사량', '회사코드',
                                            '업체명', '단가', '출장비', '일식', '검사비', 'FilmCount',
                                            '차량번호', '주행거리', '차량점검', '차량비고', '검사구분', '조인트수', '불량수'])
        self.budget_df = pd.DataFrame(columns=['Site', 'Revenue', 'UnitPrice', 'LaborCost', 'MaterialCost', 'Expense', 'OutsourceCost', 'Profit', 'Note', 'LaborDetail', 'MaterialDetail', 'ExpenseDetail'])

    
    # --- [NEW] Global Data Sanitization & Dtype Enforcement ---
    # This definitively prevents "TypeError: Invalid value '' for dtype 'float64'" by ensuring 
    # that all numeric columns are strictly float64 and free of empty strings or text.
    sanitization_map = {
        'budget_df': ['Revenue', 'UnitPrice', 'LaborCost', 'MaterialCost', 'Expense', 'OutsourceCost', 'Profit'],
        'daily_usage_df': ['Usage', '검사량', '단가', '출장비', '일식', '검사비', '수량', 
                            'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 
                            'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타'],
        'materials_df': ['가격', '원가', '수량', '재고하한'],
        'transactions_df': ['Quantity']
    }
    
    for attr, cols in sanitization_map.items():
        if hasattr(self, attr):
            df = getattr(self, attr)
            if df is not None:
                for col in cols:
                    if col in df.columns:
                        # Convert to numeric, turn errors to NaN, then NaN to 0.0, then force float
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float)
                setattr(self, attr, df)

    # Global header cleanup: remove ALL internal and edge whitespace for permanent stability
    for df_attr in ['materials_df', 'transactions_df', 'daily_usage_df', 'budget_df']:
        if hasattr(self, df_attr):
            df = getattr(self, df_attr)
            if df is not None and not df.empty:
                df.columns = [re.sub(r'\s+', '', str(c)) for c in df.columns]
                setattr(self, df_attr, df)
        
    # Refresh all inquiry filters once data is loaded
    try:
        self.refresh_inquiry_filters()
    except Exception as e:
        print(f"DEBUG: Initial refresh_inquiry_filters failed: {e}")


def save_data_impl(self):
    try:
        # [STABILITY] Ensure MaterialID is safely normalized (numeric where possible, otherwise original string)
        def normalize_id(val):
            if pd.isna(val) or str(val).strip() == '': return val
            try:
                s_val = str(val).strip()
                # Handle "10001.0" or "10001"
                num = float(s_val)
                if num == int(num): return int(num)
                return num
            except:
                # Keep as string (for PAUT/manual names)
                return str(val).strip()

        for df_name, df in [('Materials', self.materials_df), ('Transactions', self.transactions_df), 
                            ('Monthly', self.monthly_usage_df), ('Daily', self.daily_usage_df)]:
            if df is not None and 'MaterialID' in df.columns:
                df['MaterialID'] = df['MaterialID'].apply(normalize_id)
            if df is not None and 'Active' in df.columns:
                df['Active'] = pd.to_numeric(df['Active'], errors='coerce').fillna(1).astype(int)
        
        # [STABILITY] Robust saving with Retry and Backup-on-Fail logic
        max_retries = 3
        retry_delay = 0.5
        
        save_success = False
        last_err = None
        
        for attempt in range(max_retries):
            try:
                # Explicitly check for write permission/locks
                if os.path.exists(self.db_path):
                    with open(self.db_path, 'a'): pass
                
                with pd.ExcelWriter(self.db_path, engine='openpyxl') as writer:
                    self.materials_df.to_excel(writer, sheet_name='Materials', index=False)
                    self.transactions_df.to_excel(writer, sheet_name='Transactions', index=False)
                    self.monthly_usage_df.to_excel(writer, sheet_name='MonthlyUsage', index=False)
                    self.daily_usage_df.to_excel(writer, sheet_name='DailyUsage', index=False)
                    self.budget_df.to_excel(writer, sheet_name='Budget', index=False)
                    if hasattr(self, 'settings_df'):
                        self.settings_df.to_excel(writer, sheet_name='Settings', index=False)
                
                save_success = True
                break # Success!
            except PermissionError as pe:
                last_err = pe
                if attempt < max_retries - 1:
                    time.sleep(retry_delay) # Wait and try again
                    continue
                else:
                    # Final attempt failed due to lock. Try saving as Conflict Backup.
                    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    conflict_path = self.db_path.replace('.xlsx', f'_Conflict_{ts}.xlsx')
                    try:
                        with pd.ExcelWriter(conflict_path, engine='openpyxl') as writer:
                            self.materials_df.to_excel(writer, sheet_name='Materials', index=False)
                            self.transactions_df.to_excel(writer, sheet_name='Transactions', index=False)
                            self.monthly_usage_df.to_excel(writer, sheet_name='MonthlyUsage', index=False)
                            self.daily_usage_df.to_excel(writer, sheet_name='DailyUsage', index=False)
                            self.budget_df.to_excel(writer, sheet_name='Budget', index=False)
                            if hasattr(self, 'settings_df'):
                                self.settings_df.to_excel(writer, sheet_name='Settings', index=False)
                        
                        self.show_error_dialog("데이터 저장 지연/충돌", 
                            f"원본 파일('{os.path.basename(self.db_path)}')이 다른 프로그램에 의해 잠겨 있습니다.\n\n"
                            f"데이터 유실 방지를 위해 다음 경로에 임시 저장되었습니다:\n{os.path.basename(conflict_path)}\n\n"
                            f"동기화 중 오류일 수 있으니, 나중에 수동으로 이름을 변경하거나 파일을 병합해 주세요.")
                        return True # Technically "saved" somewhere
                    except Exception as e2:
                         raise Exception(f"원본 파일 잠김 및 백업 저장 실패: {e2}") from pe
            except Exception as e:
                raise e # Other errors

        return save_success
    except Exception as e:
        import traceback
        err_detail = traceback.format_exc()
        self.show_error_dialog("데이터 저장 실패", f"데이터를 저장하는데 실패했습니다:\n{e}\n\n파일이 열려있다면 닫고 다시 시도해주세요.")
        return False


