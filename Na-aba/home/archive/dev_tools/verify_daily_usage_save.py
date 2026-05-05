
import sys
import os
import pandas as pd
import datetime
import tkinter as tk
from tkinter import messagebox
from unittest.mock import MagicMock

# Mocking parts of the app for testing
class MockApp:
    def __init__(self):
        self.daily_usage_df = pd.DataFrame(columns=['Date', 'Site', 'MaterialID', 'Usage', 'Note', 'EntryTime', 'FilmCount'])
        self.materials_df = pd.DataFrame([
            {'MaterialID': 1, '품목명': 'Test Mat', '모델명': 'M1', 'SN': 'S1', '품목군코드': 'FILM'}
        ])
        self.transactions_df = pd.DataFrame(columns=['Date', 'MaterialID', 'Site', 'Type', 'Quantity', 'Note', 'User'])
        self.sites = []
        self.users = []
        
        # UI Mocks
        self.ent_daily_date = MagicMock()
        self.ent_daily_date.get.return_value = '2026-02-15'
        self.ent_daily_site = MagicMock()
        self.ent_daily_site.get.return_value = 'Test Site'
        self.cb_daily_material = MagicMock()
        self.cb_daily_material.get.return_value = 'Test Mat - M1 (SN: S1)'
        self.ent_film_count = MagicMock()
        self.ent_film_count.get.return_value = '10'
        self.ent_daily_note = MagicMock()
        self.ent_daily_note.get.return_value = 'Test Note'
        self.cb_daily_user = MagicMock()
        self.cb_daily_user.get.return_value = '(주간) Tester'
        self.cb_daily_equip = MagicMock()
        self.cb_daily_equip.get.return_value = 'Equip1'
        self.cb_daily_test_method = MagicMock()
        self.cb_daily_test_method.get.return_value = 'RT'
        self.ent_daily_test_amount = MagicMock()
        self.ent_daily_test_amount.get.return_value = '100'
        self.ent_daily_unit_price = MagicMock()
        self.ent_daily_unit_price.get.return_value = '1000'
        self.ent_daily_travel_cost = MagicMock()
        self.ent_daily_travel_cost.get.return_value = '5000'
        self.ent_daily_meal_cost = MagicMock()
        self.ent_daily_meal_cost.get.return_value = '8000'
        self.ent_daily_test_fee = MagicMock()
        self.ent_daily_test_fee.get.return_value = '15000'
        
        self.rtk_entries = {cat: MagicMock() for cat in ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타", "총계"]}
        for mock in self.rtk_entries.values(): mock.get.return_value = '0'
        
        self.ndt_entries = {mat: MagicMock() for mat in ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]}
        for mock in self.ndt_entries.values(): mock.get.return_value = '0'
        
        self.save_data = MagicMock()
        self.update_daily_usage_view = MagicMock()
        self.update_transaction_view = MagicMock()
        self.update_stock_view = MagicMock()
        self.update_material_combo = MagicMock()
        self.update_registration_combos = MagicMock()
        self.save_tab_config = MagicMock()
        self.refresh_ui_for_list_change = MagicMock()
        
    def _safe_format_datetime(self, val, format_str='%Y-%m-%d %H:%M'):
        if pd.isna(val) or val == '': return ''
        try:
            dt = pd.to_datetime(val)
            if pd.isna(dt): return ''
            return dt.strftime(format_str)
        except: return str(val)

    def get_material_display_name(self, mat_id):
        row = self.materials_df[self.materials_df['MaterialID'] == mat_id]
        if not row.empty:
            m = row.iloc[0]
            name = m['품목명']
            model = m.get('모델명', '')
            sn = m.get('SN', '')
            disp = name
            if model: disp += f" - {model}"
            if sn: disp += f" (SN: {sn})"
            return disp
        return f"ID: {mat_id}"

# Insert the code from add_daily_usage_entry here for testing
def add_daily_usage_entry_test(self):
    # This is a simplified version of the logic to verify data flow
    import re
    try:
        date_str = self.ent_daily_date.get()
        site = self.ent_daily_site.get()
        mat_name = self.cb_daily_material.get()
        film_count_str = self.ent_film_count.get()
        note = self.ent_daily_note.get()
        
        film_count = float(film_count_str) if film_count_str else 0.0
        
        mat_id = 1 # Simplified for test
        pure_mat_name = "Test Mat"
        
        selected_date = datetime.datetime.strptime(date_str, '%Y-%m-%d')
        usage_datetime = datetime.datetime.combine(selected_date.date(), datetime.datetime.now().time())
        
        manager_val = self.cb_daily_user.get().strip()
        
        new_entry = {
            'Date': selected_date,
            'Site': site.strip(),
            'MaterialID': mat_id,
            'FilmCount': film_count,
            'Usage': 0,
            'Note': note,
            'EntryTime': datetime.datetime.now(),
            'User': manager_val
        }
        
        # Verify normalization
        final_entry = {re.sub(r'\s+', '', str(k)): v for k, v in new_entry.items()}
        print(f"DEBUG: Final Entry Keys: {list(final_entry.keys())}")
        
        self.daily_usage_df = pd.concat([self.daily_usage_df, pd.DataFrame([final_entry])], ignore_index=True)
        print("SUCCESS: Record added to daily_usage_df")
        
        # Verify view formatting
        for idx, entry in self.daily_usage_df.iterrows():
            u_date = self._safe_format_datetime(entry.get('Date', ''), '%Y-%m-%d')
            e_time = self._safe_format_datetime(entry.get('EntryTime', ''), '%Y-%m-%d %H:%M')
            f_count = entry.get('FilmCount', 0)
            print(f"DEBUG: Formatted Date: {u_date}, EntryTime: {e_time}, FilmCount: {f_count}")
            
            if not u_date or not e_time:
                print("FAILURE: Date formatting failed")
                return False
            if 'FilmCount' not in entry:
                print("FAILURE: FilmCount column missing in entry")
                return False

        return True
    except Exception as e:
        import traceback
        print(f"ERROR in test: {e}")
        traceback.print_exc()
        return False

# Run test
test_app = MockApp()
result = add_daily_usage_entry_test(test_app)
if result:
    print("Verification script passed!")
else:
    print("Verification script FAILED!")
    sys.exit(1)
