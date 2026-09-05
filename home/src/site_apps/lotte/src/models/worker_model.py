import pandas as pd
import numpy as np
from datetime import datetime
import tkinter.messagebox as messagebox
import re
import math
from site_apps.lotte.src.utils.helpers import *

def get_base_salaries_impl(self):
    """Extract labor base salaries from settings_df"""
    if not hasattr(self, 'settings_df') or self.settings_df.empty:
        return {
            "이사": 55250000, "부장": 55250000, "차장": 47670000, "과장": 41170000,
            "대리": 37920000, "계장": 34670000, "주임": 31420000, "기사": 29250000
        }
    df = self.settings_df[self.settings_df['Category'] == 'Labor']
    if df.empty:
        return {
            "이사": 55250000, "부장": 55250000, "차장": 47670000, "과장": 41170000,
            "대리": 37920000, "계장": 34670000, "주임": 31420000, "기사": 29250000
        }
    return df.set_index('Name')['Rate'].to_dict()


def _calculate_ot_from_worktime_impl(self, worktime_value, calculation_date):
    """Standalone helper to calculate OT hours and amount from a worktime string"""
    try:
        if not worktime_value:
            return 0.0, 0
            
        import re
        marker_pattern = MARKER_PATTERN
        clean_val = marker_pattern.sub('', str(worktime_value)).strip()
        
        # Use regex to find separator (~ or -) more robustly
        sep_match = re.search(r'[:\d]\s*([~-])\s*[:\d]', clean_val)
        if not sep_match: return 0.0, 0
        sep = sep_match.group(1)
        
        start_time_str, end_time_str = clean_val.split(sep)
        sh, sm = map(int, start_time_str.split(':'))
        eh, em = map(int, end_time_str.split(':'))
        
        start_f = sh + sm / 60.0
        end_f = eh + em / 60.0
        if end_f < start_f: end_f += 24
        total_duration = end_f - start_f
        if total_duration <= 0: return 0.0, 0

        weekday = calculation_date.weekday()
        is_holiday = weekday >= 5
        is_friday = (weekday == 4)

        ot_hours = 0.0
        amount = 0
        if is_holiday:
            ot_hours = total_duration
            amount = ot_hours * 7500
        else:
            if end_f > 18:
                ot_start = max(start_f, 18.0)
                ot_hours = end_f - ot_start
                evening_end = min(end_f, 22.0)
                evening_hours = max(0, evening_end - ot_start)
                night_start = max(ot_start, 22.0)
                night_end = min(end_f, 24.0)
                night_hours = max(0, night_end - night_start)
                dawn_start = max(ot_start, 24.0)
                dawn_hours = max(0, end_f - dawn_start)
                dawn_rate = 7500 if is_friday else 5000
                amount = (evening_hours * 4000) + (night_hours * 5000) + (dawn_hours * dawn_rate)
        
        return ot_hours, int(amount)
    except:
        return 0.0, 0


def calculate_ot_amount_impl(self, ot_value):
    """Calculate OT amount based on time and rates, or just parse if already an amount"""
    try:
        if not ot_value or not str(ot_value).strip(): return 0
        val = str(ot_value).strip().replace(',', '')
        
        # If it's already just a large number, it's the amount
        if val.isdigit() and int(val) > 100:
            return int(val)
        
        # If it has (N원) format
        if '(' in val and '원)' in val:
            try:
                return int(val.split('(')[1].split('원')[0].replace(',', ''))
            except: pass

        hours = self._parse_ot_hours(val)
        if hours <= 0: return 0

        # (Rest of simple duration fallback)
        evening_hours = min(hours, 4)
        night_hours = max(0, hours - 4)
        return int(evening_hours * 4000 + night_hours * 5000)
    except Exception as e:
        return 0


def _parse_ot_hours_impl(self, ot_value):
    """Helper to extract numeric OT hours using regex for maximum robustness"""
    import re
    try:
        if not ot_value or not str(ot_value).strip(): return 0
        val = str(ot_value).strip().replace(' ', '').replace('익일', '')
        
        # If it's just a large number (>100), assume it's an amount, not hours
        if val.replace(',', '').isdigit() and int(val.replace(',', '')) > 100:
            return 0.0

        # 1. Check for "N시간"
        dur_match = re.search(r'(\d+\.?\d*)\s*(시간|hr|h)', val)
        if dur_match:
            return float(dur_match.group(1))

        # 2. Check for time range "18:00~22:00"
        range_match = re.search(r'(\d{1,2}):(\d{1,2})[-~](\d{1,2}):(\d{1,2})', val)
        if range_match:
            h1, m1, h2, m2 = map(int, range_match.groups())
            if h2 < h1: h2 += 24
            return (h2 * 60 + m2 - (h1 * 60 + m1)) / 60

        # 3. Check for simple ':' format "2:30"
        colon_match = re.search(r'^(\d{1,2}):(\d{1,2})$', val)
        if colon_match:
            h, m = map(int, colon_match.groups())
            return h + (m / 60)

        # 4. Fallback to just extracting the first small float/int found
        num_match = re.search(r'(\d+\.?\d*)', val)
        if num_match:
            v = float(num_match.group(1))
            if v <= 24: return v # Reasonable hour count
        
        return 0
    except:
        return 0


def _calculate_split_ot_hours_impl(self, ot_value, date_val=None):
    """Split OT hours into Day window (18-22), Night window (22-24), and Holiday window for weekends/Friday dawn"""
    import re
    import pandas as pd
    try:
        if not ot_value or not str(ot_value).strip(): return 0.0, 0.0, 0.0
        val = str(ot_value).strip().replace(' ', '')
        
        total_hours = self._parse_ot_hours(val)
        if total_hours <= 0: return 0.0, 0.0, 0.0

        # Default start at 18:00 if no range
        start_hour = 18
        range_match = re.search(r'(\d{1,2}):(\d{1,2})[-~](\d{1,2}):(\d{1,2})', val)
        if range_match:
            start_hour = int(range_match.group(1))
        
        # Handle overnight logic (e.g., 18:00~01:00)
        current_time = float(start_hour)
        remaining = total_hours
        day_hours = 0.0
        night_hours = 0.0
        holiday_hours = 0.0
        
        # Check if this is weekend (Sat/Sun) or Friday going into Saturday
        is_friday = False
        is_weekend = False
        if date_val is not None:
            try:
                dt = pd.to_datetime(date_val)
                is_friday = (dt.weekday() == 4)
                is_weekend = (dt.weekday() >= 5) # Sat=5, Sun=6
            except:
                pass
        
        # [FIXED] If it's Saturday or Sunday, all overtime is holiday work
        if is_weekend:
            return 0.0, 0.0, float(total_hours)

        # Simulate hour by hour (or portion by portion)
        while remaining > 0:
            # 18:00 ~ 22:00 구간은 연장근무 (day_hours)
            if 18 <= current_time < 22:
                can_take = 22 - current_time
                taken = min(remaining, can_take)
                day_hours += taken
                remaining -= taken
                current_time += taken
            # 22:00 ~ 24:00 구간은 야간근무 (night_hours)
            elif 22 <= current_time < 24:
                can_take = 24 - current_time
                taken = min(remaining, can_take)
                night_hours += taken
                remaining -= taken
                current_time += taken
            # 24:00 ~ (익일) 구간
            elif current_time >= 24:
                can_take = remaining # 끝까지 처리
                taken = min(remaining, can_take)
                if is_friday:
                    holiday_hours += taken
                else:
                    night_hours += taken
                remaining -= taken
                current_time += taken
            # 18:00 이전 시간이 혹시 OT로 입력되었다면 상황에 맞게 처리 (기본 시뮬레이션에서는 18시까지 대기하는 것으로 가정)
            else: 
                 current_time = 18.0
        
        return day_hours, night_hours, holiday_hours
    except:
        return 0.0, 0.0, 0.0


def sync_worker_times_impl(self):
    """작업자 1의 설정(주야/작업시간/OT)을 모든 작업자와 동기화"""
    try:
        # Get values from Worker 1
        if not hasattr(self, 'worker_group1'):
            return
            
        master_group = self.worker_group1
        wt1 = master_group.ent_worktime.get().strip()
        ot1 = master_group.ent_ot.get().strip()
        meal1 = master_group.get_meal().strip()
        shift1 = master_group.cb_shift.get()
        
        if not wt1 and not ot1 and not meal1:
            messagebox.showwarning("입력 필요", "작업자 1의 작업시간이나 OT, 또는 일비를 입력해주세요.")
            return

        for i in range(2, 11):
            group_attr = f'worker_group{i}'
            
            # Check if this worker slot exists
            if not hasattr(self, group_attr):
                continue
                
            target_group = getattr(self, group_attr)
            
            # Only apply to workers with names selected
            target_name = target_group.get_worker().strip()
            if not target_name:
                continue

            # Sync Shift
            target_group.cb_shift.set(shift1)

            # Sync Work Time
            target_group.ent_worktime.set(wt1)
            
            # Sync OT
            target_group.set_ot(ot1)
            
            # Sync Meal
            target_group.set_meal(meal1)
        
        messagebox.showinfo("완료", "작업자 1의 설정(주야/시간/OT/일비)이 성명이 입력된 모든 작업자에게 적용되었습니다.")
        
        # [FIX] Prevent RTK grid from being click-blocked after bulk apply & auto-focus
        if hasattr(self, 'rtk_grid') and self.rtk_grid.winfo_ismapped():
            self.rtk_grid.lift()
            if getattr(self, 'cb_daily_test_method', None) and self.cb_daily_test_method.get().strip() == 'RT':
                if hasattr(self, 'rtk_entries') and "센터미스" in self.rtk_entries:
                    self.rtk_entries["센터미스"].focus_set()
                    
    except Exception as e:
        messagebox.showerror("오류", f"동기화 중 오류가 발생했습니다: {e}")


def format_worker_summary_impl(self, workers):
    """Format a list of workers (or a joined string) into a compact summary string with a dropdown cue"""
    if not workers: return ""
    if isinstance(workers, str):
        if " | " in workers:
            names = [w.strip() for w in workers.split(" | ") if w.strip()]
        else:
            names = [w.strip() for w in workers.split(",") if w.strip()]
    else:
        names = [self.clean_nan(w) for w in workers if self.clean_nan(w)]
        
    unique_names = sorted(list(set(names)))
    if not unique_names: return ""
    
    return f"{unique_names[0]} [▼]"


