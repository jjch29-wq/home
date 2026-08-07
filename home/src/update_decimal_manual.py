import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_loop_2 = """            prev_total_str = prev_data.get('qty_data', {}).get(comp_key, {}).get('총누계', '0') or '0'
            try: prev_total = float(prev_total_str.replace(',',''))
            except ValueError: prev_total = 0.0
            today_val = float(entries['금일작업'].get() or 0)
            
            entries['전일누계'].delete(0, tk.END)
            entries['전일누계'].insert(0, format_val(comp_key, prev_total))"""

new_loop_2 = """            prev_total_str = prev_data.get('qty_data', {}).get(comp_key, {}).get('총누계', '0') or '0'
            try: prev_total = float(prev_total_str.replace(',',''))
            except ValueError: prev_total = 0.0
            today_val = float(entries['금일작업'].get() or 0)
            
            entries['금일작업'].delete(0, tk.END)
            entries['금일작업'].insert(0, format_val(comp_key, today_val))
            
            entries['전일누계'].delete(0, tk.END)
            entries['전일누계'].insert(0, format_val(comp_key, prev_total))"""

code = code.replace(old_loop_2, new_loop_2)

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated manual input formatting successfully")
