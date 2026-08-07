import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Fix 1: Add weather to save_current_history
old_save = """        data = {
            'qty_data': {},
            'equip_data': {},
            'personnel_data': {},
            'remarks': self.remarks_text.get("1.0", "end-1c"),
            'ndt_results': []
        }"""
new_save = """        data = {
            'weather': self.weather_entry.get(),
            'qty_data': {},
            'equip_data': {},
            'personnel_data': {},
            'remarks': self.remarks_text.get("1.0", "end-1c"),
            'ndt_results': []
        }"""
code = code.replace(old_save, new_save)

# Fix 2: Add weather loading to on_date_change
old_load = """        # Update Remarks
        self.remarks_text.delete("1.0", tk.END)
        if 'remarks' in curr_data:
            self.remarks_text.insert("1.0", curr_data['remarks'])"""

new_load = """        # Update Weather
        self.weather_entry.delete(0, tk.END)
        if 'weather' in curr_data:
            self.weather_entry.insert(0, curr_data['weather'])
            
        # Update Remarks
        self.remarks_text.delete("1.0", tk.END)
        if 'remarks' in curr_data:
            self.remarks_text.insert("1.0", curr_data['remarks'])"""

code = code.replace(old_load, new_load)

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated weather save/load logic successfully")
