import os
import re

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Update save_current_history
new_save = """    def save_current_history(self):
        history = self.load_history()
        current_date = self.date_entry.get()
        
        data = {
            'qty_data': {},
            'equip_data': {},
            'personnel_data': {},
            'remarks': self.remarks_text.get("1.0", "end-1c"),
            'ndt_results': []
        }
        for comp_key, entries in self.qty_entries.items():
            data['qty_data'][comp_key] = {k: v.get() for k, v in entries.items()}
        for eq, entries in self.equip_entries.items():
            data['equip_data'][eq] = {k: v.get() for k, v in entries.items()}
        for p_key, ent in self.personnel_entries.items():
            data['personnel_data'][p_key] = ent.get()
            
        for row_entries in self.ndt_grid_entries:
            row_dict = {}
            for col, ent in row_entries.items():
                if hasattr(ent, 'get'):
                    row_dict[col] = ent.get()
            data['ndt_results'].append(row_dict)
            
        history[current_date] = data
        self.save_history(history)"""

# replace save_current_history
start_idx = code.find("    def save_current_history(self):")
end_idx = code.find("    def on_date_change(self, event=None):")
code = code[:start_idx] + new_save + "\n\n" + code[end_idx:]

# 2. Update 구간정보 set_val
old_set_val = """                    def set_val(val, ents=entries):
                        for e in ents: e.delete(0, tk.END)
                        ents[0].insert(0, val)"""
new_set_val = """                    def set_val(val, ents=entries):
                        for e in ents: e.delete(0, tk.END)
                        parts = val.split(',') if val else []
                        for i, p in enumerate(parts):
                            if i < len(ents): ents[i].insert(0, p)"""
code = code.replace(old_set_val, new_set_val)

# 3. Update on_date_change
new_on_date = """    def on_date_change(self, event=None):
        current_date = self.date_entry.get()
        history = self.load_history()
        
        past_dates = [d for d in history.keys() if d < current_date]
        if past_dates:
            latest_past_date = max(past_dates)
            prev_data = history[latest_past_date]
        else:
            prev_data = {'qty_data': {}, 'equip_data': {}, 'personnel_data': {}}
            
        curr_data = history.get(current_date, {})
            
        # Update Qty
        for comp_key, entries in self.qty_entries.items():
            # Load from today if exists
            curr_qty = curr_data.get('qty_data', {}).get(comp_key, {})
            for field in ['예상량', '금일작업', '총누계', '공정률', '불량', '불량률', '비고']:
                entries[field].delete(0, tk.END)
                if field in curr_qty:
                    entries[field].insert(0, curr_qty[field])
                    
            # Always calculate 전일누계 from past
            prev_total = prev_data.get('qty_data', {}).get(comp_key, {}).get('총누계', '0')
            entries['전일누계'].delete(0, tk.END)
            entries['전일누계'].insert(0, prev_total)
            
        # Update Equip
        for eq, entries in self.equip_entries.items():
            curr_eq = curr_data.get('equip_data', {}).get(eq, {})
            for field in ['금일', '누계']:
                entries[field].delete(0, tk.END)
                if field in curr_eq:
                    entries[field].insert(0, curr_eq[field])
                    
        # Update Personnel
        for p_key, ent in self.personnel_entries.items():
            ent.delete(0, tk.END)
            if p_key in curr_data.get('personnel_data', {}):
                ent.insert(0, curr_data['personnel_data'][p_key])
                
        # Update Remarks
        self.remarks_text.delete("1.0", tk.END)
        if 'remarks' in curr_data:
            self.remarks_text.insert("1.0", curr_data['remarks'])
            
        # Update NDT Results
        curr_ndt = curr_data.get('ndt_results', [])
        for i, row_entries in enumerate(self.ndt_grid_entries):
            row_data = curr_ndt[i] if i < len(curr_ndt) else {}
            for col, ent in row_entries.items():
                val = row_data.get(col, '')
                if hasattr(ent, 'set') and col != '구간정보': 
                    # If it's a combobox, we can just set it
                    if isinstance(ent, tk.ttk.Combobox):
                        ent.set(val)
                    else: # 구간정보 Frame uses custom set
                        ent.set(val)
                elif hasattr(ent, 'delete'):
                    ent.delete(0, tk.END)
                    ent.insert(0, val)
"""

start_idx = code.find("    def on_date_change(self, event=None):")
end_idx = code.find("    def auto_calculate_and_save(self):")
code = code[:start_idx] + new_on_date + "\n" + code[end_idx:]

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated daily_work_log_tab.py successfully")
