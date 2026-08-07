import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Remove FocusOut binding
code = code.replace(
    'self.date_entry.bind("<FocusOut>", self.on_date_change)',
    '# self.date_entry.bind("<FocusOut>", self.on_date_change) # Removed to prevent accidental UI wipes'
)

# 2. Fix NDT Results loading in on_date_change
old_load = """        # Update NDT Results
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
                    ent.insert(0, val)"""

new_load = """        # Update NDT Results
        curr_ndt = curr_data.get('ndt_results', [])
        for i, row_entries in enumerate(self.ndt_grid_entries):
            row_data = curr_ndt[i] if i < len(curr_ndt) else {}
            for col, ent in row_entries.items():
                val = row_data.get(col, '')
                if col == '구간정보':
                    ent.set(val)
                elif isinstance(ent, tk.ttk.Combobox):
                    ent.set(val)
                elif hasattr(ent, 'delete'):
                    ent.delete(0, tk.END)
                    ent.insert(0, val)"""

code = code.replace(old_load, new_load)

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Fixed successfully")
