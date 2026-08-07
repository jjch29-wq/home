import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_loop_1 = """        for comp_key, val in today_qty.items():
            if '소계' not in comp_key:
                ent = self.qty_entries[comp_key]['금일작업']
                ent.delete(0, tk.END)
                ent.insert(0, f"{val:.1f}" if val % 1 else f"{int(val)}")"""

new_loop_1 = """        def format_val(ckey, v):
            if ckey.startswith(('PAUT', 'MT', 'PT')):
                return f"{v:.3f}"
            return f"{v:.1f}" if v % 1 else f"{int(v)}"

        for comp_key, val in today_qty.items():
            if '소계' not in comp_key:
                ent = self.qty_entries[comp_key]['금일작업']
                ent.delete(0, tk.END)
                ent.insert(0, format_val(comp_key, val))"""

old_loop_2 = """            entries['전일누계'].delete(0, tk.END)
            entries['전일누계'].insert(0, f"{prev_total:.1f}" if prev_total % 1 else f"{int(prev_total)}")
            
            total = prev_total + today_val
            entries['총누계'].delete(0, tk.END)
            entries['총누계'].insert(0, f"{total:.1f}" if total % 1 else f"{int(total)}")"""

new_loop_2 = """            entries['전일누계'].delete(0, tk.END)
            entries['전일누계'].insert(0, format_val(comp_key, prev_total))
            
            total = prev_total + today_val
            entries['총누계'].delete(0, tk.END)
            entries['총누계'].insert(0, format_val(comp_key, total))"""

old_loop_3 = """        paut_total = sum(float(self.qty_entries[f"PAUT_{s}"]['총누계'].get() or 0) for s in ['300A이상', '300A이상-야간', '250A', '200A', '200A-야간'])
        self.qty_entries['PAUT_소계']['총누계'].delete(0, tk.END)
        self.qty_entries['PAUT_소계']['총누계'].insert(0, f"{paut_total:.1f}" if paut_total % 1 else f"{int(paut_total)}")
        
        rt_total = sum(float(self.qty_entries[f"RT_{s}"]['총누계'].get() or 0) for s in ['150A~100A', '150A~100A-야간', '80A이하', '80A이하-야간'])
        self.qty_entries['RT_소계']['총누계'].delete(0, tk.END)
        self.qty_entries['RT_소계']['총누계'].insert(0, f"{rt_total:.1f}" if rt_total % 1 else f"{int(rt_total)}")"""

new_loop_3 = """        paut_total = sum(float(self.qty_entries[f"PAUT_{s}"]['총누계'].get() or 0) for s in ['300A이상', '300A이상-야간', '250A', '200A', '200A-야간'])
        self.qty_entries['PAUT_소계']['총누계'].delete(0, tk.END)
        self.qty_entries['PAUT_소계']['총누계'].insert(0, format_val('PAUT', paut_total))
        
        rt_total = sum(float(self.qty_entries[f"RT_{s}"]['총누계'].get() or 0) for s in ['150A~100A', '150A~100A-야간', '80A이하', '80A이하-야간'])
        self.qty_entries['RT_소계']['총누계'].delete(0, tk.END)
        self.qty_entries['RT_소계']['총누계'].insert(0, format_val('RT', rt_total))"""

code = code.replace(old_loop_1, new_loop_1)
code = code.replace(old_loop_2, new_loop_2)
code = code.replace(old_loop_3, new_loop_3)

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated decimal formatting successfully")
