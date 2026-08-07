import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_loop = """            expected = float(entries['예상량'].get() or 0)
            if expected > 0:"""

new_loop = """            expected = float(entries['예상량'].get() or 0)
            if expected > 0:
                entries['예상량'].delete(0, tk.END)
                entries['예상량'].insert(0, f"{int(expected)}")"""

code = code.replace(old_loop, new_loop)

# Also need to clean up subtotal expected quantities just in case
old_sub = """        paut_total = sum(float(self.qty_entries[f"PAUT_{s}"]['총누계'].get() or 0) for s in ['300A이상', '300A이상-야간', '250A', '200A', '200A-야간'])"""
new_sub = """        paut_expected = sum(float(self.qty_entries[f"PAUT_{s}"]['예상량'].get() or 0) for s in ['300A이상', '300A이상-야간', '250A', '200A', '200A-야간'])
        if paut_expected > 0:
            self.qty_entries['PAUT_소계']['예상량'].delete(0, tk.END)
            self.qty_entries['PAUT_소계']['예상량'].insert(0, f"{int(paut_expected)}")
        paut_total = sum(float(self.qty_entries[f"PAUT_{s}"]['총누계'].get() or 0) for s in ['300A이상', '300A이상-야간', '250A', '200A', '200A-야간'])"""

code = code.replace(old_sub, new_sub)

old_sub2 = """        rt_total = sum(float(self.qty_entries[f"RT_{s}"]['총누계'].get() or 0) for s in ['150A~100A', '150A~100A-야간', '80A이하', '80A이하-야간'])"""
new_sub2 = """        rt_expected = sum(float(self.qty_entries[f"RT_{s}"]['예상량'].get() or 0) for s in ['150A~100A', '150A~100A-야간', '80A이하', '80A이하-야간'])
        if rt_expected > 0:
            self.qty_entries['RT_소계']['예상량'].delete(0, tk.END)
            self.qty_entries['RT_소계']['예상량'].insert(0, f"{int(rt_expected)}")
        rt_total = sum(float(self.qty_entries[f"RT_{s}"]['총누계'].get() or 0) for s in ['150A~100A', '150A~100A-야간', '80A이하', '80A이하-야간'])"""

code = code.replace(old_sub2, new_sub2)

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated expected qty formatting in UI successfully")
