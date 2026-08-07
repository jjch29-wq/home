import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_load = """        for comp_key, entries in self.qty_entries.items():
            # Load from today if exists
            curr_qty = curr_data.get('qty_data', {}).get(comp_key, {})
            for field in ['예상량', '금일작업', '총누계', '공정률', '불량', '불량률', '비고']:
                entries[field].delete(0, tk.END)
                if field in curr_qty:
                    entries[field].insert(0, curr_qty[field])"""

new_load = """        def load_format_val(ckey, field_name, v):
            if not str(v).strip(): return v
            if field_name == '예상량':
                try: return f"{int(float(str(v).replace(',','')))}"
                except ValueError: return v
            if field_name in ['전일누계', '금일작업', '총누계']:
                try:
                    fv = float(str(v).replace(',', ''))
                    if ckey.startswith(('PAUT', 'MT', 'PT')): return f"{fv:.4f}"
                    return f"{fv:.1f}" if fv % 1 else f"{int(fv)}"
                except ValueError: return v
            return v

        for comp_key, entries in self.qty_entries.items():
            # Load from today if exists
            curr_qty = curr_data.get('qty_data', {}).get(comp_key, {})
            for field in ['예상량', '금일작업', '총누계', '공정률', '불량', '불량률', '비고']:
                entries[field].delete(0, tk.END)
                if field in curr_qty:
                    val = load_format_val(comp_key, field, curr_qty[field])
                    entries[field].insert(0, val)"""

code = code.replace(old_load, new_load)

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated load formatting successfully")
