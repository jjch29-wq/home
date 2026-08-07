import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_save = """        for row_entries in self.ndt_grid_entries:
            row_dict = {}
            for col, ent in row_entries.items():
                if hasattr(ent, 'get'):
                    row_dict[col] = ent.get()
            data['ndt_results'].append(row_dict)"""

new_save = """        for row_entries in self.ndt_grid_entries:
            row_dict = {}
            for col, ent in row_entries.items():
                if hasattr(ent, 'get'):
                    val = ent.get()
                    if callable(val): # In case get() returned a method? No, ent.get is the method.
                        pass
                    row_dict[col] = str(val) if val else ""
            data['ndt_results'].append(row_dict)
            
        with open('debug_save.txt', 'w', encoding='utf-8') as debug_f:
            import json
            json.dump(data['ndt_results'], debug_f, ensure_ascii=False, indent=2)"""

code = code.replace(old_save, new_save)

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Debug added")
