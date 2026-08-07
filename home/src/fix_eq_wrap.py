import os

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Define align_nowrap
old_align = """        self.align_right = Alignment(horizontal='right', vertical='center', wrap_text=True)"""
new_align = """        self.align_right = Alignment(horizontal='right', vertical='center', wrap_text=True)
        self.align_nowrap = Alignment(horizontal='center', vertical='center', wrap_text=False, shrink_to_fit=True)"""
code = code.replace(old_align, new_align)

# Update equip rows
old_eq = """        for i, eq in enumerate(equip_rows):
            r = 10 + i
            set_cell(f'L{r}', eq)
            eq_data = data.get('equip_data', {}).get(eq, {})"""

new_eq = """        for i, eq in enumerate(equip_rows):
            r = 10 + i
            set_cell(f'L{r}', eq, align=self.align_nowrap)
            eq_data = data.get('equip_data', {}).get(eq, {})"""

code = code.replace(old_eq, new_eq)

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated equipment alignment successfully")
