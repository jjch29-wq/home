import re

file_path = 'src/Material-Master-Manager-V13.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Update the combobox to use the dynamic list
content = content.replace(
    "self.cb_daily_unit = ttk.Combobox(form_content, width=13, values=['매', 'P,M,I/D', 'M,I/D', 'Point', 'Meter', 'Inch', 'Dia'])",
    "self.cb_daily_unit = ttk.Combobox(form_content, width=13, values=self.daily_units)"
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print('Combobox updated to use dynamic daily_units list.')
