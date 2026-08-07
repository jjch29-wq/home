import os

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_line = "set_qty_cell(f'C{row_idx}', row_data.get('예상량', ''), force_m_format=is_m_unit)"
new_line = "set_qty_cell(f'C{row_idx}', row_data.get('예상량', ''), force_m_format=False)"

code = code.replace(old_line, new_line)

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated expected qty formatting in exporter successfully")
