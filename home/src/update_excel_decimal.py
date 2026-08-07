import os

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_loop = """        for method, spec in qty_rows:
            set_cell(f'A{row_idx}', method)
            set_cell(f'B{row_idx}', spec)
            
            # Fetch data from input dictionary
            row_data = data.get('qty_data', {}).get(f"{method}_{spec}", {})
            set_cell(f'C{row_idx}', row_data.get('예상량', ''))
            set_cell(f'D{row_idx}', row_data.get('전일누계', ''))
            set_cell(f'E{row_idx}', row_data.get('금일작업', ''))
            set_cell(f'F{row_idx}', row_data.get('총누계', ''))
            set_cell(f'G{row_idx}', row_data.get('공정률', ''))
            set_cell(f'H{row_idx}', row_data.get('불량', ''))
            set_cell(f'I{row_idx}', row_data.get('불량률', ''))
            set_cell(f'J{row_idx}', row_data.get('비고', ''))
            
            row_idx += 1"""

new_loop = """        for method, spec in qty_rows:
            set_cell(f'A{row_idx}', method)
            set_cell(f'B{row_idx}', spec)
            
            # Fetch data from input dictionary
            row_data = data.get('qty_data', {}).get(f"{method}_{spec}", {})
            
            is_m_unit = method in ['PAUT', 'MT', 'PT']
            
            def set_qty_cell(coord, val_str, force_m_format=False):
                if not val_str:
                    set_cell(coord, '')
                    return
                try:
                    fval = float(str(val_str).replace(',', ''))
                    cell = ws[coord]
                    cell.value = fval
                    if force_m_format:
                        cell.number_format = '0.000'
                    elif fval % 1 != 0:
                        cell.number_format = '0.0'
                    else:
                        cell.number_format = '0'
                    cell.font = self.font_normal
                    cell.alignment = self.align_center
                    cell.border = self.border_thin
                except ValueError:
                    set_cell(coord, val_str)

            set_qty_cell(f'C{row_idx}', row_data.get('예상량', ''), force_m_format=is_m_unit)
            set_qty_cell(f'D{row_idx}', row_data.get('전일누계', ''), force_m_format=is_m_unit)
            set_qty_cell(f'E{row_idx}', row_data.get('금일작업', ''), force_m_format=is_m_unit)
            set_qty_cell(f'F{row_idx}', row_data.get('총누계', ''), force_m_format=is_m_unit)
            
            set_qty_cell(f'G{row_idx}', row_data.get('공정률', ''))
            set_qty_cell(f'H{row_idx}', row_data.get('불량', ''))
            set_qty_cell(f'I{row_idx}', row_data.get('불량률', ''))
            set_cell(f'J{row_idx}', row_data.get('비고', ''))
            
            row_idx += 1"""

code = code.replace(old_loop, new_loop)

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated exporter formatting successfully")
