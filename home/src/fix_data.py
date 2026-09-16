with open(r'C:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src\비파괴검사보고서.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if 'no_col = self.col_to_num(self.config.get(\'PT_COL_NO\', \'1\'))' in line:
        skip = True
        
        indent = '                '
        new_lines.append(indent + "# PT 고정 열 하드코딩 (H:Q 병합 주의)\n")
        new_lines.append(indent + "self.safe_set_value(ws, ws.cell(row=current_row, column=1).coordinate, item.get('Dwg', ''))\n")
        new_lines.append(indent + "self.safe_set_value(ws, ws.cell(row=current_row, column=4).coordinate, item.get('Joint', ''))\n")
        new_lines.append(indent + "\n")
        new_lines.append(indent + "res_val = str(item.get('Result', 'Acc')).strip().upper()\n")
        new_lines.append(indent + "is_acc = (res_val in ['ACC', 'ACCEPT', 'PASS', '합격', 'O', 'OK', 'V', ''])\n")
        new_lines.append(indent + "\n")
        new_lines.append(indent + "if is_acc:\n")
        new_lines.append(indent + "    self.safe_set_value(ws, ws.cell(row=current_row, column=5).coordinate, 'V')\n")
        new_lines.append(indent + "    self.safe_set_value(ws, ws.cell(row=current_row, column=8).coordinate, 'NO RECORDABLE INDICATION')\n")
        new_lines.append(indent + "else:\n")
        new_lines.append(indent + "    self.safe_set_value(ws, ws.cell(row=current_row, column=6).coordinate, 'V')\n")
        new_lines.append(indent + "    self.safe_set_value(ws, ws.cell(row=current_row, column=8).coordinate, res_val)\n")
        new_lines.append(indent + "\n")
        new_lines.append(indent + "self.safe_set_value(ws, ws.cell(row=current_row, column=18).coordinate, item.get('Welder', ''))\n")
        new_lines.append(indent + "self.safe_set_value(ws, ws.cell(row=current_row, column=19).coordinate, item.get('NPS', ''))\n")
        new_lines.append(indent + "\n")
        new_lines.append(indent + "# 스타일링\n")
        new_lines.append(indent + "for c in range(1, 20):\n")

    if skip and 'for c in range(1, 12):' in line:
        skip = False
        continue
        
    if not skip:
        new_lines.append(line)

with open(r'C:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src\비파괴검사보고서.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print('Replaced data block!')
