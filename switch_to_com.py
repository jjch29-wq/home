import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i in range(11960, 11980):
    if 'import openpyxl' in lines[i]:
        lines[i] = "                import win32com.client\n                import os\n                excel = win32com.client.Dispatch('Excel.Application')\n                excel.Visible = False\n"
        lines[i+1] = "                wb = excel.Workbooks.Open(os.path.abspath(save_path))\n"
        break

for i in range(11970, 11990):
    if 'ws.cell(row=row, column=col, value=' in lines[i]:
        lines[i] = lines[i].replace('ws.cell(row=row, column=col, value=', 'ws.Cells(row, col).Value = ')
        # The closing parenthesis of ws.cell needs to be removed.
        # But wait, it's `value=round(...) if ... else ...)`
        # Let's just do a regex replace on that line.
        import re
        lines[i] = re.sub(r'ws\.cell\(row=row,\s*column=col,\s*value=(.*)\)', r'ws.Cells(row, col).Value = \1', lines[i])
        break

for i in range(12080, 12110):
    if 'if sheet_name in wb.sheetnames:' in lines[i]:
        # Need to insert `sheet_names = [sheet.Name for sheet in wb.Sheets]` before `if main_agg:`
        pass

# Find `if main_agg:` and insert sheet_names
for i in range(12080, 12110):
    if 'if main_agg:' in lines[i]:
        lines.insert(i, "                sheet_names = [sheet.Name for sheet in wb.Sheets]\n")
        break

# Find and replace wb.sheetnames and wb[sheet_name]
for i in range(12080, 12120):
    if 'if sheet_name in wb.sheetnames:' in lines[i]:
        lines[i] = lines[i].replace('wb.sheetnames', 'sheet_names')
    if 'write_ndt_sheet(wb[sheet_name]' in lines[i]:
        lines[i] = lines[i].replace('wb[sheet_name]', 'wb.Sheets(sheet_name)')

for i in range(12100, 12130):
    if 'wb.save(save_path)' in lines[i]:
        lines[i] = lines[i].replace('wb.save(save_path)', 'wb.Save()')
    if 'wb.close()' in lines[i]:
        lines[i] = lines[i].replace('wb.close()', 'wb.Close()\n                excel.Quit()')

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
