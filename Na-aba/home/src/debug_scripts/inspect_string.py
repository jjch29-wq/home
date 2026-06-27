import openpyxl
import os

current_dir = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src"
home_dir = os.path.dirname(current_dir)
template_path = os.path.join(home_dir, 'resources', 'Template_DailyWorkReport.xlsx')

wb = openpyxl.load_workbook(template_path, data_only=True)
sheet = wb.active

val = sheet['E29'].value
print(f"E29 value: [{val}]")
if val:
    print(f"E29 hex: {val.encode('utf-16').hex()}")

val_n = sheet['N29'].value
print(f"N29 value: [{val_n}]")
if val_n:
    print(f"N29 hex: {val_n.encode('utf-16').hex()}")
