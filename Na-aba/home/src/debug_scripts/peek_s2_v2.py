import openpyxl
import os

# Use the path from the error but fixed
path = r"C:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Project PROVIDENCE 작업일보(Template).xlsx"

wb = openpyxl.load_workbook(path, data_only=True)
ws = wb.active
for r in range(11, 16):
    print(f'Row {r}:', [c.value for c in ws[r]])
