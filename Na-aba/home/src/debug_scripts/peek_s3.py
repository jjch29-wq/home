import openpyxl
import os

path = r"C:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Project PROVIDENCE 작업일보(Template).xlsx"

wb = openpyxl.load_workbook(path, data_only=True)
ws = wb.active
# Section 3 is around Row 26-30
for r in range(26, 31):
    print(f'Row {r}:', [c.value for c in ws[r]])
