import openpyxl
import os

path = r"C:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\resources\Template_DailyWorkReport.xlsx"

wb = openpyxl.load_workbook(path, data_only=True)
ws = wb.active
for r in range(11, 16):
    print(f'Row {r}:', [c.value for c in ws[r]])
