import openpyxl
import os

path = r"C:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\resources\Template_DailyWorkReport.xlsx"

wb = openpyxl.load_workbook(path, data_only=True)
ws = wb.active
# Cell N29 is Row 29, Column 14 (N)
cell = ws.cell(row=29, column=14)
print(f"Cell N29 value: '{cell.value}'")
print(f"Cell N29 repr: {repr(cell.value)}")
