import openpyxl
import os

path = r"C:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\resources\Template_DailyWorkReport.xlsx"

wb = openpyxl.load_workbook(path, data_only=True)
ws = wb.active
cell = ws.cell(row=29, column=14)
val = cell.value if cell.value else ""
print(f"Hex: {' '.join([hex(ord(c)) for c in val])}")
# Try to decode to Korean
try:
    print(f"Decoded: {val}")
except:
    pass
