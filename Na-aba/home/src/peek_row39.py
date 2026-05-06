import openpyxl
import os

path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx'
if os.path.exists(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    
    row = 39
    print(f"Row {row} Cells:")
    for col in range(1, 20):
        val = ws.cell(row=row, column=col).value
        print(f"{chr(64+col)}{row}: {val}")
else:
    print("File not found")
