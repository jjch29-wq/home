import openpyxl
import os

path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx'
if os.path.exists(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    
    print("Range K43:S50 Values:")
    for row in range(43, 51):
        row_vals = []
        for col_idx in range(11, 20): # K is 11, S is 19
            val = ws.cell(row=row, column=col_idx).value
            row_vals.append(f"{chr(64+col_idx)}{row}: {val}")
        print(" | ".join(row_vals))
else:
    print("File not found")
