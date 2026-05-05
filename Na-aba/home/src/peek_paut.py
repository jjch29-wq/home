import openpyxl
import os

path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx'
if os.path.exists(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    
    row = 20 # PAUT Row
    print(f"Row {row} (PAUT Row) Cells:")
    for col in range(1, 22): # A to U
        val = ws.cell(row=row, column=col).value
        addr = f"{chr(64+col)}{row}" if col <= 26 else f"A{chr(64+col-26)}{row}"
        print(f"{addr}: {val}")
        
    # Check for merging
    for merge_range in ws.merged_cells.ranges:
        if f'A{row}' in merge_range or f'B{row}' in merge_range:
            print(f"Merged Range involving row {row}: {merge_range}")
else:
    print("File not found")
