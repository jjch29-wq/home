import openpyxl

file_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\SISGI-KOGAS-RT-001.xlsx"
wb = openpyxl.load_workbook(file_path)
ws = wb.active

for r in range(28, 60):
    vals = [ws.cell(row=r, column=c).value for c in range(1, 28)]
    if any(v is not None for v in vals):
        print(f"Row {r}: {vals}")
