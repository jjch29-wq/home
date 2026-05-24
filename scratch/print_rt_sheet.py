import openpyxl

file_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\SISGI-KOGAS-RT-001.xlsx"
wb = openpyxl.load_workbook(file_path)
ws = wb['RT()']

for r in range(8, 36):
    vals = [ws.cell(row=r, column=c).value for c in range(1, 28)]
    print(f"Row {r}: {vals}")
