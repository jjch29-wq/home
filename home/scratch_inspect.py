import openpyxl
import os

wb = openpyxl.load_workbook('c:/Users/-/PMI/home/data/monthly_report_template.xlsx')
ws = wb.active
print('Print titles:', ws.print_title_rows)

start = None
end = None
for r in range(1, 200):
    val = str(ws.cell(row=r, column=2).value or '').replace(' ', '')
    if '2.0안전관리교육' in val or '2.0안전관리' in val:
        start = r
    if start and '3.0' in val:
        end = r
        break

print('Start:', start, 'End:', end)
