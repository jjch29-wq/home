
import openpyxl
import os

template_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx'
wb = openpyxl.load_workbook(template_path)
sheet = wb.active

# Section 6 title is usually at 41
start = 41
for r in range(start, start + 15):
    row_values = [sheet.cell(row=r, column=c).value for c in range(1, 10)]
    print(f"Row {r:2}: {row_values}")
