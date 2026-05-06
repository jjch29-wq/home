import openpyxl
import os

path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx"

if os.path.exists(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = wb.active
    for r in range(17, 27):
        row_data = [sheet.cell(row=r, column=c).value for c in range(1, 10)]
        print(f"Row {r}: {row_data}")
