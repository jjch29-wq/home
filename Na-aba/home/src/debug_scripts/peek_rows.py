import openpyxl
import os

current_dir = os.path.dirname(os.path.abspath(__file__))
home_dir = os.path.dirname(current_dir)
template_path = os.path.join(home_dir, 'resources', 'Template_DailyWorkReport.xlsx')

wb = openpyxl.load_workbook(template_path, data_only=True)
sheet = wb.active

print(f"--- Template Rows 18-25 ---")
for r in range(18, 26):
    row_vals = [str(sheet.cell(row=r, column=c).value or "").strip() for c in range(1, 10)]
    print(f"Row {r}: {' | '.join(row_vals)}")
