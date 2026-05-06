import openpyxl
import os

current_dir = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src"
home_dir = os.path.dirname(current_dir)
template_path = os.path.join(home_dir, 'resources', 'Template_DailyWorkReport.xlsx')

wb = openpyxl.load_workbook(template_path)
sheet = wb.active

print("Merged cells in Section 3 area (Rows 27-31):")
for merge_range in sheet.merged_cells.ranges:
    if merge_range.min_row >= 27 and merge_range.max_row <= 31:
        print(merge_range)
