
import openpyxl
import os

template_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx'

wb = openpyxl.load_workbook(template_path)
sheet = wb.active

print(f"Template Max Row: {sheet.max_row}")
print(f"Template Print Area: {sheet.print_area}")

# List all named ranges in case print area is defined there
for name in wb.defined_names.definedName:
    print(f"Name: {name.name}, Value: {name.value}")
