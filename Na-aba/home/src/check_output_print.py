
import openpyxl
import os

output_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src\test_print_area.xlsx'
wb = openpyxl.load_workbook(output_path)
sheet = wb.active

print(f"Output Max Row: {sheet.max_row}")
print(f"Output Print Area: {sheet.print_area}")

try:
    print(f"FitToWidth: {sheet.page_setup.fitToWidth}")
    print(f"FitToHeight: {sheet.page_setup.fitToHeight}")
except:
    print("PageSetup info not accessible or not set.")
