
import openpyxl
import os
import datetime
from daily_work_report_manager import DailyWorkReportManager

template_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx'
output_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src\test_print_area.xlsx'

def check_print_info(path, label):
    wb = openpyxl.load_workbook(path)
    sheet = wb.active
    print(f"--- {label} ---")
    print(f"File: {path}")
    print(f"Max Row: {sheet.max_row}")
    print(f"Print Area: {sheet.print_area}")
    if sheet.page_setup:
        print(f"Fit to Page: {sheet.page_setup.fitToPage}")
        print(f"Fit to Width: {sheet.page_setup.fitToWidth}")
        print(f"Fit to Height: {sheet.page_setup.fitToHeight}")
    print("")

# 1. Check Template
check_print_info(template_path, "TEMPLATE")

# 2. Generate with 5 RT items
data = {
    'date': datetime.date(2026, 5, 4),
    'materials': {
        f'RT_ROW_{i}': {'name': f'Item {i}', 'used': 1} for i in range(1, 6)
    }
}
manager = DailyWorkReportManager(template_path)
manager.generate_report(data, output_path)

# 3. Check Output
check_print_info(output_path, "OUTPUT (5 Items)")
