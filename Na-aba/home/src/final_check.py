
import openpyxl
import os
import datetime
from daily_work_report_manager import DailyWorkReportManager

template_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx'
output_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src\test_print_area_final.xlsx'

# Generate with 5 RT items
data = {
    'date': datetime.date(2026, 5, 4),
    'materials': {
        f'RT_ROW_{i}': {'name': f'Item {i}', 'used': 1} for i in range(1, 6)
    }
}
manager = DailyWorkReportManager(template_path)
manager.generate_report(data, output_path)

# Verify
wb = openpyxl.load_workbook(output_path)
sheet = wb.active

print(f"Output Max Row: {sheet.max_row}")
print(f"Output Print Area: {sheet.print_area}")
try:
    print(f"FitToWidth: {sheet.page_setup.fitToWidth}")
    print(f"FitToHeight: {sheet.page_setup.fitToHeight}")
except:
    print("PageSetup info missing.")
