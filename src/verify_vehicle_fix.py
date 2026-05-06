import os
from daily_work_report_manager import DailyWorkReportManager

# Setup paths
current_dir = os.path.dirname(os.path.abspath(__file__))
home_dir = os.path.dirname(current_dir)
template_path = os.path.join(home_dir, 'resources', 'Template_DailyWorkReport.xlsx')
output_path = os.path.join(current_dir, 'test_fix_vehicle.xlsx')

manager = DailyWorkReportManager(template_path)

import datetime
data = {
    'date': datetime.date.today(),
    'vehicles': [
        {
            'out_locking': '잠김',
            'in_locking': '잠김',
            'out_exterior': '양호',
            'in_exterior': '양호',
            'out_cleanliness': '양호',
            'in_cleanliness': '양호',
            'out_cleaning': '함',
            'in_cleaning': '함',
            'vehicle_info': 'TEST-1234',
            'mileage': '100',
            'remarks': 'Fix verification test'
        }
    ]
}

print(f"Generating report to {output_path}...")
manager.generate_report(data, output_path)

# Verify the result
import openpyxl
wb = openpyxl.load_workbook(output_path, data_only=True)
sheet = wb.active

def find_row_by_text(sheet, text, start_r=1, end_r=100, col='B'):
    for r in range(start_r, end_r + 1):
        val = str(sheet.cell(row=r, column=ord(col.upper())-64).value or "")
        if text in val:
            return r
    return None

s3_row = find_row_by_text(sheet, "3. 차량관리", 10, 60, 'A') or 18
out_row = find_row_by_text(sheet, "출차시", s3_row, s3_row + 20, 'B')

if out_row:
    locking_val = sheet[f'N{out_row}'].value
    print(f"Result N{out_row}: '{locking_val}'")
    if '■' in locking_val and '잠금' in locking_val:
        print("SUCCESS: '잠금' was correctly checked even though '잠김' was selected!")
    else:
        print("FAILURE: '잠금' was not checked.")
else:
    print("Could not find vehicle section in output.")

os.remove(output_path)
