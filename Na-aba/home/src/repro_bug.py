import os
import datetime
from daily_work_report_manager import DailyWorkReportManager
import openpyxl

# Setup
current_dir = os.path.dirname(os.path.abspath(__file__))
home_dir = os.path.dirname(current_dir)
template_path = os.path.join(home_dir, 'resources', 'Template_DailyWorkReport.xlsx')
output_path = os.path.join(current_dir, 'test_output_bug.xlsx')

manager = DailyWorkReportManager(template_path)

# Data with 5 methods (triggering 1 insertion)
data = {
    'date': datetime.date.today(),
    'methods': {
        'RT': {'qty': 10}, 'UT': {'qty': 20}, 'MT': {'qty': 30}, 'PT': {'qty': 40}, 'PAUT': {'qty': 50}
    },
    'rtk': {'센터미스': 1, '농도': 2, '총계': 3},
    'vehicles': [{'out_locking': '잠김', 'remarks': 'Test Remarks'}]
}

print("Generating report...")
manager.generate_report(data, output_path)

print(f"--- Output File Rows 25-35 ---")
wb = openpyxl.load_workbook(output_path, data_only=True)
sheet = wb.active
for r in range(25, 36):
    row_vals = [str(sheet.cell(row=r, column=c).value or "").strip() for c in range(1, 6)]
    print(f"Row {r}: {' | '.join(row_vals)}")

os.remove(output_path)
