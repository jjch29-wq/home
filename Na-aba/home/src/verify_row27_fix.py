import os
import openpyxl
from daily_work_report_manager import DailyWorkReportManager
import datetime

# Setup paths
current_dir = os.path.dirname(os.path.abspath(__file__))
home_dir = os.path.dirname(current_dir)
template_path = os.path.join(home_dir, 'resources', 'Template_DailyWorkReport.xlsx')
output_path = os.path.join(current_dir, 'test_row27_fix.xlsx')

manager = DailyWorkReportManager(template_path)

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
    ],
    'methods': {'RT': {'qty': 1, 'price': 1000, 'total': 1000}} # Just some data
}

print(f"Generating report to {output_path}...")
manager.generate_report(data, output_path)

# Verify the result
wb = openpyxl.load_workbook(output_path)
sheet = wb.active

# Find Section 3 Title Row
def find_row_by_text(sheet, text, start_r=1, end_r=100, col='A'):
    for r in range(start_r, end_r + 1):
        val = str(sheet.cell(row=r, column=ord(col.upper())-64).value or "")
        if text in val:
            return r
    return None

s3_row = find_row_by_text(sheet, "3. 차량관리", 10, 60, 'A')
if s3_row:
    print(f"Found '3. 차량관리' at row {s3_row}")
    print(f"Row {s3_row} Merge Analysis:")
    for col_char in ['B', 'E', 'H', 'K', 'N', 'Q']:
        coord = f"{col_char}{s3_row}"
        is_merged = False
        for merged_range in sheet.merged_cells.ranges:
            if coord in merged_range:
                print(f"  {coord}: MERGED in {merged_range}")
                is_merged = True
                break
        if not is_merged:
            print(f"  {coord}: NOT MERGED")
    
    # Check value
    print(f"  Value at A{s3_row}: '{sheet[f'A{s3_row}'].value}'")
    # Check height
    print(f"  Row Height: {sheet.row_dimensions[s3_row].height}")
else:
    print("Could not find '3. 차량관리' row.")

# Cleanup
# os.remove(output_path)
