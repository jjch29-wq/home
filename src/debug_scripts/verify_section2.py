import os
import sys
import datetime
from openpyxl import load_workbook

# Add src to path to import DailyWorkReportManager
sys.path.append(os.getcwd())
from daily_work_report_manager import DailyWorkReportManager

template_path = os.path.abspath("../resources/Template_DailyWorkReport.xlsx")
output_path = os.path.abspath("test_report.xlsx")

# Sample data
data = {
    'date': datetime.date(2026, 5, 2),
    'company': 'Test Company',
    'project_name': 'Test Project',
    'methods': {
        'RT': {'unit': '매', 'qty': 400, 'price': 25500, 'travel': 0, 'total': 10200000},
        'UT': {'unit': '매', 'qty': 100, 'price': 20000, 'travel': 0, 'total': 2000000},
    },
    'vehicles': [{'out_exterior': '양호', 'remarks': 'Test remarks'}]
}

print(f"Generating report using template: {template_path}")
manager = DailyWorkReportManager(template_path)
manager.generate_report(data, output_path)

print(f"Report generated: {output_path}")

# Verify layout
wb = load_workbook(output_path)
sheet = wb.active

def check_row(row_idx, is_edge=False):
    merged_ranges = sheet.merged_cells.ranges
    blocks = [('B','D'), ('E','G'), ('H','J'), ('K','M'), ('N','P'), ('Q','S')]
    print(f"\nChecking Row {row_idx}:")
    for c1, c2 in blocks:
        rng_str = f"{c1}{row_idx}:{c2}{row_idx}"
        is_merged = any(str(r) == rng_str for r in merged_ranges)
        cell = sheet[f"{c1}{row_idx}"]
        
        # Internal lines should be hair
        # Edges should be thin
        left_style = cell.border.left.style
        right_style = sheet[f"{c2}{row_idx}"].border.right.style
        top_style = cell.border.top.style
        
        print(f"  {rng_str}: Merged={is_merged}, Left={left_style}, Right={right_style}, Top={top_style}")

# Check rows 13 to 30
for r in range(13, 31):
    check_row(r)

# Check an empty row in the section
check_row(21)

wb.close()
