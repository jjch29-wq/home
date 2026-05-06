import openpyxl
import os

current_dir = os.path.dirname(os.path.abspath(__file__))
home_dir = os.path.dirname(current_dir)
template_path = os.path.join(home_dir, 'resources', 'Template_DailyWorkReport.xlsx')

wb = openpyxl.load_workbook(template_path)
sheet = wb.active
target_text = "3. 차량관리"
found_row = None
for r in range(1, 100):
    val = str(sheet.cell(row=r, column=1).value or "")
    if target_text in val:
        found_row = r
        break

if found_row:
    print(f"'{target_text}' found at row {found_row}")
    print(f"Row {found_row} Merge Analysis:")
    for col_char in ['B', 'E', 'H', 'K', 'N', 'Q']:
        coord = f"{col_char}{found_row}"
        is_merged = False
        for merged_range in sheet.merged_cells.ranges:
            if coord in merged_range:
                print(f"  {coord}: MERGED in {merged_range}")
                is_merged = True
                break
        if not is_merged:
            print(f"  {coord}: NOT MERGED")
else:
    print(f"'{target_text}' NOT found in template.")
