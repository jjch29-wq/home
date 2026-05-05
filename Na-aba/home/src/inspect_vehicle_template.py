import openpyxl
import os

current_dir = os.path.dirname(os.path.abspath(__file__))
home_dir = os.path.dirname(current_dir) # current_dir is home/src, so dirname is home
template_path = os.path.join(home_dir, 'resources', 'Template_DailyWorkReport.xlsx')

print(f"Inspecting {template_path}")
if not os.path.exists(template_path):
    print("Template not found!")
    exit()

wb = openpyxl.load_workbook(template_path, data_only=True)
sheet = wb.active

# Find the vehicle section
# Based on daily_work_report_manager.py:
# chk_header_row = find_row_by_text("차량 외부상태", ...)
# out_row_base = find_row_by_text("출차시", ...)

def find_row_by_text(sheet, text, start_r=1, end_r=100, col='B'):
    for r in range(start_r, end_r + 1):
        val = str(sheet.cell(row=r, column=ord(col.upper())-64).value or "")
        if text in val:
            return r
    return None

s3_title_row = find_row_by_text(sheet, "3. 차량관리", 10, 60, 'A') or 18
chk_header_row = find_row_by_text(sheet, "차량 외부상태", s3_title_row, s3_title_row + 20, 'E')
print(f"S3 Title Row: {s3_title_row}")
print(f"Check Header Row: {chk_header_row}")

if chk_header_row:
    for col in ['B', 'E', 'H', 'K', 'N']:
        val = sheet[f'{col}{chk_header_row}'].value
        print(f"Header {col}{chk_header_row}: '{val}'")

# Check rows for 'out' and 'in'
out_row = find_row_by_text(sheet, "출차시", s3_title_row, s3_title_row + 20, 'B')
in_row = find_row_by_text(sheet, "입차시", out_row if out_row else s3_title_row, s3_title_row + 20, 'B')

print(f"Out row: {out_row}, In row: {in_row}")

# Check the locking device column (N)
def print_cell_info(cell_coord):
    val = sheet[cell_coord].value
    if val:
        print(f"{cell_coord}: '{val}' | Hex: {str(val).encode('utf-8').hex()}")
    else:
        print(f"{cell_coord}: None")

print("\n--- Locking (Column N) ---")
if out_row: print_cell_info(f'N{out_row}')
if in_row: print_cell_info(f'N{in_row}')

print("\n--- Cleaning (Column K) ---")
if out_row: print_cell_info(f'K{out_row}')
if in_row: print_cell_info(f'K{in_row}')
