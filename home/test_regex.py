import openpyxl
import re

wb = openpyxl.Workbook()
sheet = wb.active

# Setup mock data
sheet['C29'] = "□ 양호 □ 불량"
sheet['G29'] = "□ 양호 □ 불량"
sheet['K29'] = "□ 함 □ 안함"
sheet['O29'] = "□ 잠금 □ 안함"

# Mock variables
v = {'out_exterior': '양호', 'out_cleanliness': '불량', 'out_cleaning': '함', 'out_locking': '잠금'}
chk_rows = {'out': 29}
chk_map_ranges = {
    'exterior': range(3, 7),
    'cleanliness': range(7, 11),
    'cleaning': range(11, 15),
    'locking': range(15, 20)
}

for rk, row_idx in chk_rows.items():
    for ck, col_range in chk_map_ranges.items():
        val = v.get(f"{rk}_{ck}")
        if val:
            pattern = f"([\u25a1\u2610\u3141]|\\[\\s*\\]|\\(\\s*\\))(\\s*){re.escape(val)}"
            for c in col_range:
                cell = sheet.cell(row=row_idx, column=c)
                if cell.value and isinstance(cell.value, str):
                    if re.search(pattern, cell.value):
                        cell.value = re.sub(pattern, f"\u25a0\\2{val}", cell.value)
                        print(f"[{ck}] REPLACED: {cell.value}")
                        break

print("Final values:")
print(f"C29: {sheet['C29'].value}")
print(f"G29: {sheet['G29'].value}")
print(f"K29: {sheet['K29'].value}")
print(f"O29: {sheet['O29'].value}")
