import openpyxl, re
path = r"f:/내 드라이브/07_Antigravity/PMI_한국지역난방/home/data/Report_Template_현장사용량.xlsx"
wb = openpyxl.load_workbook(path)
sheet = wb.active

chk_rows = {'out': 29, 'in': 30}
chk_map = {'exterior':'E','cleanliness':'H','cleaning':'K','locking':'N'}
v = {'out_exterior': '양호', 'out_cleanliness': '불량', 'out_cleaning': '함', 'out_locking': '잠금'}

for rk, row_idx in chk_rows.items():
    for ck, col_let in chk_map.items():
        val = v.get(f"{rk}_{ck}")
        if val:
            coord = f"{col_let}{row_idx}"
            cell = sheet[coord]
            if isinstance(cell, openpyxl.cell.cell.MergedCell) or cell.value is None:
                for mr in sheet.merged_cells.ranges:
                    if coord in mr:
                        cell = sheet.cell(row=mr.min_row, column=mr.min_col)
                        break
            
            if cell.value and isinstance(cell.value, str):
                print(f"[{coord}] Original: {repr(cell.value)}")
                pattern = f"([\u25a1\u2610\u3141]|\[\s*\])(\\s*){re.escape(val)}"
                if re.search(pattern, cell.value):
                    new_val = re.sub(pattern, f"\u25a0\\2{val}", cell.value)
                    print(f"[{coord}] REPLACED: {repr(new_val)}")
                else:
                    print(f"[{coord}] NO MATCH for val: {val} with pattern: {pattern}")
