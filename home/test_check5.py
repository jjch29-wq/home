import openpyxl, re, os
base = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(base, 'data', 'Report_Template_현장사용량.xlsx')
wb = openpyxl.load_workbook(path)
sheet = wb.active
for coord in ['E29', 'H29', 'K29', 'N29']:
    cell = sheet[coord]
    if isinstance(cell, openpyxl.cell.cell.MergedCell) or cell.value is None:
        for mr in sheet.merged_cells.ranges:
            if coord in mr:
                cell = sheet.cell(row=mr.min_row, column=mr.min_col)
                break
    print(f"[{coord}] Value: {repr(cell.value)}")
    pattern = f"([\u25a1\u2610])(\\s*)양호"
    if cell.value and re.search(pattern, cell.value):
        print(f"[{coord}] Match for '양호'!")
