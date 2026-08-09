import openpyxl

wb = openpyxl.load_workbook(r'C:/Users/jjch2/Desktop/템플릿_최종완성본_V74.xlsx', data_only=True)
ws = wb.worksheets[0]

# Check rows 520-530 for ALL columns including merged
print("=== 2.1 PAUT 섹션 헤더 (row 519-526) 전체 열 ===")
for row in range(519, 527):
    cells = {}
    for col in range(1, 30):
        c = ws.cell(row=row, column=col)
        val = c.value
        is_merged = isinstance(c, openpyxl.cell.cell.MergedCell)
        if val is not None and not is_merged:
            v = str(val).strip().replace('\n', ' ')
            if v:
                cells[col] = v
    print(f"Row {row}: {cells}")

# Check merge ranges in this area
print("\n=== Row 519-530 병합 범위 ===")
for merge in ws.merged_cells.ranges:
    if 519 <= merge.min_row <= 530:
        print(f"  {merge.coord}")
