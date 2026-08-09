import openpyxl
import codecs

wb = openpyxl.load_workbook(r'C:/Users/jjch2/Desktop/템플릿_최종완성본_V74.xlsx', data_only=True)
ws = wb.worksheets[0]

# Dump rows 520-545 with ALL info including merged cell details
out = []
for row in range(520, 548):
    row_info = {'row': row, 'cells': {}, 'merged_anchors': []}
    for col in range(1, 30):
        c = ws.cell(row=row, column=col)
        if c.value is not None:
            row_info['cells'][col] = str(c.value).strip().replace('\n', ' ')
    # Find merged cell anchors in this row
    for merge in ws.merged_cells.ranges:
        if merge.min_row == row:
            row_info['merged_anchors'].append(str(merge.coord))
    out.append(row_info)

with codecs.open('section21_full.txt', 'w', 'utf-8') as f:
    for r in out:
        if r['cells'] or r['merged_anchors']:
            f.write(f"Row {r['row']}: cells={r['cells']} anchors={r['merged_anchors']}\n")
print("Done")
