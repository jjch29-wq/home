import openpyxl
import codecs

wb = openpyxl.load_workbook(r'C:/Users/jjch2/Desktop/템플릿_최종완성본_V74.xlsx', data_only=True)
ws = wb.worksheets[0]

# Check if cleared properly - rows 405~482
out = []
for row in range(403, 490):
    cells = {}
    for col in range(1, 30):
        c = ws.cell(row=row, column=col)
        if c.value is not None:
            v = str(c.value).strip().replace('\n', ' ')
            if v:
                cells[col] = v
    anchors = []
    for merge in ws.merged_cells.ranges:
        if merge.min_row == row:
            anchors.append(str(merge.coord))
    if cells or anchors:
        out.append(f"Row {row}: {cells} | merges_start={anchors}")

with codecs.open('after_clear.txt', 'w', 'utf-8') as f:
    for line in out:
        f.write(line + '\n')

print("Done! Check after_clear.txt")
