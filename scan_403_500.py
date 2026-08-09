import openpyxl
import codecs

# Load the original template
wb = openpyxl.load_workbook(r'C:/Users/jjch2/Desktop/템플릿_최종완성본_V74.xlsx', data_only=True)
ws = wb.worksheets[0]

# Scan ALL rows from 403 to 500 to see the full structure
out = []
for row in range(403, 500):
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
    out.append((row, cells, anchors))

with codecs.open('full_403_500.txt', 'w', 'utf-8') as f:
    for row, cells, anchors in out:
        if cells or anchors:
            f.write(f"Row {row}: {cells}  |  merges={anchors}\n")
        else:
            f.write(f"Row {row}: (empty)\n")

print("Done! Check full_403_500.txt")
