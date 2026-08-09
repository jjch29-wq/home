import openpyxl
import codecs

wb = openpyxl.load_workbook(r'C:/Users/jjch2/Desktop/템플릿_최종완성본_V74.xlsx', data_only=True)
ws = wb.worksheets[0]

# Dump all rows from 486 to 600 to see all sections
out = []
for row in range(486, 605):
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
            anchors.append(f"{merge.coord}")
    if cells or anchors:
        out.append((row, cells, anchors))

with codecs.open('all_ndt_sections.txt', 'w', 'utf-8') as f:
    for row, cells, anchors in out:
        f.write(f"Row {row}: {cells}  |  merges={anchors}\n")

print("Done! Check all_ndt_sections.txt")
