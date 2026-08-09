import openpyxl

wb = openpyxl.load_workbook(r'C:/Users/jjch2/Desktop/템플릿_최종완성본_V74.xlsx', data_only=True)
ws = wb.worksheets[0]

# Dump header rows with their Korean content to file
import codecs

with codecs.open('ndt_headers.txt', 'w', 'utf-8') as out:
    # Check rows 483-530 for MT/RT NDT tables
    for row_num in [403, 487, 504, 511, 523, 556, 565, 572]:
        out.write(f"\n=== Row {row_num} ===\n")
        # Also dump row+1 (sub-headers)
        for r in range(row_num, row_num + 3):
            cells = {}
            for col in range(1, 30):
                c = ws.cell(row=r, column=col)
                if c.value and isinstance(c.value, str):
                    cv = c.value.strip().replace('\n', ' ')
                    if cv:
                        cells[col] = cv
            out.write(f"  Row {r}: {cells}\n")
        
print("Done! Check ndt_headers.txt")
