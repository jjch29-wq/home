import openpyxl

wb = openpyxl.load_workbook(r'C:/Users/jjch2/Desktop/템플릿_최종완성본_V74.xlsx', data_only=True)
ws = wb.worksheets[0]

# check rows 395-415 for all non-None cells
for row in range(395, 415):
    cells = []
    for col in range(1, 30):
        cell = ws.cell(row=row, column=col)
        if cell.value is not None:
            v = str(cell.value).strip().replace('\n', ' ')
            if v:
                b = v.encode('ascii', 'ignore').decode('ascii')
                cells.append(f'C{col}={repr(b)}')
    if cells:
        print(f'Row {row}: {cells}')
