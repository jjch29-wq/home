import sys, json, openpyxl
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src')

# Check Debug_NDT.xlsx that was just saved
wb = openpyxl.load_workbook(r'C:\Users\jjch2\Desktop\Debug_NDT.xlsx', data_only=True)
ws = wb.worksheets[0]

print("=== PAUT 섹션 (row 403-425) ===")
for row in range(403, 425):
    cells = {}
    for col in range(1, 25):
        c = ws.cell(row=row, column=col)
        if c.value is not None:
            v = str(c.value).strip().replace('\n', ' ')
            if v:
                cells[col] = v
    if cells:
        print(f"  Row {row}: {cells}")

print("\n=== 2.1 PAUT 섹션 (row 520-530) ===")
for row in range(520, 535):
    cells = {}
    for col in range(1, 25):
        c = ws.cell(row=row, column=col)
        if c.value is not None:
            v = str(c.value).strip().replace('\n', ' ')
            if v:
                cells[col] = v
    if cells:
        print(f"  Row {row}: {cells}")
