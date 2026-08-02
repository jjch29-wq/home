import openpyxl, os
base = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(base, 'data', 'Report_Template_현장사용량.xlsx')
wb = openpyxl.load_workbook(path)
cell_val = wb.active['E29'].value
with open('cell_val.txt', 'w', encoding='utf-8') as f:
    f.write(repr(cell_val) + "\n")
    if cell_val:
        for c in cell_val:
            f.write(f"{c}: {hex(ord(c))}\n")
