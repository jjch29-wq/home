import openpyxl
wb = openpyxl.load_workbook(r'f:/내 드라이브/07_Antigravity/PMI_한국지역난방/home/data/Report_Template_현장사용량.xlsx')
cell_val = wb.active['E29'].value
with open('cell_val.txt', 'w', encoding='utf-8') as f:
    f.write(repr(cell_val) + "\n")
    if cell_val:
        for c in cell_val:
            f.write(f"{c}: {hex(ord(c))}\n")
