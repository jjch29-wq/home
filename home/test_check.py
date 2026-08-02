import openpyxl
wb = openpyxl.load_workbook(r'f:/내 드라이브/07_Antigravity/PMI_한국지역난방/home/data/Report_Template_현장사용량.xlsx')
print(repr(wb.active['E29'].value))
print(repr(wb.active['E30'].value))
print(repr(wb.active['H29'].value))
