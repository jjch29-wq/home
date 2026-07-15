import openpyxl

wb = openpyxl.load_workbook(r'C:\Users\-\OneDrive\바탕 화면\지역난방 PAUT.xlsx')
ws0 = wb.worksheets[0]
ws1 = wb.copy_worksheet(ws0)

for r in range(17, 51):
    ws1.row_dimensions[r].height = 25.0
    
wb.save(r'C:\Users\-\PMI\home\scratch_copy_bug.xlsx')

wb2 = openpyxl.load_workbook(r'C:\Users\-\PMI\home\scratch_copy_bug.xlsx')
ws_copied = wb2.worksheets[-1]
print(f"Copied sheet Row 17 height: {ws_copied.row_dimensions[17].height}")
