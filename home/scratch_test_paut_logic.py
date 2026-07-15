import openpyxl

wb = openpyxl.load_workbook(r'C:\Users\-\OneDrive\바탕 화면\지역난방 PAUT.xlsx')
ws = wb.worksheets[0]

# mimic run_pmi_process
for r in range(17, 50):
    ws.row_dimensions[r].height = 20.55
    
# mimic apply_custom_dimensions
for r in range(17, 51):
    ws.row_dimensions[r].height = 25.0
    
wb.save(r'C:\Users\-\OneDrive\바탕 화면\TEST_PAUT2.xlsx')

# reload and check
wb2 = openpyxl.load_workbook(r'C:\Users\-\OneDrive\바탕 화면\TEST_PAUT2.xlsx')
ws2 = wb2.worksheets[0]
print(f"Row 17 height: {ws2.row_dimensions[17].height}")
print(f"Row 50 height: {ws2.row_dimensions[50].height}")
