import openpyxl

template_path = r'C:\Users\-\OneDrive\바탕 화면\템플릿 가산가평월간용역진도보고서.xlsx'
wb = openpyxl.load_workbook(template_path, data_only=True)

for sheet_name in wb.sheetnames:
    if '5' in sheet_name or '용접' in sheet_name or '불량' in sheet_name:
        print(f"Found sheet: '{sheet_name}'")
