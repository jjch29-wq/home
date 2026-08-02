import openpyxl

wb = openpyxl.load_workbook(r'C:\Users\jjch2\Desktop\중앙지사 종합 리스트.xlsx')

for sheet_name in ['PT', 'MT']:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
        
rt_sheet = wb['RT']

# Copy RT sheet twice
pt_sheet = wb.copy_worksheet(rt_sheet)
pt_sheet.title = 'PT'

mt_sheet = wb.copy_worksheet(rt_sheet)
mt_sheet.title = 'MT'

wb.save(r'C:\Users\jjch2\Desktop\중앙지사 종합 리스트.xlsx')
print("RT format successfully copied to PT and MT.")
