import win32com.client
import os

try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(r'c:\Users\jjch2\Desktop\PMI\home\data\기성서류_기본양식.xlsx')
    sheet_names = [sheet.Name for sheet in wb.Sheets]
    with open('out_sheets_final.txt', 'w', encoding='utf-8') as f:
        f.write(str(sheet_names))
    wb.Close(False)
    excel.Quit()
except Exception as e:
    with open('out_sheets_final.txt', 'w', encoding='utf-8') as f:
        f.write(f"Error: {e}")
