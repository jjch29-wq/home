import os
import win32com.client as win32

filepath = r"C:\Users\-\OneDrive\바탕 화면\JHC RT\박스라벨.xls"
outpath = r"C:\Users\-\OneDrive\바탕 화면\JHC RT\Final_BoxLabel_COM_Test.xlsx"

try:
    print("Starting Excel COM...")
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    print(f"Opening workbook: {filepath}")
    wb = excel.Workbooks.Open(filepath)
    ws = wb.Sheets('2021')
    
    # We want to clear everything below row 6, then copy A1:G6 downwards.
    # To be safe, just copy A1:G6.
    template_range = ws.Range("A1:G6")
    
    # Paste to Row 8 (gap of 1 row)
    template_range.Copy()
    ws.Range("A8").PasteSpecial(Paste=-4104) # xlPasteAll
    
    # Modify data in the copied block (Row 8)
    ws.Cells(8, 1).Value = "BOX NO. < 3 >"
    ws.Cells(9, 2).Value = "SIT-K1-JHC-PIP-RT-TEST"
    
    print(f"Saving to {outpath}")
    wb.SaveAs(outpath, FileFormat=51) # 51 = xlOpenXMLWorkbook (.xlsx)
    wb.Close(SaveChanges=False)
    excel.Quit()
    print("Done!")
    
except Exception as e:
    print(f"Error: {e}")
    try:
        wb.Close(SaveChanges=False)
        excel.Quit()
    except:
        pass
