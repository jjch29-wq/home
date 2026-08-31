import win32com.client.dynamic
import os

excel_path = r"C:\Users\-\OneDrive\바탕 화면\템플릿 가산가평월간용역진도보고서.xlsx"

# Make sure path is absolute
excel_path = os.path.abspath(excel_path)

excel = win32com.client.dynamic.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(excel_path)

try:
    # 2. 공정율 (주배관)
    ws_main = wb.Sheets("2. 공정율 (주배관)")
    ws_main.Range("C9").Value = 14205
    ws_main.Range("C10").Value = 4920
    ws_main.Range("C12").Value = 237.66
    ws_main.Range("C13").Value = 81.36
    ws_main.Range("C15").Value = 237.66
    ws_main.Range("C16").Value = 81.35
    
    ws_main.Range("C22").Value = 19125
    ws_main.Range("C23").Value = 0
    ws_main.Range("C24").Value = 0
    
    # 2. 공정율 (관리소)
    ws_station = wb.Sheets("2. 공정율 (관리소)")
    ws_station.Range("C9").Value = 4314
    ws_station.Range("C10").Value = 1097
    ws_station.Range("C12").Value = 0
    ws_station.Range("C13").Value = 0
    ws_station.Range("C15").Value = 15.4
    ws_station.Range("C16").Value = 4.22
    
    ws_station.Range("C22").Value = 1243
    ws_station.Range("C23").Value = 2464
    ws_station.Range("C24").Value = 1704
    
    # 2. 공정율 (전체)
    ws_total = wb.Sheets("2. 공정율 (전체)")
    ws_total.Range("C9").Value = 18519
    ws_total.Range("C10").Value = 6017
    ws_total.Range("C12").Value = 237.66
    ws_total.Range("C13").Value = 81.36
    ws_total.Range("C15").Value = 253.06
    ws_total.Range("C16").Value = 85.57
    
    ws_total.Range("C22").Value = 20368
    ws_total.Range("C23").Value = 2464
    ws_total.Range("C24").Value = 1704
    
    wb.Save()
    print("Successfully updated the Excel file.")
finally:
    wb.Close(SaveChanges=False)
    excel.Quit()
