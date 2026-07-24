import os
import win32com.client as win32

files_to_convert = [
    (r"C:\Users\-\OneDrive\바탕 화면\05.인원 및 장비투입계획서 동탄.xls", r"C:\Users\-\PMI\home\src\templates\인원_장비투입_동탄양식.xlsx"),
    (r"C:\Users\-\OneDrive\바탕 화면\04.인원 및 장비투입계획서.xls", r"C:\Users\-\PMI\home\src\templates\인원_장비투입_기본양식.xlsx")
]

try:
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    for src, dst in files_to_convert:
        if os.path.exists(dst):
            os.remove(dst)
        print(f"Converting {src}...")
        wb = excel.Workbooks.Open(src)
        # 51 = xlOpenXMLWorkbook (.xlsx)
        wb.SaveAs(dst, FileFormat=51)
        wb.Close()
        print(f"Saved to {dst}")
        
except Exception as e:
    print("Error:", e)
finally:
    excel.Quit()
