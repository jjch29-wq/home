import openpyxl
import sys

try:
    file_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx'
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active
    
    print(f"SHEET NAME: {sheet.title}")
    print("CELL MAPPING (OPENPYXL):")
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None:
                print(f"Cell {cell.coordinate}: {cell.value}")
except Exception as e:
    print(f"Error: {e}")
