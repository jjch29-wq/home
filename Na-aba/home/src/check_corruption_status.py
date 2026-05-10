import openpyxl
import os

path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\RT_Report_20260510_173619.xlsx"

try:
    print(f"Attempting to load: {path}")
    wb = openpyxl.load_workbook(path)
    print("SUCCESS: openpyxl loaded the workbook.")
    print(f"Sheets: {wb.sheetnames}")
    for sheet in wb.worksheets:
        print(f"Sheet '{sheet.title}' has {len(sheet._images)} images.")
except Exception as e:
    print(f"FAILURE: openpyxl failed to load the workbook.")
    print(f"Error: {e}")
