import os
import openpyxl

desktop = r'C:\Users\-\OneDrive\바탕 화면'
files = [f for f in os.listdir(desktop) if '산출내역서' in f and f.endswith('.xlsx')]
print("Found files:", files)

if files:
    filepath = os.path.join(desktop, files[0])
    print(f"Reading {filepath}...")
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        print("Sheets:", wb.sheetnames)
    except Exception as e:
        print("Error:", e)
