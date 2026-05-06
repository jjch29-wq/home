import openpyxl
import os

# Use relative path if possible, or list files to find it
for f in os.listdir('.'):
    if 'Template' in f and f.endswith('.xlsx'):
        path = f
        break
else:
    path = "Template_DailyWorkReport.xlsx"

print(f"Checking: {path}")
if not os.path.exists(path):
    print("File not found.")
else:
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = wb.active
    for r in range(1, 100):
        val_a = sheet.cell(row=r, column=1).value
        if val_a and any(str(val_a).startswith(s) for s in ["3.", "4.", "5.", "6."]):
            print(f"Row {r}: {val_a}")
