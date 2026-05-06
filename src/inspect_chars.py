from openpyxl import load_workbook
import os

path = os.path.abspath("../resources/Template_DailyWorkReport.xlsx")
wb = load_workbook(path)
sheet = wb.active

# Section 3 checklist items usually start around Row 29
cell_val = sheet['E29'].value
if cell_val:
    print(f"E29 value: {repr(cell_val)}")
    print("Hex codes:")
    for char in cell_val:
        print(f"  {char}: {hex(ord(char))}")

wb.close()
