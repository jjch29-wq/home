import openpyxl
import os
import glob

# Find template
folder = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data"
pattern = os.path.join(folder, "RT KS*.xlsx")
matches = glob.glob(pattern)
if not matches:
    print("TEMPLATE NOT FOUND")
    exit(1)
template_path = matches[0]

try:
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[1] # Sheet 1 (Data Sheet)
    print(f"Inspecting Sheet 1 (Data) rows 1-40:")
    for r in range(1, 41):
        row_vals = [str(ws.cell(row=r, column=c).value) for c in range(1, 14)]
        print(f"Row {r}: {' | '.join(row_vals)}")
except Exception as e:
    print(f"ERROR: {e}")
