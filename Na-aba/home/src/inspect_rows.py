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
    print(f"Inspecting template content: {template_path}")
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]
    for r in range(30, 50):
        row_vals = [str(ws.cell(row=r, column=c).value) for c in range(1, 10)]
        print(f"Row {r}: {' | '.join(row_vals)}")
except Exception as e:
    print(f"ERROR: {e}")
