import openpyxl
import glob

files = glob.glob(r'C:\Users\-\PMI\home\assets\*.xlsx')
if not files:
    print("No xlsx files found.")
else:
    template_path = files[0]
    print("Using:", template_path)
    wb = openpyxl.load_workbook(template_path)
    ws = wb['표지']
    print(f"Merged cells in 표지: {len(ws.merged_cells.ranges)}")
    for rng in list(ws.merged_cells.ranges)[:5]:
        print(f"  {rng}")
