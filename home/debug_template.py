import openpyxl
import glob

# Try to find the exact template file
files = glob.glob('C:/Users/**/템플릿*.xlsx', recursive=True)
desktop_template = [f for f in files if 'V70' in f and ('바탕' in f or 'Desktop' in f or '' in f)][0]

wb = openpyxl.load_workbook(desktop_template, data_only=True)
ws = wb.worksheets[0]
print(f"Sheet name: {ws.title}")
print(f"Number of merged cells in [표지]: {len(ws.merged_cells.ranges)}")
for rng in list(ws.merged_cells.ranges)[:5]:
    print(f"  {rng}")
