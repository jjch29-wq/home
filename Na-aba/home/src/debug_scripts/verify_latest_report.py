import openpyxl
import os

report_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\RT_Report_20260510_205846.xlsx"

try:
    print(f"Inspecting report: {report_path}")
    wb = openpyxl.load_workbook(report_path)
    ws0 = wb.worksheets[0]
    print(f"Sheet 0 ({ws0.title}): {len(ws0._images)} images")
    for i, img in enumerate(ws0._images):
        anchor = img.anchor
        pos = "Unknown"
        if hasattr(anchor, '_from'):
            marker = anchor._from
            pos = f"Col {marker.col}, Row {marker.row}"
        print(f"  Image {i}: {pos}")
    
    print("\nChecking rows 32-42 content (Gapji):")
    for r in range(32, 43):
        row_vals = [str(ws0.cell(row=r, column=c).value) for c in range(1, 10)]
        print(f"Row {r}: {' | '.join(row_vals)}")
        
except Exception as e:
    print(f"ERROR: {e}")
