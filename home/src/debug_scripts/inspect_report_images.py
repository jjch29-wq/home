import openpyxl
import os

report_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\RT_Report_20260510_180003.xlsx"

try:
    print(f"Inspecting report: {report_path}")
    wb = openpyxl.load_workbook(report_path)
    for i, ws in enumerate(wb.worksheets):
        print(f"Sheet {i} ({ws.title}): {len(ws._images)} images")
        for j, img in enumerate(ws._images):
            anchor = img.anchor
            pos = "Unknown"
            if hasattr(anchor, '_from'):
                marker = anchor._from
                pos = f"Col {marker.col}, Row {marker.row}"
            elif hasattr(anchor, 'pos'):
                pos = f"X={anchor.pos.x}, Y={anchor.pos.y}"
            
            print(f"  Image {j}: {pos}")
except Exception as e:
    print(f"ERROR: {e}")
