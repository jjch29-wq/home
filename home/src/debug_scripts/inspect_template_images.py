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
    print(f"Inspecting template: {template_path}")
    wb = openpyxl.load_workbook(template_path)
    for i, ws in enumerate(wb.worksheets):
        print(f"Sheet {i} ({ws.title}): {len(ws._images)} images")
        for j, img in enumerate(ws._images):
            anchor = img.anchor
            pos = "Unknown"
            if hasattr(anchor, '_from'):
                marker = anchor._from
                pos = f"Col {marker.col}, Row {marker.row}"
            elif hasattr(anchor, 'pos'):
                # AbsoluteAnchor
                pos = f"X={anchor.pos.x}, Y={anchor.pos.y}"
            
            print(f"  Image {j}: {pos}")
except Exception as e:
    print(f"ERROR: {e}")
