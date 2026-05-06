
import openpyxl
import os

template_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx'
wb = openpyxl.load_workbook(template_path)
sheet = wb.active

print(f"Images count: {len(sheet._images)}")
for img in sheet._images:
    anchor = img.anchor
    # Anchor can be a OneCellAnchor or TwoCellAnchor
    # Usually it's a TwoCellAnchor in openpyxl
    try:
        from_row = anchor._from.row
        print(f"Image anchor row: {from_row + 1}")
    except:
        print(f"Image anchor: {anchor}")
