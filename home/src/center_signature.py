import os

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_code = """                # N5 cell is at row 5, col 14. openpyxl add_image anchors to top-left of cell.
                # A bit of offset for centering could be done, but default anchor is usually fine.
                ws.add_image(img, 'N5')"""

new_code = """                # Use OneCellAnchor to center the image in N5
                from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
                from openpyxl.drawing.xdr import XDRPositiveSize2D
                from openpyxl.utils.units import pixels_to_EMU
                
                # N is column index 13, row 5 is index 4
                marker = AnchorMarker(col=13, colOff=pixels_to_EMU(12), row=4, rowOff=pixels_to_EMU(3))
                size = XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height))
                img.anchor = OneCellAnchor(_from=marker, ext=size)
                
                ws.add_image(img)"""

code = code.replace(old_code, new_code)

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated signature anchor successfully")
