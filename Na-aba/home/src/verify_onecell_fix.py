import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker, XDRPositiveSize2D
import os
import tempfile
from PIL import Image as PILImage
import glob

# Find template
folder = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data"
pattern = os.path.join(folder, "RT KS*.xlsx")
matches = glob.glob(pattern)
if not matches:
    print("TEMPLATE NOT FOUND")
    exit(1)
template_path = matches[0]

logo_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\SITCO.png"
output_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src\test_onecell_anchor.xlsx"

try:
    print(f"Loading template: {template_path}")
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]
    
    # New logic
    w, h = 100, 50
    x_offset, y_offset = 10, 10
    anchor_cell = "A1"
    
    with PILImage.open(logo_path).convert("RGBA") as img_orig:
        temp_png = os.path.join(tempfile.gettempdir(), "test_logo_v2.png")
        img_orig.resize((w, h), PILImage.Resampling.LANCZOS).save(temp_png, "PNG")
        
        xl_img = XLImage(temp_png)
        xl_img.width = w
        xl_img.height = h
        
        col_idx, row_idx = 0, 0 # A1
        
        # 1 Pixel = 9525 EMU
        final_emu_x = max(0, int(float(x_offset) * 9525))
        final_emu_y = max(0, int(float(y_offset) * 9525))
        emu_w = max(9525, int(float(w) * 9525))
        emu_h = max(9525, int(float(h) * 9525))
        
        marker = AnchorMarker(col=col_idx, colOff=final_emu_x, row=row_idx, rowOff=final_emu_y)
        size = XDRPositiveSize2D(cx=emu_w, cy=emu_h)
        
        xl_img.anchor = OneCellAnchor(_from=marker, ext=size)
        ws.add_image(xl_img)
        
    wb.save(output_path)
    print(f"SUCCESS: Saved to {output_path}")
    
    # Verify by loading again
    wb2 = openpyxl.load_workbook(output_path)
    print(f"Verification: Loaded successfully. Sheet 0 has {len(wb2.worksheets[0]._images)} images.")
    
except Exception as e:
    print(f"FAILURE: {e}")
