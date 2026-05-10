import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor, XDRPoint2D, XDRPositiveSize2D
import os
import tempfile
from PIL import Image as PILImage

template_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\RT KS.xlsx"
logo_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\SITCO.png"
output_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src\test_absolute_anchor.xlsx"

try:
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]
    
    # Simulate place_image_freely logic
    w, h = 100, 50
    x_offset, y_offset = 10, 10
    
    with PILImage.open(logo_path).convert("RGBA") as img_orig:
        temp_png = os.path.join(tempfile.gettempdir(), "test_logo.png")
        img_orig.resize((w, h), PILImage.Resampling.LANCZOS).save(temp_png, "PNG")
        
        xl_img = XLImage(temp_png)
        xl_img.width = w
        xl_img.height = h
        
        # Using the same logic as the Archived script
        final_emu_x = max(0, int(float(x_offset) * 12700))
        final_emu_y = max(0, int(float(y_offset) * 12700))
        emu_w = max(12700, int(float(w) * 9525))
        emu_h = max(12700, int(float(h) * 9525))
        
        pos = XDRPoint2D(x=final_emu_x, y=final_emu_y)
        size = XDRPositiveSize2D(cx=emu_w, cy=emu_h)
        
        xl_img.anchor = AbsoluteAnchor(pos=pos, ext=size)
        ws.add_image(xl_img)
        
    wb.save(output_path)
    print(f"SUCCESS: Saved to {output_path}")
except Exception as e:
    print(f"FAILURE: {e}")
