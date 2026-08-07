import os

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_code = """        set_cell('N5', '') # Signature space
        set_cell('O5', '')
        set_cell('P5', '')
        ws.row_dimensions[5].height = 30 # Make signature box taller"""

new_code = """        set_cell('N5', '') # Signature space
        set_cell('O5', '')
        set_cell('P5', '')
        ws.row_dimensions[5].height = 30 # Make signature box taller
        
        # Add signature image
        try:
            from openpyxl.drawing.image import Image
            sign_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'signs', '주진철.png')
            if os.path.exists(sign_path):
                img = Image(sign_path)
                # Resize image to fit the cell (approx width 70, height 38)
                img.width = 50
                img.height = 35
                
                # N5 cell is at row 5, col 14. openpyxl add_image anchors to top-left of cell.
                # A bit of offset for centering could be done, but default anchor is usually fine.
                ws.add_image(img, 'N5')
        except Exception as e:
            print(f"Error adding signature: {e}")"""

code = code.replace(old_code, new_code)

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Added signature insertion logic successfully")
