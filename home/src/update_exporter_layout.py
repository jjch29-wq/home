import os

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Update title span
old_title = "merge_and_set('A2:P2', \"비파괴검사 결과서 및 작업/감독일보\""
new_title = "merge_and_set('A2:R2', \"비파괴검사 결과서 및 작업/감독일보\""
code = code.replace(old_title, new_title)

# 2. Update signatures and anchor
old_sigs = """        set_cell('N4', '현장대리인', font=self.font_small, fill=self.fill_header)
        set_cell('O4', '감독', font=self.font_small, fill=self.fill_header)
        set_cell('P4', '확인', font=self.font_small, fill=self.fill_header)
        
        set_cell('N5', '') # Signature space
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
                
                # Use OneCellAnchor to center the image in N5
                from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
                from openpyxl.drawing.xdr import XDRPositiveSize2D
                from openpyxl.utils.units import pixels_to_EMU
                
                # N is column index 13, row 5 is index 4
                marker = AnchorMarker(col=13, colOff=pixels_to_EMU(12), row=4, rowOff=pixels_to_EMU(3))"""

new_sigs = """        set_cell('O4', '현장대리인', font=self.font_small, fill=self.fill_header)
        set_cell('P4', '감독', font=self.font_small, fill=self.fill_header)
        merge_and_set('Q4:R4', '확인', font=self.font_small, fill=self.fill_header)
        
        set_cell('O5', '') # Signature space
        set_cell('P5', '')
        merge_and_set('Q5:R5', '')
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
                
                # Use OneCellAnchor to center the image in O5
                from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
                from openpyxl.drawing.xdr import XDRPositiveSize2D
                from openpyxl.utils.units import pixels_to_EMU
                
                # O is column index 14, row 5 is index 4
                marker = AnchorMarker(col=14, colOff=pixels_to_EMU(12), row=4, rowOff=pixels_to_EMU(3))"""
code = code.replace(old_sigs, new_sigs)

# 3. Update personnel block
old_pers = """        # Personnel
        merge_and_set('O8:P8', '금일 투입인원(명)', font=self.font_bold, fill=self.fill_header)
        set_cell('O9', '구분(관리/안전)', font=self.font_small, fill=self.fill_header, align=self.align_nowrap)
        set_cell('P9', '검사원', font=self.font_bold, fill=self.fill_header, align=self.align_nowrap)
        
        personnel = data.get('personnel_data', {})
        set_cell('O10', '인원', align=self.align_nowrap)
        set_cell('P10', personnel.get('검사원_인원', ''), align=self.align_nowrap)
        set_cell('O11', '현장대리인', align=self.align_nowrap)
        set_cell('P11', personnel.get('검사원_현장대리인', ''), align=self.align_nowrap)
        set_cell('O12', '누계', align=self.align_nowrap)
        set_cell('P12', personnel.get('검사원_누계', ''), align=self.align_nowrap)
        
        # Merge empty blocks under personnel to match equipment height
        merge_and_set('O13:P14', '', border=self.border_thin)

        # Remarks (특이사항 및 작업계획)
        merge_and_set('L15:P15', '특이사항 및 작업계획', font=self.font_bold, fill=self.fill_header)
        merge_and_set('L16:P23', data.get('remarks', ''), align=Alignment(horizontal='left', vertical='top', wrap_text=True))"""

new_pers = """        # Personnel
        merge_and_set('O8:R8', '금일 투입인원(명)', font=self.font_bold, fill=self.fill_header)
        set_cell('O9', '구분(관리/안전)', font=self.font_small, fill=self.fill_header, align=self.align_nowrap)
        set_cell('P9', '검사원', font=self.font_bold, fill=self.fill_header, align=self.align_nowrap)
        merge_and_set('Q9:R9', '안전담당', font=self.font_bold, fill=self.fill_header, align=self.align_nowrap)
        
        personnel = data.get('personnel_data', {})
        set_cell('O10', '인원', align=self.align_nowrap)
        set_cell('P10', personnel.get('검사원_인원', ''), align=self.align_nowrap)
        merge_and_set('Q10:R10', personnel.get('안전_인원', ''), align=self.align_nowrap)
        
        set_cell('O11', '현장대리인', align=self.align_nowrap)
        set_cell('P11', personnel.get('검사원_현장대리인', ''), align=self.align_nowrap)
        merge_and_set('Q11:R11', personnel.get('안전_현장대리인', ''), align=self.align_nowrap)
        
        set_cell('O12', '누계', align=self.align_nowrap)
        set_cell('P12', personnel.get('검사원_누계', ''), align=self.align_nowrap)
        merge_and_set('Q12:R12', personnel.get('안전_누계', ''), align=self.align_nowrap)
        
        # Merge empty blocks under personnel to match equipment height
        merge_and_set('O13:R14', '', border=self.border_thin)

        # Remarks (특이사항 및 작업계획)
        merge_and_set('L15:R15', '특이사항 및 작업계획', font=self.font_bold, fill=self.fill_header)
        merge_and_set('L16:R23', data.get('remarks', ''), align=Alignment(horizontal='left', vertical='top', wrap_text=True))"""

code = code.replace(old_pers, new_pers)

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated daily_work_log_exporter.py")
