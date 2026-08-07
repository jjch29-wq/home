import os

with open('ndt_summary_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Update _write_summary_sheet
old_summary_headers = """        headers = ['구간(Sec.No)', '라인번호', '관경', '검사방법', '검사 횟수(Joint 카운트)', '물량 누계']
        col_widths = {'A': 15, 'B': 25, 'C': 12, 'D': 12, 'E': 22, 'F': 25}
        
        for col_idx, width in enumerate(col_widths.values(), start=1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width
            
        # Write Title
        ws.merge_cells('A1:F1')"""

new_summary_headers = """        headers = ['업체', '구간(Sec.No)', '라인번호', '관경', '검사방법', '검사 횟수(Joint 카운트)', '물량 누계']
        col_widths = {'A': 15, 'B': 15, 'C': 25, 'D': 12, 'E': 12, 'F': 22, 'G': 25}
        
        for col_idx, width in enumerate(col_widths.values(), start=1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width
            
        # Write Title
        ws.merge_cells('A1:G1')"""
code = code.replace(old_summary_headers, new_summary_headers)

old_summary_dict = """        summary_dict = {}
        for method, rows in grouped_data.items():
            for r in rows:
                sec = str(r.get('구간', '')).strip()
                line = str(r.get('라인번호', '')).strip()
                size = str(r.get('관경', '')).strip()
                
                key = (sec, line, size, method)"""

new_summary_dict = """        summary_dict = {}
        for method, rows in grouped_data.items():
            for r in rows:
                company = str(r.get('업체', '')).strip()
                sec = str(r.get('구간', '')).strip()
                line = str(r.get('라인번호', '')).strip()
                size = str(r.get('관경', '')).strip()
                
                key = (company, sec, line, size, method)"""
code = code.replace(old_summary_dict, new_summary_dict)

old_summary_write = """        for key in sorted_keys:
            sec, line, size, method = key
            stats = summary_dict[key]
            
            # 1~4: Group keys
            for i, val in enumerate(key, start=1):
                cell = ws.cell(row=row_idx, column=i, value=val)
                cell.font = self.font_normal
                cell.alignment = self.align_center
                cell.border = self.border_thin
                
            # 5: Count
            cell = ws.cell(row=row_idx, column=5, value=f"{stats['count']} 개")
            cell.font = self.font_normal
            cell.alignment = self.align_center
            cell.border = self.border_thin
            
            # 6: Qty
            cell = ws.cell(row=row_idx, column=6)
            if method == 'RT':"""

new_summary_write = """        for key in sorted_keys:
            company, sec, line, size, method = key
            stats = summary_dict[key]
            
            # 1~5: Group keys
            for i, val in enumerate(key, start=1):
                cell = ws.cell(row=row_idx, column=i, value=val)
                cell.font = self.font_normal
                cell.alignment = self.align_center
                cell.border = self.border_thin
                
            # 6: Count
            cell = ws.cell(row=row_idx, column=6, value=f"{stats['count']} 개")
            cell.font = self.font_normal
            cell.alignment = self.align_center
            cell.border = self.border_thin
            
            # 7: Qty
            cell = ws.cell(row=row_idx, column=7)
            if method == 'RT':"""
code = code.replace(old_summary_write, new_summary_write)


# 2. Update _write_sheet
old_sheet_headers = """    def _write_sheet(self, ws, method, rows):
        # Common headers
        headers = ['일자', '순번', '구간(Sec.No)', '라인번호', 'Joint No.', '관경', '용접사', '구간정보', '결과', '규격']
        
        # Quantity headers
        if method == 'RT':
            qty_headers = ['RT_OR(매)', 'RT_RE(매)']
        else:
            qty_headers = ['물량(m)']
            
        all_headers = headers + qty_headers
        
        # Set column widths
        col_widths = {
            'A': 12, 'B': 6, 'C': 12, 'D': 25, 'E': 10, 'F': 8, 'G': 15, 'H': 15, 'I': 10, 'J': 10, 'K': 12, 'L': 12
        }"""

new_sheet_headers = """    def _write_sheet(self, ws, method, rows):
        # Common headers
        headers = ['일자', '순번', '업체', '구간(Sec.No)', '라인번호', 'Joint No.', '관경', '용접사', '구간정보', '결과', '규격']
        
        # Quantity headers
        if method == 'RT':
            qty_headers = ['RT_OR(매)', 'RT_RE(매)']
        else:
            qty_headers = ['물량(m)']
            
        all_headers = headers + qty_headers
        
        # Set column widths
        col_widths = {
            'A': 12, 'B': 6, 'C': 15, 'D': 12, 'E': 25, 'F': 10, 'G': 8, 'H': 15, 'I': 15, 'J': 10, 'K': 10, 'L': 12, 'M': 12
        }"""
code = code.replace(old_sheet_headers, new_sheet_headers)

old_col_map = """            # Map standard columns
            col_map = {
                '일자': '일자',
                '순번': '순번',
                '구간(Sec.No)': '구간',
                '라인번호': '라인번호',
                'Joint No.': 'Joint No.',
                '관경': '관경',
                '용접사': '용접사',
                '구간정보': '구간정보',
                '결과': '결과',
                '규격': '규격'
            }"""

new_col_map = """            # Map standard columns
            col_map = {
                '일자': '일자',
                '순번': '순번',
                '업체': '업체',
                '구간(Sec.No)': '구간',
                '라인번호': '라인번호',
                'Joint No.': 'Joint No.',
                '관경': '관경',
                '용접사': '용접사',
                '구간정보': '구간정보',
                '결과': '결과',
                '규격': '규격'
            }"""
code = code.replace(old_col_map, new_col_map)

with open('ndt_summary_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated ndt_summary_exporter.py successfully")
