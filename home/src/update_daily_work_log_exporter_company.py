import os

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Update widths
old_widths = """        col_widths = {
            'A': 6, 'B': 14, 'C': 12, 'D': 21, 'E': 9, 'F': 8, 'G': 15, 'H': 7, 'I': 7, 'J': 7, 'K': 7,
            'L': 7, 'M': 6, 'N': 10, 'O': 10, 'P': 10, 'Q': 7, 'R': 7, 'S': 4, 'T': 4, 'U': 4, 'V': 4, 'W': 4, 'X': 4, 'Y': 4, 'Z': 4
        }"""
new_widths = """        col_widths = {
            'A': 5, 'B': 12, 'C': 8, 'D': 14, 'E': 21, 'F': 9, 'G': 8, 'H': 15, 'I': 7, 'J': 7, 'K': 7, 'L': 7,
            'M': 7, 'N': 6, 'O': 10, 'P': 10, 'Q': 10, 'R': 7, 'S': 7, 'T': 4, 'U': 4, 'V': 4, 'W': 4, 'X': 4, 'Y': 4, 'Z': 4
        }"""
code = code.replace(old_widths, new_widths)

# Update Headers
old_headers = """        headers_ndt = [
            ('A26:A27', '순번'), ('B26:B27', '검사방법'), ('C26:C27', '구간(Sec.No)'), ('D26:D27', '라인번호'),
            ('E26:E27', 'Joint No.'), ('F26:F27', '관경'), ('G26:G27', '용접사'), 
            ('L26:L27', '결과'), ('M26:M27', '규격')
        ]
        for rng, text in headers_ndt:
            merge_and_set(rng, text, font=self.font_small, fill=self.fill_header, align=self.align_nowrap)
            
        merge_and_set('H26:K26', '구간정보(Start/Length)', font=self.font_small, fill=self.fill_header)
        set_cell('H27', '1', font=self.font_small, fill=self.fill_header)
        set_cell('I27', '2', font=self.font_small, fill=self.fill_header)
        set_cell('J27', '3', font=self.font_small, fill=self.fill_header)
        set_cell('K27', '4', font=self.font_small, fill=self.fill_header)
            
        merge_and_set('N26:O26', 'RT매수', font=self.font_small, fill=self.fill_header)
        set_cell('N27', 'OR', font=self.font_small, fill=self.fill_header)
        set_cell('O27', 'RE', font=self.font_small, fill=self.fill_header)
        
        merge_and_set('P26:P27', 'PAUT(m)', font=self.font_small, fill=self.fill_header)
        merge_and_set('Q26:Q27', 'MT(m)', font=self.font_small, fill=self.fill_header)
        merge_and_set('R26:R27', 'PT(m)', font=self.font_small, fill=self.fill_header)"""

new_headers = """        headers_ndt = [
            ('A26:A27', '순번'), ('B26:B27', '업체'), ('C26:C27', '검사방법'), ('D26:D27', '구간(Sec.No)'), ('E26:E27', '라인번호'),
            ('F26:F27', 'Joint No.'), ('G26:G27', '관경'), ('H26:H27', '용접사'), 
            ('M26:M27', '결과'), ('N26:N27', '규격')
        ]
        for rng, text in headers_ndt:
            merge_and_set(rng, text, font=self.font_small, fill=self.fill_header, align=self.align_nowrap)
            
        merge_and_set('I26:L26', '구간정보(Start/Length)', font=self.font_small, fill=self.fill_header)
        set_cell('I27', '1', font=self.font_small, fill=self.fill_header)
        set_cell('J27', '2', font=self.font_small, fill=self.fill_header)
        set_cell('K27', '3', font=self.font_small, fill=self.fill_header)
        set_cell('L27', '4', font=self.font_small, fill=self.fill_header)
            
        merge_and_set('O26:P26', 'RT매수', font=self.font_small, fill=self.fill_header)
        set_cell('O27', 'OR', font=self.font_small, fill=self.fill_header)
        set_cell('P27', 'RE', font=self.font_small, fill=self.fill_header)
        
        merge_and_set('Q26:Q27', 'PAUT(m)', font=self.font_small, fill=self.fill_header)
        merge_and_set('R26:R27', 'MT(m)', font=self.font_small, fill=self.fill_header)
        merge_and_set('S26:S27', 'PT(m)', font=self.font_small, fill=self.fill_header)"""
code = code.replace(old_headers, new_headers)

# Update cell mapping
old_mapping = """            set_cell(f'A{row_idx}', i + 1)
            set_cell(f'B{row_idx}', res.get('검사방법', ''))
            set_cell(f'C{row_idx}', res.get('구간', ''))
            set_cell(f'D{row_idx}', res.get('라인번호', ''))
            set_cell(f'E{row_idx}', res.get('Joint No.', ''))
            set_cell(f'F{row_idx}', res.get('관경', ''))
            set_cell(f'G{row_idx}', res.get('용접사', ''))
            
            # 구간정보 H ~ K
            loc_info = str(res.get('구간정보', '')).split(',')
            for col_idx, loc_val in enumerate(loc_info[:4]):
                col_letter = chr(ord('H') + col_idx)
                set_cell(f'{col_letter}{row_idx}', loc_val.strip())
                
            set_cell(f'L{row_idx}', res.get('결과', ''))
            set_cell(f'M{row_idx}', res.get('규격', ''))
            
            set_cell(f'N{row_idx}', res.get('RT_OR', ''))
            set_cell(f'O{row_idx}', res.get('RT_RE', ''))
            set_cell(f'P{row_idx}', res.get('PAUT', ''))
            set_cell(f'Q{row_idx}', res.get('MT', ''))
            set_cell(f'R{row_idx}', res.get('PT', ''))"""

new_mapping = """            set_cell(f'A{row_idx}', i + 1)
            set_cell(f'B{row_idx}', res.get('업체', ''))
            set_cell(f'C{row_idx}', res.get('검사방법', ''))
            set_cell(f'D{row_idx}', res.get('구간', ''))
            set_cell(f'E{row_idx}', res.get('라인번호', ''))
            set_cell(f'F{row_idx}', res.get('Joint No.', ''))
            set_cell(f'G{row_idx}', res.get('관경', ''))
            set_cell(f'H{row_idx}', res.get('용접사', ''))
            
            # 구간정보 I ~ L
            loc_info = str(res.get('구간정보', '')).split(',')
            for col_idx, loc_val in enumerate(loc_info[:4]):
                col_letter = chr(ord('I') + col_idx)
                set_cell(f'{col_letter}{row_idx}', loc_val.strip())
                
            set_cell(f'M{row_idx}', res.get('결과', ''))
            set_cell(f'N{row_idx}', res.get('규격', ''))
            
            set_cell(f'O{row_idx}', res.get('RT_OR', ''))
            set_cell(f'P{row_idx}', res.get('RT_RE', ''))
            set_cell(f'Q{row_idx}', res.get('PAUT', ''))
            set_cell(f'R{row_idx}', res.get('MT', ''))
            set_cell(f'S{row_idx}', res.get('PT', ''))"""
code = code.replace(old_mapping, new_mapping)

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated daily_work_log_exporter.py successfully")
