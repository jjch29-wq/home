with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'r', encoding='utf-8') as f:
    lines = f.read().split('\n')

for i, line in enumerate(lines):
    if '# [FIX] Merge H1:O4 properly for Cover title' in line:
        insert_code = """
            # [FIX] Merge H1:O4 and adjust row heights to match Data sheet so title isn't cut off
            try:
                ws0.unmerge_cells('H1:O3')
            except: pass
            try:
                ws0.merge_cells('H1:O4')
                ws0.row_dimensions[1].height = 32.25
                ws0.row_dimensions[2].height = 14.1
                ws0.row_dimensions[3].height = 14.1
                ws0.row_dimensions[4].height = 14.1
            except: pass"""
        
        # Remove old fix block
        del lines[i:i+6] 
        # Insert new fix block
        lines.insert(i, insert_code.strip('\n'))
        break

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print('SUCCESS')
