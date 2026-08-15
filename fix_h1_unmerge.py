with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'r', encoding='utf-8') as f:
    lines = f.read().split('\n')

for i, line in enumerate(lines):
    if '# [FIX] Merge H1:O4' in line:
        insert_code = """
            # [FIX] Merge H1:O4 and adjust row heights to match Data sheet so title isn't cut off
            # Robustly unmerge to avoid openpyxl KeyError
            to_remove = []
            for m_range in list(ws0.merged_cells.ranges):
                if str(m_range) == "H1:O3":
                    to_remove.append(m_range)
            for m_range in to_remove:
                ws0.merged_cells.ranges.remove(m_range)
            
            try:
                ws0.merge_cells('H1:O4')
                ws0.row_dimensions[1].height = 32.25
                ws0.row_dimensions[2].height = 14.1
                ws0.row_dimensions[3].height = 14.1
                ws0.row_dimensions[4].height = 14.1
            except Exception as e:
                self.log(f"[ERROR] Failed to merge H1:O4: {e}")
"""
        # Find the end of the previous fix block
        end_idx = i
        for j in range(i, i+20):
            if 'except: pass' in lines[j] and 'row_dimensions' in lines[j-1]:
                end_idx = j
                break
        
        del lines[i:end_idx+1]
        lines.insert(i, insert_code.strip('\n'))
        break

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print('SUCCESS')
