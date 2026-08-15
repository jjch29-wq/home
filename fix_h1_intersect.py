with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'r', encoding='utf-8') as f:
    lines = f.read().split('\n')

for i, line in enumerate(lines):
    if '# [FIX] Merge H1:O4' in line:
        insert_code = """
            # [FIX] Merge H1:O4 and adjust row heights to match Data sheet so title isn't cut off
            # Robustly unmerge to avoid openpyxl KeyError
            to_remove = []
            for m_range in list(ws0.merged_cells.ranges):
                # Intersects with H1:O4 (col 8-15, row 1-4)?
                # bounds = (min_col, min_row, max_col, max_row)
                min_c, min_r, max_c, max_r = m_range.bounds
                if not (max_c < 8 or min_c > 15 or max_r < 1 or min_r > 4):
                    to_remove.append(m_range)
            
            for m_range in to_remove:
                try:
                    ws0.merged_cells.ranges.remove(m_range)
                except Exception as e:
                    self.log(f"[ERROR] unmerge fail: {e}")
            
            try:
                ws0.merge_cells(start_row=1, start_column=8, end_row=4, end_column=15)
                ws0.row_dimensions[1].height = 32.25
                ws0.row_dimensions[2].height = 14.1
                ws0.row_dimensions[3].height = 14.1
                ws0.row_dimensions[4].height = 14.1
            except Exception as e:
                self.log(f"[ERROR] Failed to merge H1:O4: {e}")
"""
        # Find the end of the previous fix block
        end_idx = i
        for j in range(i, i+25):
            if 'self.add_logos_to_sheet' in lines[j]:
                end_idx = j - 1
                break
        
        del lines[i:end_idx+1]
        lines.insert(i, insert_code.strip('\n'))
        break

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print('SUCCESS')
