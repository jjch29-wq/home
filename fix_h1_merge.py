with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'r', encoding='utf-8') as f:
    lines = f.read().split('\n')

for i, line in enumerate(lines):
    if 'ws0 = wb.worksheets[0]' in line and 'def _run_rt_process' in '\n'.join(lines[max(0, i-50):i]):
        insert_code = """
            # [FIX] Merge H1:O4 properly for Cover title so it doesn't get cut off vertically
            try:
                ws0.unmerge_cells('H1:O3')
                ws0.merge_cells('H1:O4')
            except Exception:
                pass"""
        lines.insert(i + 1, insert_code)
        break

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print('SUCCESS')
