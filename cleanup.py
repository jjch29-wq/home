with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'r', encoding='utf-8') as f:
    lines = f.read().split('\n')

new_lines = []
for line in lines:
    if 'debug_merge.txt' in line or 'f_dbg.write' in line or 'for m in ws0.merged_cells.ranges:' in line or "if 'H' in str(m) or 'O' in str(m): f_dbg.write(str(m)" in line or line.strip() == "BEFORE:" or line.strip() == "AFTER:" or line.strip() == "BEFORE SAVE:":
        continue
    new_lines.append(line)

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(new_lines))
print('Cleaned up debug code')
