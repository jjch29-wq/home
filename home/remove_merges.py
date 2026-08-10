with open('src/services/monthly_report_manager.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if "cover_merges = []" in line and "시트의 병합 셀이" in lines[i-1]:
        skip = True
        
    if "ws = wb.active" in line:
        skip = False
        
    if "if '표지' in wb.sheetnames and cover_merges:" in line:
        skip = True
        
    if "wb.save(output_path)" in line:
        skip = False
        
    if not skip:
        new_lines.append(line)

with open('src/services/monthly_report_manager.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)
