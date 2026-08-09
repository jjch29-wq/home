import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

insert_idx = -1
for i in range(12800, 12830):
    if 'filepath = filepath.replace("/", "\\\\")' in lines[i]:
        insert_idx = i
        break

if insert_idx != -1:
    code_to_insert = """                import datetime
                report_ym = f"{year}년 {month:02d}월"
                today_str = datetime.datetime.now().strftime("%Y. %m. %d.")
                for s_idx in range(1, wb.Sheets.Count + 1):
                    ws_temp = wb.Sheets(s_idx)
                    for r in range(1, 150):
                        for c in range(1, 20):
                            val = str(ws_temp.Cells(r, c).Value or '')
                            if '[[문서번호]]' in val or '[[보고서_연월]]' in val or '[[작성일자]]' in val:
                                new_val = val.replace('[[문서번호]]', doc_num)
                                new_val = new_val.replace('[[보고서_연월]]', report_ym)
                                new_val = new_val.replace('[[작성일자]]', today_str)
                                ws_temp.Cells(r, c).Value = new_val
"""
    lines.insert(insert_idx, code_to_insert)
    
    with codecs.open(file_path, 'w', 'utf-8') as f:
        f.writelines(lines)
    print(f'Successfully injected tag replacement logic at line {insert_idx}')
else:
    print('Could not find insertion point.')
