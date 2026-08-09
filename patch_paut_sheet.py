import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

new_logic = """                # 1.1 비파괴검사 물량표 작성 (PAUT)
                # 시트 이름은 '3. 비파괴검사 현황 (열배관)' 입니다.
                paut_sheet_name = '3. 비파괴검사 현황 (열배관)'
"""

for i, line in enumerate(lines):
    if "paut_sheet_name = None" in line:
        lines[i] = new_logic
        lines[i+1] = "" # for name in sheet_names:
        lines[i+2] = "" # if '위상배열' in name
        lines[i+3] = "" # paut_sheet_name = name
        lines[i+4] = "" # break
        lines[i+5] = "" # if paut_sheet_name:
        lines[i+6] = "" # log DEBUG
        print("Patched PAUT sheet name to '3. 비파괴검사 현황 (열배관)'")
        break

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
