import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

new_logic = """                # 1.1 비파괴검사 물량표 작성 (PAUT)
                # 시트 이름이 약간 다를 수 있으므로 '위상배열'이 포함된 시트를 찾습니다.
                paut_sheet_name = None
                for name in sheet_names:
                    if '위상배열' in name or 'PAUT' in name.upper():
                        paut_sheet_name = name
                        break
                
                if paut_sheet_name:
                    log(f"DEBUG: Found PAUT sheet: {paut_sheet_name}")
"""

for i, line in enumerate(lines):
    if "paut_sheet_name = '1.1.2.1 " in line:
        # replace the hardcoded assignment and the if check
        lines[i-1] = "" # # 1.1 비파괴...
        lines[i] = new_logic
        lines[i+1] = "" # if paut_sheet_name in sheet_names:
        print("Patched PAUT sheet name logic!")
        break

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
