import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

new_logic = """                            # 데이터 개수에 맞춰 행 삽입 (TOTAL 행 밀어내기)
                            if len(groups) > 1:
                                for _ in range(len(groups) - 1):
                                    ws_paut.Rows(406).Insert(Shift=-4121, CopyOrigin=0)
                                    
                            current_row = 405
                            for i, (key, data) in enumerate(groups.items()):
"""

for i in range(12170, 12200):
    if 'ws_paut.Range("B405:M1000").ClearContents()' in lines[i]:
        # Comment this out
        lines[i] = lines[i].replace('ws_paut.Range("B405:M1000").ClearContents()', '# ws_paut.Range("B405:M1000").ClearContents()')
        
    if 'for i, (key, data) in enumerate(groups.items()):' in lines[i]:
        # Replace current_row = 405 and the for loop with the new logic
        lines[i-1] = "" # remove current_row = 405
        lines[i] = new_logic
        break

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
