import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

with codecs.open('out_lines.txt', 'w', 'utf-8') as out:
    for i, line in enumerate(lines):
        if 'paut_sheet_name =' in line:
            for j in range(i-5, i+60):
                out.write(f'{j+1}: {lines[j]}')
            break
