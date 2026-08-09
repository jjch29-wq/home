import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if 'year = year_var.get()' in line and line.startswith('                                year = year_var.get()'):
        lines[i] = '                year = year_var.get()\n'
        print(f'Fixed indentation at line {i+1}')
        break

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
