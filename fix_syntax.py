import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i in range(11980, 11995):
    if 'ws.Cells(row, col).Value =' in lines[i] and lines[i].endswith(')\r\n'):
        lines[i] = lines[i][:-3] + '\r\n'
    elif 'ws.Cells(row, col).Value =' in lines[i] and lines[i].endswith(')\n'):
        lines[i] = lines[i][:-2] + '\n'

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
