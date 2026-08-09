import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i in range(12100, 12130):
    if 'if main_agg:' in lines[i]:
        lines.insert(i, "                sheet_names = [sheet.Name for sheet in wb.Sheets]\n")
        break

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
