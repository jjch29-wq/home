import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "paut_sheet_name = '3. 비파괴검사 현황 (열배관)'" in line:
        indent = line[:line.find("paut_sheet_name")]
        lines.insert(i, indent + "messagebox.showinfo('디버그', 'PAUT 처리 시작!')\n")
        break

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
