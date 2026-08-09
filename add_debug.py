import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "paut_sheet_name = '1.1.2.1 위상배열초음파탐상검사'" in line:
        # inject debug log
        indent = line[:line.find('paut_sheet_name')]
        lines.insert(i+1, indent + "log(f\"DEBUG: paut_sheet_name: '{paut_sheet_name}', in sheet_names? {paut_sheet_name in sheet_names}\")\n")
        lines.insert(i+2, indent + "log(f\"DEBUG: sheet_names: {sheet_names}\")\n")
        break

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
