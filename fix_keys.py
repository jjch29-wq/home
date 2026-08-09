import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "if str(r.get('검사', '')).strip().upper() == 'PAUT':" in line:
        lines[i] = line.replace("'검사'", "'검사방법'")
        print(f"Fixed line {i+1}: '검사' -> '검사방법'")
    
    if "shift = str(r.get('주야간', '주간')).strip()" in line:
        lines[i] = line.replace("'주야간'", "'규격'")
        print(f"Fixed line {i+1}: '주야간' -> '규격'")

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
