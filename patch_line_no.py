import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "key = (str(r.get('업체', '')), str(r.get('구간', '')), str(r.get('도면번호', '')), str(r.get('관경', '')), str(r.get('Joint No.', '')))" in line:
        lines[i] = line.replace("'도면번호'", "'라인번호'")
        print("Patched 도면번호 to 라인번호 in key")
    
    if "업체, 구간, 도면번호, 관경, joint = key" in line:
        lines[i] = line.replace("도면번호", "라인번호")
        print("Patched 도면번호 to 라인번호 in unpack")
        
    if "ws_paut.Cells(current_row, 5).Value = 도면번호" in line:
        lines[i] = line.replace("도면번호", "라인번호")
        print("Patched 도면번호 to 라인번호 in cell assignment")

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
