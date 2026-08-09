import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "raw_data = log_data.get('ndt_results', [])" in line:
        lines[i] = line.replace("raw_data", "ndt_results")
        print(f"Fixed line {i+1}: raw_data -> ndt_results")

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
