file_path = r"C:\Users\-\PMI\home\src\한국지역난방 중앙지사.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace(r'\"\"\"', '"""')

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print("Fixed quotes.")
