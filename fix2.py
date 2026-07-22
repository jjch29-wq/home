import os

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\자재작업일보기성서류Ver.1.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace(r"\'날짜\'", "'날짜'")
content = content.replace(r"\'업체명\'", "'업체명'")
content = content.replace(r"\'적용코드\'", "'적용코드'")
content = content.replace(r"\'현장\'", "'현장'")
content = content.replace(r"\'구분\'", "'구분'")
content = content.replace(r"\'검사품명\'", "'검사품명'")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Syntax errors fixed.")
