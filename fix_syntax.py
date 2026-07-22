import os

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\자재작업일보기성서류Ver.1.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace(r"\'연도\'", "'연도'")
content = content.replace(r"\'월\'", "'월'")
content = content.replace(r"\'현장\'", "'현장'")
content = content.replace(r"\'구분\'", "'구분'")
content = content.replace(r"\'작업자\'", "'작업자'")
content = content.replace(r"\'작업시간\'", "'작업시간'")
content = content.replace(r"\'Site\'", "'Site'")
content = content.replace(r"\'\'", "''")
content = content.replace(r"\'--- 전체 누계 ---\'", "'--- 전체 누계 ---'")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Syntax errors fixed.")
