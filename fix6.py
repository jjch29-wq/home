import os

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\자재작업일보기성서류Ver.1.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace(r"\'tv_recent_col_widths\'", "'tv_recent_col_widths'")
content = content.replace(r"\'tv_recent\'", "'tv_recent'")
content = content.replace(r"\'columns\'", "'columns'")
content = content.replace(r"\'width\'", "'width'")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Syntax errors fixed.")
