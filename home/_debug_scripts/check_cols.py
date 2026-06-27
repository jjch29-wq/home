import os
import re

file_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\src\main.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

for m in re.finditer(r'(self\.[a-z_]*preview_tree\s*=\s*ttk\.Treeview[^)]*columns=\(([^)]+)\))', content):
    print(m.group(0))
    
# Check rt display cols
for m in re.finditer(r'display_cols\s*=\s*\[(.*?)\]', content):
    print(m.group(0)[:200])

