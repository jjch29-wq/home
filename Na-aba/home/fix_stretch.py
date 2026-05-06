import os
import re

file_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\src\main.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace stretch=True with stretch=False inside treeview configuration
# Usually looks like: self.preview_tree.column(col, width=w, anchor='center', stretch=True)
# or self.rt_preview_tree.column(col, width=w, anchor='center', stretch=True)

pattern = re.compile(r"(\.column\(col,\s*width=w,\s*anchor='center',\s*)stretch=True\)")
content, count = pattern.subn(r"\g<1>stretch=False)", content)

print(f"Replaced stretch=True with stretch=False in {count} places.")

if count > 0:
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
