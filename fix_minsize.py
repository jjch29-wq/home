import os

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\자재작업일보기성서류Ver.1.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace(
    "self.master_form_panel.columnconfigure(1, weight=1)",
    "self.master_form_panel.columnconfigure(1, weight=1, minsize=550)"
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Minsize added to column 1.")
