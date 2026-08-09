import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

text = re.sub(
    r'(year = year_var\.get\(\)\r?\n\s+month = month_var\.get\(\)\r?\n\s+filepath = file_var\.get\(\)\.strip\(\))',
    r'year = year_var.get()\n                month = month_var.get()\n                doc_num = doc_num_var.get().strip() or "01"\n                filepath = file_var.get().strip()',
    text
)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
print("Patch applied successfully.")
