import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

old_do_export = '''        def do_export():
            try:
                year = year_var.get()
                month = month_var.get()
                filepath = file_var.get().strip()'''

new_do_export = '''        def do_export():
            try:
                year = year_var.get()
                month = month_var.get()
                doc_num = doc_num_var.get().strip() or "01"
                filepath = file_var.get().strip()'''

if old_do_export in text:
    text = text.replace(old_do_export, new_do_export)
    print('Replaced do_export')
else:
    print('old_do_export NOT FOUND!')

# Replace exporter instantiation using regex
text = re.sub(r'exporter = MonthlyReportExporter\(history, target_month_str, filepath.*?\)', r'exporter = MonthlyReportExporter(history, target_month_str, filepath, doc_num)', text)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
print("Patch applied successfully.")
