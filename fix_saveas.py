import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\monthly_report_exporter.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

text = text.replace('wb.SaveAs(abs_output, FileFormat=51)', 'abs_output = abs_output.replace("/", "\\\\")\n        wb.SaveAs(abs_output, FileFormat=51)')

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
