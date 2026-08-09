import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\monthly_report_exporter.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if 'def __init__(self, history, target_month, template_path):' in line:
        lines[i] = '    def __init__(self, history, target_month, template_path, doc_num="01"):\n'
    elif 'self.template_path = template_path' in line:
        lines[i] = '        self.template_path = template_path\n        self.doc_num = doc_num\n'
    elif 'self.replacements[\'[[보고서_월]]\']' in line:
        lines.insert(i+1, '        self.replacements[\'[[문서번호]]\'] = self.doc_num\n')
        lines.insert(i+2, '        self.replacements[\'[[작성일자]]\'] = datetime.now().strftime("%Y. %m. %d.")\n')
        break

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
