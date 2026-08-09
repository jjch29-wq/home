import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

# 1. Add doc_num UI
old_ui = '''        now = datetime.datetime.now()
        ttk.Label(period_frame, text="연도:").pack(side='left', padx=5)
        year_var = tk.IntVar(value=now.year)
        ttk.Spinbox(period_frame, from_=2024, to=2030, textvariable=year_var, width=6).pack(side='left')
        ttk.Label(period_frame, text="  월:").pack(side='left', padx=5)
        month_var = tk.IntVar(value=now.month)
        ttk.Spinbox(period_frame, from_=1, to=12, textvariable=month_var, width=4).pack(side='left')'''

new_ui = old_ui + '''
        
        ttk.Label(period_frame, text="  문서번호:").pack(side='left', padx=5)
        doc_num_var = tk.StringVar(value="01")
        ttk.Entry(period_frame, textvariable=doc_num_var, width=5).pack(side='left')'''

text = text.replace(old_ui, new_ui)

# 2. Add doc_num extraction
old_do = '''        def do_export():
            try:
                year = year_var.get()
                month = month_var.get()
                filepath = file_var.get().strip()'''

new_do = '''        def do_export():
            try:
                year = year_var.get()
                month = month_var.get()
                doc_num = doc_num_var.get().strip() or "01"
                filepath = file_var.get().strip()'''

text = text.replace(old_do, new_do)

# 3. MonthlyReportExporter instantiation
text = re.sub(r'exporter = MonthlyReportExporter\(history, target_month_str, filepath.*?\)', r'exporter = MonthlyReportExporter(history, target_month_str, filepath, doc_num)', text)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
print('Patched doc_num!')
