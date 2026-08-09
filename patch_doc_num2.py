import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

# 1. UI addition
ui_target = r"ttk\.Spinbox\(period_frame, from_=1, to=12, textvariable=month_var, width=4\)\.pack\(side='left'\)"
ui_replacement = r"""ttk.Spinbox(period_frame, from_=1, to=12, textvariable=month_var, width=4).pack(side='left')
        
        ttk.Label(period_frame, text="  문서번호:").pack(side='left', padx=5)
        doc_num_var = tk.StringVar(value="01")
        ttk.Entry(period_frame, textvariable=doc_num_var, width=5).pack(side='left')"""
text = re.sub(ui_target, ui_replacement, text)

# 2. Variable extraction in do_export
do_export_target = r"month = month_var\.get\(\)\n\s+filepath = file_var\.get\(\)\.strip\(\)"
do_export_replacement = r"""month = month_var.get()
                doc_num = doc_num_var.get().strip() or "01"
                filepath = file_var.get().strip()"""
text = re.sub(do_export_target, do_export_replacement, text)

# 3. MonthlyReportExporter instantiation
exporter_target = r"exporter = MonthlyReportExporter\(history, target_month_str, filepath.*?\)"
exporter_replacement = r"exporter = MonthlyReportExporter(history, target_month_str, filepath, doc_num)"
text = re.sub(exporter_target, exporter_replacement, text)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)

print("Patch applied successfully.")
