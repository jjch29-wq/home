import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

# 1. Inject doc_num into UI
target_ui = """ttk.Spinbox(period_frame, from_=1, to=12, textvariable=month_var, width=4).pack(side='left')"""
replacement_ui = """ttk.Spinbox(period_frame, from_=1, to=12, textvariable=month_var, width=4).pack(side='left')

        ttk.Label(period_frame, text="  문서번호:").pack(side='left', padx=(15, 5))
        doc_num_var = tk.StringVar(value="")
        ttk.Entry(period_frame, textvariable=doc_num_var, width=15).pack(side='left')"""
text = text.replace(target_ui, replacement_ui)

# 2. Inject doc_num extraction in do_export
target_export = """            def do_export():
                try:
                    year = year_var.get()
                    month = month_var.get()
                    filepath = file_var.get().strip()"""
replacement_export = """            def do_export():
                try:
                    year = year_var.get()
                    month = month_var.get()
                    doc_num = doc_num_var.get().strip()
                    filepath = file_var.get().strip()"""
text = text.replace(target_export, replacement_export)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
