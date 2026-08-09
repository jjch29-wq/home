import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\daily_work_log_tab.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

import re

# Find the export_monthly_report function
pattern = r'(def export_monthly_report\(self\):.*?)(target_month = simpledialog\.askstring.*?if not target_month:\s*return)(.*?)(template_path = filedialog\.askopenfilename)'

custom_dialog_code = """        # Custom dialog for Year, Month, Doc Num
        top = tk.Toplevel(self.root)
        top.title("월간진도보고서 생성")
        top.geometry("350x200")
        top.transient(self.root)
        top.grab_set()
        
        result_vars = {}
        
        ttk.Label(top, text="연도:").place(x=30, y=30)
        year_var = tk.IntVar(value=datetime.today().year)
        ttk.Spinbox(top, from_=2024, to=2030, textvariable=year_var, width=6).place(x=80, y=30)
        
        ttk.Label(top, text="월:").place(x=170, y=30)
        month_var = tk.IntVar(value=datetime.today().month)
        ttk.Spinbox(top, from_=1, to=12, textvariable=month_var, width=4).place(x=200, y=30)
        
        ttk.Label(top, text="문서번호(예: 01):").place(x=30, y=80)
        doc_num_var = tk.StringVar(value="01")
        ttk.Entry(top, textvariable=doc_num_var, width=10).place(x=150, y=80)
        
        def on_confirm():
            result_vars['year'] = year_var.get()
            result_vars['month'] = month_var.get()
            result_vars['doc_num'] = doc_num_var.get().strip()
            top.destroy()
            
        def on_cancel():
            top.destroy()
            
        ttk.Button(top, text="확인", command=on_confirm).place(x=80, y=140)
        ttk.Button(top, text="취소", command=on_cancel).place(x=180, y=140)
        
        self.root.wait_window(top)
        
        if 'year' not in result_vars:
            return
            
        target_month = f"{result_vars['year']}-{result_vars['month']:02d}"
        doc_num = result_vars['doc_num']
"""

text = re.sub(pattern, r'\1' + custom_dialog_code + r'\3\4', text, flags=re.DOTALL)

# Also pass doc_num to exporter
text = text.replace(
    'exporter = MonthlyReportExporter(history, target_month, template_path)',
    'exporter = MonthlyReportExporter(history, target_month, template_path, doc_num)'
)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
