import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

# 1. Add doc_num_var UI
insert_ui = '''        ttk.Label(period_frame, text="  \uc6d4:").pack(side='left', padx=5)
        month_var = tk.IntVar(value=now.month)
        ttk.Spinbox(period_frame, from_=1, to=12, textvariable=month_var, width=4).pack(side='left')
        
        ttk.Label(period_frame, text="  \ubb38\uc11c\ubc88\ud638(01~99):").pack(side='left', padx=5)
        doc_num_var = tk.StringVar(value="01")
        ttk.Entry(period_frame, textvariable=doc_num_var, width=5).pack(side='left')'''

text = re.sub(r'ttk\.Label\(period_frame, text=\"\s*\uc6d4:\"\)\.pack.*?ttk\.Spinbox\(period_frame, from_=1, to=12, textvariable=month_var, width=4\)\.pack\(side=\'left\'\)', insert_ui, text, flags=re.DOTALL)

# 2. Extract doc_num in do_export
insert_extract = '''                year = year_var.get()
                month = month_var.get()
                doc_num = doc_num_var.get().strip()'''

text = re.sub(r'year = year_var\.get\(\)\s*month = month_var\.get\(\)', insert_extract, text, flags=re.DOTALL)

# 3. Add Replacement Logic before SaveAs
insert_replace = '''
                # ----------------- PLACEHOLDER REPLACEMENT -----------------
                import datetime
                try:
                    today_str = datetime.datetime.now().strftime("%Y. %m. %d.")
                    ym_str = f"{year}\ub144 {month}\uc6d4"
                    for s_idx in range(1, wb.Sheets.Count + 1):
                        temp_ws = wb.Sheets(s_idx)
                        # xlPart = 2
                        temp_ws.Cells.Replace(What="[[\ubcf4\uace0\uc11c_\uc5f0\uc6d4]]", Replacement=ym_str, LookAt=2)
                        temp_ws.Cells.Replace(What="[[\ubb38\uc11c\ubc88\ud638]]", Replacement=doc_num, LookAt=2)
                        temp_ws.Cells.Replace(What="[[\uc791\uc131\uc77c\uc790]]", Replacement=today_str, LookAt=2)
                except Exception as e:
                    print("Placeholder replacement error:", e)
                # -----------------------------------------------------------

                filepath = filepath.replace("/", "\\\\")
                wb.SaveAs(filepath)'''

text = text.replace('filepath = filepath.replace("/", "\\\\")\n                wb.SaveAs(filepath)', insert_replace)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
print("Successfully modified code!")
