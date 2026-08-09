import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i in range(11740, 11760):
    if 'month_var = tk.IntVar(value=now.month)' in lines[i]:
        lines.insert(i+2, "        ttk.Label(period_frame, text=\"  문서번호:\").pack(side='left', padx=(15, 5))\n")
        lines.insert(i+3, "        doc_num_var = tk.StringVar(value=\"\")\n")
        lines.insert(i+4, "        ttk.Entry(period_frame, textvariable=doc_num_var, width=15).pack(side='left')\n")
        break

for i in range(11790, 11815):
    if 'filepath = file_var.get().strip()' in lines[i]:
        lines.insert(i, "                doc_num = doc_num_var.get().strip()\n")
        break

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
