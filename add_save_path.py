import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

for i in range(11810, 11825):
    if 'if not filepath:' in lines[i]:
        # insert save_path prompt after the if not filepath block
        insert_code = """
                from tkinter import filedialog
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=f"월간진도보고서_{year}년_{month:02d}월.xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    title="월간진도보고서 통합 저장"
                )
                if not save_path:
                    return
"""
        lines.insert(i+3, insert_code)
        break

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
