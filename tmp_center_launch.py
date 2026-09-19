import sys, os
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src')
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src')
import tkinter as tk, re

src_path = r'c:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src\비파괴검사보고서.py'
src = open(src_path, encoding='utf-8').read()
src_no_main = re.sub(r'if __name__\s*==\s*[' + "'\"" + r']__main__[' + "'\"" + r'].*', '', src, flags=re.DOTALL)
g = {'__name__': '__main__', '__file__': src_path}
exec(compile(src_no_main, src_path, 'exec'), g)

root = tk.Tk()
# 화면 중앙으로 강제 이동
root.update_idletasks()
w, h = 1400, 900
sw = root.winfo_screenwidth()
sh = root.winfo_screenheight()
x = (sw - w) // 2
y = (sh - h) // 2
root.geometry(f'{w}x{h}+{x}+{y}')
root.lift()
root.focus_force()

PMIReportApp = g.get('PMIReportApp')
app = PMIReportApp(root)
root.mainloop()
