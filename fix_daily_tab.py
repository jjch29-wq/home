import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\daily_work_log_tab.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

text = text.replace('top = tk.Toplevel(self.root)', 'top = tk.Toplevel(self.winfo_toplevel())')
text = text.replace('top.transient(self.root)', 'top.transient(self.winfo_toplevel())')
text = text.replace('self.root.wait_window(top)', 'self.wait_window(top)')

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
