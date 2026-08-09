import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\daily_work_log_tab.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    text = f.read()

pattern = r'(def export_monthly_report\(self\):)(.*?)(    def load_history\(self\):)'

replacement = r'''\1
        if hasattr(self, 'main_app') and hasattr(self.main_app, 'export_monthly_ndt_report'):
            self.main_app.export_monthly_ndt_report()
        else:
            messagebox.showerror("오류", "메인 어플리케이션과 연결되지 않았습니다.")

\3'''

text = re.sub(pattern, replacement, text, flags=re.DOTALL)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(text)
