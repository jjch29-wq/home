import os

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_header = "('C26:C27', '구간(Section No.)')"
new_header = "('C26:C27', '구간(Sec.No)')"

code = code.replace(old_header, new_header)

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated header to 구간(Sec.No) successfully")
