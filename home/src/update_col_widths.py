import os

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_widths = "'A': 6, 'B': 14, 'C': 7, 'D': 25, 'E': 10, 'F': 8, 'G': 15, 'H': 7, 'I': 7, 'J': 7, 'K': 7"
new_widths = "'A': 6, 'B': 14, 'C': 12, 'D': 21, 'E': 9, 'F': 8, 'G': 15, 'H': 7, 'I': 7, 'J': 7, 'K': 7"

code = code.replace(old_widths, new_widths)

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated column widths successfully")
