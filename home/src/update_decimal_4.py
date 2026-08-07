import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

code = code.replace('f"{v:.3f}"', 'f"{v:.4f}"')
code = code.replace("f\"{v:.3f}\"", "f\"{v:.4f}\"")

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code_ex = f.read()

code_ex = code_ex.replace("cell.number_format = '0.000'", "cell.number_format = '0.0000'")

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code_ex)

print("Updated formatting to 4 decimal places successfully")
