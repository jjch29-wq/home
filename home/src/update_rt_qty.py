import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_rt_qty = """                ('RT', '150A~100A'): '95', ('RT', '150A~100A-야간'): '14',
                ('RT', '80A이하'): '34', ('RT', '80A이하-야간'): '16', ('RT', '소계'): '159',"""

new_rt_qty = """                ('RT', '150A~100A'): '293', ('RT', '150A~100A-야간'): '43',
                ('RT', '80A이하'): '105', ('RT', '80A이하-야간'): '49', ('RT', '소계'): '490',"""

code = code.replace(old_rt_qty, new_rt_qty)

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated RT default quantities to film counts successfully")
