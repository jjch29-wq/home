import os
import re

# 1. Update kogas_daily_work_log_tab.py
tab_path = r"c:\Users\jjch2\Desktop\PMI\home\src\kogas_daily_work_log_tab.py"
with open(tab_path, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("class DailyWorkLogTab", "class KogasDailyWorkLogTab")
content = content.replace("from services.monthly_report_manager import MonthlyReportManager", "from services.kogas_monthly_report_manager import KogasMonthlyReportManager")
content = content.replace("from daily_work_log_exporter import DailyWorkLogExporter", "from kogas_daily_work_log_exporter import KogasDailyWorkLogExporter")
content = content.replace("MonthlyReportManager(", "KogasMonthlyReportManager(")
content = content.replace("DailyWorkLogExporter()", "KogasDailyWorkLogExporter()")

with open(tab_path, 'w', encoding='utf-8') as f:
    f.write(content)

# 2. Update kogas_daily_work_log_exporter.py
exp_path = r"c:\Users\jjch2\Desktop\PMI\home\src\kogas_daily_work_log_exporter.py"
with open(exp_path, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("class DailyWorkLogExporter:", "class KogasDailyWorkLogExporter:")

with open(exp_path, 'w', encoding='utf-8') as f:
    f.write(content)

# 3. Update kogas_monthly_report_manager.py
mon_path = r"c:\Users\jjch2\Desktop\PMI\home\src\services\kogas_monthly_report_manager.py"
with open(mon_path, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("class MonthlyReportManager:", "class KogasMonthlyReportManager:")

with open(mon_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Updated KOGAS specific classes and imports.")
