import sys
import os
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'home', 'src'))

from services.monthly_report_manager import MonthlyReportManager

manager = MonthlyReportManager(r"C:\Users\jjch2\Desktop\템플릿_최종완성본_V70.xlsx")
history_path = r"c:\Users\jjch2\Desktop\PMI\home\src\daily_work_history.json"
output_path = r"c:\Users\jjch2\Desktop\PMI\Test_MonthlyReport_2026-08.xlsx"

try:
    print("Generating report...")
    manager.generate_report(history_path, "2026-08", output_path)
    print("Success. Saved to:", output_path)
except Exception as e:
    import traceback
    traceback.print_exc()
