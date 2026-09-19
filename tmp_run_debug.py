import traceback, sys, os
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src')
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src')
try:
    import runpy
    runpy.run_path(r'c:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src\비파괴검사보고서.py', run_name='__main__')
except Exception:
    with open(r'c:\Users\jjch2\Desktop\PMI\app_error.log', 'w', encoding='utf-8') as f:
        traceback.print_exc(file=f)
    raise
