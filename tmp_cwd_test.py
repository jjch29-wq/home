import sys, os, traceback
os.chdir(r'c:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src')
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src')
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src')

try:
    with open(r'c:\Users\jjch2\Desktop\PMI\app_error.log', 'w', encoding='utf-8') as _log:
        _log.write('[START]\n')
    exec(open(r'c:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src\비파괴검사보고서.py', encoding='utf-8').read())
except Exception as e:
    with open(r'c:\Users\jjch2\Desktop\PMI\app_error.log', 'a', encoding='utf-8') as _log:
        _log.write(f'[ERROR] {type(e).__name__}: {e}\n')
        traceback.print_exc(file=_log)
