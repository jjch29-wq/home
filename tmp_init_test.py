import sys, traceback, os
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src')
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src')

log = open(r'c:\Users\jjch2\Desktop\PMI\app_error.log', 'w', encoding='utf-8')
sys.stdout = log
sys.stderr = log

try:
    import tkinter as tk
    log.write('[OK] tkinter import\n'); log.flush()
    
    src_path = r'c:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src\비파괴검사보고서.py'
    g = {'__name__': '__main__', '__file__': src_path}
    
    src = open(src_path, encoding='utf-8').read()
    # __main__ 블록 제거하고 클래스 정의만 실행
    import re
    src_no_main = re.sub(r"if __name__\s*==\s*['\"]__main__['\"].*", '', src, flags=re.DOTALL)
    log.write('[OK] 소스 준비\n'); log.flush()
    
    exec(compile(src_no_main, src_path, 'exec'), g)
    log.write('[OK] exec 완료 - 클래스 로딩 성공\n'); log.flush()
    
    root = tk.Tk()
    log.write('[OK] tk.Tk() 생성\n'); log.flush()
    
    PMIReportApp = g.get('PMIReportApp')
    if PMIReportApp:
        app = PMIReportApp(root)
        log.write('[OK] PMIReportApp 초기화 성공\n'); log.flush()
    else:
        log.write('[ERR] PMIReportApp 클래스를 찾을 수 없음\n')
    
    root.after(500, root.destroy)
    root.mainloop()
    log.write('[OK] mainloop 완료\n')

except Exception as e:
    log.write(f'[ERROR] {type(e).__name__}: {e}\n')
    traceback.print_exc(file=log)
finally:
    log.flush()
    log.close()
