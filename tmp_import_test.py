import sys, traceback, os
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src')
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src')
os.chdir(r'c:\Users\jjch2\Desktop\PMI')

log = open(r'c:\Users\jjch2\Desktop\PMI\app_error.log', 'w', encoding='utf-8')

try:
    log.write('[1] sys.path OK\n')
    log.flush()
    
    # 소스 파일 상단 100라인만 exec해서 import 오류 확인
    src = open(r'c:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src\비파괴검사보고서.py', encoding='utf-8').read()
    log.write(f'[2] 파일 읽기 OK ({len(src)} chars)\n')
    log.flush()
    
    # 컴파일만 시도
    compile(src, 'test', 'exec')
    log.write('[3] 컴파일 OK\n')
    log.flush()

except Exception as e:
    log.write(f'[ERROR] {type(e).__name__}: {e}\n')
    traceback.print_exc(file=log)
finally:
    log.close()
