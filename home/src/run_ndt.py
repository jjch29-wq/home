# -*- coding: utf-8 -*-
"""ASCII-named entry point for 비파괴검사보고서 app.
CMD/bat 파일에서 한글 파일명 인코딩 문제를 우회하기 위한 런처.
"""
from pathlib import Path
import runpy
import sys
import traceback

APP_SRC = Path(__file__).resolve().parent / "report_apps" / "ndt_report" / "src"
SRC_ROOT = Path(__file__).resolve().parent
CRASH_LOG = Path(__file__).resolve().parent.parent.parent / "crash.log"

sys.path.insert(0, str(SRC_ROOT))
sys.path.insert(0, str(APP_SRC))

try:
    runpy.run_path(str(APP_SRC / "비파괴검사보고서.py"), run_name="__main__")
except Exception as e:
    with open(CRASH_LOG, "w", encoding="utf-8") as f:
        f.write(f"[CRASH] {type(e).__name__}: {e}\n\n")
        traceback.print_exc(file=f)
    raise
except SystemExit as e:
    if e.code not in (0, None):
        with open(CRASH_LOG, "w", encoding="utf-8") as f:
            f.write(f"[SystemExit] code={e.code}\n")
            traceback.print_exc(file=f)
    raise
