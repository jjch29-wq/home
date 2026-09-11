# -*- coding: utf-8 -*-
"""Compatibility launcher for the PAUT capture application.

The implementation lives under ``tools/paut_capture``. Keep this launcher so
existing shortcuts and IDE run targets continue to work.
"""

from pathlib import Path
import runpy
import sys


APP_SRC = Path(__file__).resolve().parent / "tools" / "paut_capture" / "src"
sys.path.insert(0, str(APP_SRC))
runpy.run_path(str(APP_SRC / "paut_capture.py"), run_name="__main__")
