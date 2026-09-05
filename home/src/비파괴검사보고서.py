# -*- coding: utf-8 -*-
"""Launcher for the NDT report app.

The application code lives under report_apps/ndt_report so report-specific
modules can be changed without mixing them with site app code.
"""
from pathlib import Path
import runpy
import sys


APP_SRC = Path(__file__).resolve().parent / "report_apps" / "ndt_report" / "src"
SRC_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(SRC_ROOT))
sys.path.insert(0, str(APP_SRC))
runpy.run_path(str(APP_SRC / "비파괴검사보고서.py"), run_name="__main__")
