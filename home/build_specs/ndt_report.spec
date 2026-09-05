# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files

HOME = Path(SPECPATH).parent
SRC = HOME / "src"
NDT_SRC = SRC / "report_apps" / "ndt_report" / "src"

datas = [
    (str(HOME / "config" / "logo_settings_unified.json"), "."),
    (str(HOME / "resources"), "resources"),
]

a = Analysis(
    [str(NDT_SRC / "비파괴검사보고서.py")],
    pathex=[str(SRC), str(NDT_SRC)],
    binaries=[],
    datas=datas + collect_data_files("babel") + collect_data_files("tkcalendar"),
    hiddenimports=[
        "babel", "babel.core", "babel.dates", "babel.numbers", "babel.plural",
        "lxml", "lxml.etree", "openpyxl", "openpyxl.styles", "openpyxl.utils",
        "PIL", "PIL._imagingtk", "PIL.Image", "PIL.ImageTk", "tkcalendar",
        "xlsxwriter", "ndt_history_import",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["IPython", "jupyter", "matplotlib", "notebook", "pytest", "scipy"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="SITCO-NDT-Report",
    debug=False, bootloader_ignore_signals=False, strip=False, upx=True,
    console=False, disable_windowed_traceback=False,
)
