# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import sys

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

HOME = Path(SPECPATH).parent
SRC = HOME / "src"
PACKAGE = "site_apps.kogas"
sys.path.insert(0, str(SRC))

a = Analysis(
    [str(SRC / "site_apps" / "kogas" / "main.py")],
    pathex=[str(SRC)],
    binaries=[],
    datas=collect_data_files(PACKAGE),
    hiddenimports=collect_submodules(PACKAGE),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["IPython", "jupyter", "matplotlib", "notebook", "pytest", "scipy"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="SITCO-KOGAS-Material-Manager",
    debug=False, bootloader_ignore_signals=False, strip=False, upx=True,
    console=False, disable_windowed_traceback=False,
)
