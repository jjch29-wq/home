# -*- mode: python ; coding: utf-8 -*-
# SITCO Material Master Manager — PyInstaller spec
# DB 파일(Material_Inventory.xlsx) 미포함: exe 옆에 자동 생성됨

block_cipher = None

a = Analysis(
    ['src/Material-Master-Manager-V13.py'],
    pathex=[],
    binaries=[],
    datas=[
        # DB 파일은 의도적으로 제외 (첫 저장 시 exe 옆에 자동 생성)
        # config 파일도 제외 (Documents/MaterialManager/ 에 자동 생성)
    ],
    hiddenimports=[
        # pandas / openpyxl / xlsxwriter
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.workbook',
        'xlsxwriter',
        # tkcalendar & babel (locale 관련)
        'tkcalendar',
        'babel',
        'babel.numbers',
        'babel.dates',
        'babel.core',
        # PIL
        'PIL',
        'PIL._imagingtk',
        'PIL.Image',
        'PIL.ImageTk',
        # numpy
        'numpy',
        'numpy.core._multiarray_umath',
        # lxml (openpyxl 선택적 의존)
        'lxml',
        'lxml.etree',
        # 기타
        'pkg_resources.py2_warn',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'IPython',
        'jupyter',
        'notebook',
        'pytest',
        'setuptools',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SITCO-Material-Manager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # 콘솔 창 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='resources/icon.ico',  # 아이콘 파일 있으면 주석 해제
)
