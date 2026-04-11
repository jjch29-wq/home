# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['src/main.py'],
    pathex=['c:/Users/jjch2/Desktop/보고서/Project PROVIDENCE/Request/PMI'],
    binaries=[],
    datas=[
        ('resources', 'resources'),
        ('config', 'config'),
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'openpyxl.cell.cell',
        'openpyxl.worksheet.pagebreak',
        'openpyxl.drawing.image',
        'openpyxl.worksheet.datavalidation',
        'openpyxl.styles',
        'openpyxl.drawing.spreadsheet_drawing',
        'openpyxl.drawing.xdr',
        'openpyxl.utils.cell',
        'xlsxwriter',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'PIL.ImageChops',
        'lxml',
        'lxml._elementpath',
        'lxml.etree',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.simpledialog',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='SITCO-Report-Manager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 콘솔 창 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
