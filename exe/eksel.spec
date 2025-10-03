# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['..\\eksel.py'],
    pathex=[],
    binaries=[],
    datas=[('..\\assets', 'assets')],
    hiddenimports=[
        'PIL._tkinter_finder',
        'pyperclip',
        'win32com.client',
        'pythoncom',
        'xlwings',
        'pywintypes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='eksel',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['..\\assets\\split.ico'],
)
