# -*- mode: python ; coding: utf-8 -*-

"""
PyInstaller spec file for BariSircle GUI Application
This file configures the build process for creating a standalone executable.

Usage:
    pyinstaller xml_to_excel_mapper_gui.spec
"""

block_cipher = None

a = Analysis(
    ['xml_to_excel_mapper_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('BariSircle.ico', '.'),  # Include icon file in the bundle
    ],
    hiddenimports=[
        'xml.etree.ElementTree',
        'openpyxl',
        'openpyxl.cell._writer',
        'openpyxl.styles.stylesheet',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
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
    name='BariSircle',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window (GUI only)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='BariSircle.ico',  # Application icon for the executable and taskbar
    version_file=None,
)
