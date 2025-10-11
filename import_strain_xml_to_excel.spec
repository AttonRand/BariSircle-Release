# -*- mode: python ; coding: utf-8 -*-

"""
PyInstaller spec file for BariSircle Import Script
This file configures the build process for creating a standalone executable.

Usage:
    pyinstaller import_strain_xml_to_excel.spec
"""

block_cipher = None

a = Analysis(
    ['import_strain_xml_to_excel.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('BariSircle.ico', '.'),  # Include icon file in the bundle
    ],
    hiddenimports=[
        'xml.etree.ElementTree',
        'pandas',
        'openpyxl',
        'openpyxl.cell._writer',
        'openpyxl.styles.stylesheet',
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
    name='BariSircle_Import',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Console window for this script (interactive prompts)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='BariSircle.ico',  # Application icon
    version_file=None,
)
