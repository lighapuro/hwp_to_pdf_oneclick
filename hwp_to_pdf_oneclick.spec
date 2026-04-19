# -*- mode: python ; coding: utf-8 -*-
import os
import sys
import tkinterdnd2

_tkdnd_src = os.path.join(os.path.dirname(tkinterdnd2.__file__), "tkdnd", "win-x64")

a = Analysis(
    ["hwp_to_pdf_oneclick.py"],
    pathex=[],
    binaries=[],
    datas=[
        (_tkdnd_src, "tkinterdnd2/tkdnd/win-x64"),
    ],
    hiddenimports=[
        "tkinterdnd2",
        "win32com",
        "win32com.client",
        "win32com.server",
        "pythoncom",
        "pywintypes",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="HWP_PDF변환기",
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
)
