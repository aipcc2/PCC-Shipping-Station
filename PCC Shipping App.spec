# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for the PCC Shipping App — SINGLE-FILE (onefile) build.
# ---------------------------------------------------------------------------
# Produces ONE standalone file: dist\PCC Shipping App.exe
#   * No _internal folder, no external app_logic.py, no logo.png — everything
#     (code + the base64 logo + all libraries) is baked inside the one exe.
#   * Built DIRECTLY from app_logic.py (which has its own
#     `if __name__ == "__main__": ShippingApp().mainloop()` entry point). There
#     is no separate launcher anymore.
#
# Build:   py -3.12 -m PyInstaller --noconfirm --clean "PCC Shipping App.spec"
#   (or just double-click build_exe.bat)
# Output:  dist\PCC Shipping App.exe   (a single ~36 MB file)
#
# To change the app: edit app_logic.py, then rebuild with the command above.

import os
from PyInstaller.utils.hooks import collect_all

APP_NAME = "PCC Shipping App"

# Bundle the real logo.png (from Scripts\) INTO the exe so the app uses it instead
# of the black-background base64 fallback embedded in the code. Keep logo.png in
# Scripts\ when you build. (The exe stays fully self-contained — logo.png is baked
# inside, nothing has to sit next to the exe at runtime.)
datas = []
if os.path.exists('logo.png'):
    datas.append(('logo.png', '.'))
binaries = []

# app_logic.py imports these lazily (inside functions); list them so the
# single-file exe is guaranteed to contain them rather than left to luck.
hiddenimports = [
    'tkinter', 'tkinter.ttk', 'tkinter.filedialog',
    'tkinter.messagebox', 'tkinter.simpledialog',
    'PIL', 'PIL.Image', 'PIL.ImageTk',
    'openpyxl', 'openpyxl.styles', 'openpyxl.utils', 'openpyxl.utils.units',
    'openpyxl.drawing.image', 'openpyxl.drawing.spreadsheet_drawing',
    'openpyxl.drawing.xdr',
    'barcode', 'barcode.writer',
    'docx', 'docx.shared',
    # Customer Portal upload + PDF copies (portal_upload.py, pdf_docs.py sit next to app_logic.py
    # and are compiled INTO the exe like everything else — still one file at runtime).
    'portal_upload', 'pdf_docs',
    'reportlab', 'reportlab.pdfgen.canvas', 'reportlab.graphics.barcode.code128',
    # Monday.com WO auto-lookup networking (stdlib).
    'ssl', '_ssl', 'urllib', 'urllib.request', 'urllib.parse', 'urllib.error',
    'http', 'http.client', 'json', 'threading', 'queue', 'socket', 'subprocess',
]

# Pull in the FULL third-party packages (code + data + submodules).
for _pkg in ('openpyxl', 'barcode', 'PIL', 'docx', 'reportlab'):
    _d, _b, _h = collect_all(_pkg)
    datas += _d
    binaries += _b
    hiddenimports += _h


a = Analysis(
    ['app_logic.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# ONEFILE: pack scripts + binaries + datas all INTO the exe (no COLLECT step).
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                       # UPX not installed; keep off for reproducible builds
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                   # windowed GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
