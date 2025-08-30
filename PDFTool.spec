# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['pdf_tool\\app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pdf2docx', 'pdf2docx.converter', 'pdf2docx.parser', 'fitz', 'PIL', 'PIL.Image', 'pypdf', 'pypdf.generic', 'pypdf.reader', 'pypdf.writer'],
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
    name='PDFTool',
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
