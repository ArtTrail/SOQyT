# -*- mode: python ; coding: utf-8 -*-
import sys

a = Analysis(
    ['star_query_tool.py'],
    pathex=[],
    binaries=[],
    datas=[('SOQyT.ico', '.')],
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name='Stellar Object Query Tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='SOQyT.ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Stellar Object Query Tool',
)

if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='Stellar Object Query Tool.app',
        icon='SOQyT.icns',
        bundle_identifier='com.arttrail.soqyt',
    )
