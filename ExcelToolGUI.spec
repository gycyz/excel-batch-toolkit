# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

# 拖入支持需要 tkinterdnd2 自带的 tkdnd 二进制
_dnd_datas, _dnd_binaries, _dnd_hidden = collect_all('tkinterdnd2')


a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=_dnd_binaries,
    datas=_dnd_datas,
    hiddenimports=_dnd_hidden,
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
    name='ExcelToolGUI',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ExcelToolGUI',
)
