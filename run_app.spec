# -*- mode: python ; coding: utf-8 -*-


from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import copy_metadata
 
datas = [("C:/Users/yefan/miniforge3/envs/data_merge_advanced/Lib/site-packages/streamlit/runtime","./streamlit/runtime"),('module\\*.*', '.\\module')]
datas += collect_data_files("streamlit")
datas += copy_metadata("streamlit")

block_cipher = None

a = Analysis(
    ['run_app.py'],
    pathex=["D:\audit_project\AUTO_TB\run_app.py"],
    binaries=[],
    datas=datas,
    hiddenimports=['xlwings','pandas','duckdb'],
    hookspath=['./hooks'],
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
    name='试算填写辅助工具_v20250117',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
