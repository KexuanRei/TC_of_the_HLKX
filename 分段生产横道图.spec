# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['server', 'panel', 'core', 'build', 'gantt_sync', 'paths', 'net', 'webview',
                 'kdocs_download', 'playwright.sync_api']
hiddenimports += collect_submodules('openpyxl')


a = Analysis(
    ['window.py'],
    pathex=[],
    binaries=[],
    datas=[('app.html', '.'), ('viewer.html', '.'), ('data.js', '.'), ('config.json', '.'), ('gantt_config.json', '.'), ('TM.ico', '.'), ('wps_gantt.js', '.'), ('airscript_sync_task.js', '.'), ('README.md', '.')],
    hiddenimports=hiddenimports,
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
    name='分段生产横道图',
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
    icon=['TM.ico'],
)
