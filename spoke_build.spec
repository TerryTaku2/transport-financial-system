# Builds a standalone TransportERP.exe for a spoke PC that has no Python
# installed at all. Rebuild after any app.py/requirements.txt change with:
#   venv\Scripts\python.exe -m PyInstaller spoke_build.spec
# Output lands in dist\TransportERP\ -- copy that whole folder to the spoke
# PC (see the onboarding runbook), it does not need to stay in this repo.
#
# --collect-all is used for flask_limiter and reportlab specifically
# because both failed with ModuleNotFoundError at runtime under plain
# auto-detection during real testing (reportlab's import wasn't picked up
# by static analysis at all; flask_limiter needed its full package, not
# just the top-level module) -- if a future dependency is added and the
# built exe crashes with ModuleNotFoundError for it, add it here the same
# way rather than guessing at a narrower --hidden-import.
# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('templates', 'templates'), ('static', 'static')]
binaries = []
hiddenimports = []
tmp_ret = collect_all('flask_limiter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('reportlab')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['app.py'],
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

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TransportERP',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
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
    name='TransportERP',
)
