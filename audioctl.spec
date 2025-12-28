# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

# This spec file is configured for a --onefile build.
# The key is that there is an EXE block but no COLLECT block at the end.

# Corresponds to: --add-data "audio.ico;."
datas = [('audio.ico', '.')]
binaries = []

# Corresponds to all the --hidden-import flags
hiddenimports = [
    'comtypes.automation',
    'comtypes._post_coinit',
    'comtypes._post_coinit.unknwn',
    'comtypes._post_coinit.misc'
]

# Corresponds to: --collect-all pycaw
tmp_ret = collect_all('pycaw')
datas += tmp_ret[0]
binaries += tmp_ret[1]
hiddenimports += tmp_ret[2]

# Corresponds to: --collect-all comtypes
tmp_ret = collect_all('comtypes')
datas += tmp_ret[0]
binaries += tmp_ret[1]
hiddenimports += tmp_ret[2]

a = Analysis(
    ['audioctl.py'], # The main script
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
    a.binaries,
    a.datas,
    [],
    # Corresponds to: --name audioctl
    name='audioctl',
    debug=False,
    # Corresponds to: --bootloader-ignore-signals
    bootloader_ignore_signals=True,
    strip=False,
    # Corresponds to: --noupx
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    # Corresponds to: --console
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # Corresponds to: --version-file version.txt
    version='version.txt',
    # Corresponds to: --icon audio.ico
    icon=['audio.ico'],
)
