# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_data_files

datas = [
    (os.path.join(SPECPATH, '.env'),             '.'),
    (os.path.join(SPECPATH, 'credentials.json'), '.'),
    (os.path.join(SPECPATH, 'icon.ico'),         '.'),
]
datas += collect_data_files('customtkinter')
datas += collect_data_files('playwright')

a = Analysis(
    [os.path.join(SPECPATH, 'gui.py')],
    pathex=[SPECPATH],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'gspread',
        'paramiko',
        'keyring',
        'keyring.backends',
        'keyring.backends.Windows',
        'google_auth_oauthlib',
        'google.auth.transport.requests',
    ],
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
    name='AHA Bot',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    icon=os.path.join(SPECPATH, 'icon.ico'),
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
    name='AHA Bot',
)
