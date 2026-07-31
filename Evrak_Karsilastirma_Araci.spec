# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules
import os

hiddenimports = []
hiddenimports += collect_submodules('openpyxl')

root = os.path.abspath('.')
locale_datas = [
    (os.path.join(root, 'locales', 'tr.json'), 'locales'),
    (os.path.join(root, 'locales', 'ru.json'), 'locales'),
    (os.path.join(root, 'locales', 'de.json'), 'locales'),
    (os.path.join(root, 'locales', 'en.json'), 'locales'),
]

a = Analysis(
    ['karsilastir.py'],
    pathex=[],
    binaries=[],
    datas=locale_datas,
    hiddenimports=hiddenimports + [
        'i18n', 'app_log', 'karsilastir_motor',
        'report_extra', 'batch_compare', 'archive_db',
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
    a.binaries,
    a.datas,
    [],
    name='Evrak_Karsilastirma_Araci',
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
