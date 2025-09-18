# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('symbol.csv', '.'),
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'pandas._libs.tslibs.base',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.offsets',
        'pandas._libs.tslibs.parsing',
        'pandas._libs.tslibs.period',
        'pandas._libs.tslibs.timestamps',
        'pandas._libs.tslibs.timezones',
        'pandas._libs.tslibs.tzconversion',
        'pandas._libs.tslibs.fields',
        'pandas._libs.tslibs.resolution',
        'pandas._libs.window.aggregations',
        'pandas._libs.window.indexers',
        'openpyxl.cell._writer',
        'openpyxl.worksheet._reader',
        'openpyxl.worksheet._write_only',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='pt-report-to-csv',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    icon='logo.ico'
)