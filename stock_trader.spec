# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['stock_trader.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'talib',
        'talib.stream',
        'talib.abstract',
        'numpy',
        'pandas',
        'matplotlib',
        'mplfinance',
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
    name='stock_trader',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI 프로그램이므로 False
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='stock_trader.ico'  # 아이콘 파일 경로
)