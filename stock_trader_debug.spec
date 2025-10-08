# -*- mode: python ; coding: utf-8 -*-
"""
Stock Trader PyInstaller Spec File (DEBUG VERSION)
오류 확인을 위한 디버그 모드 빌드
"""

block_cipher = None

datas = [
    ('settings.ini.example', '.'),
]

hiddenimports = [
    'PyQt5.QtWidgets',
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtPrintSupport',
    'win32com',
    'win32com.client',
    'win32com.client.gencache',
    'win32com.client.dynamic',
    'pythoncom',
    'pywintypes',
    'win32api',
    'win32con',
    'win32gui',
    'win32process',
    'win32event',
    'win32file',
    'matplotlib',
    'matplotlib.pyplot',
    'matplotlib.figure',
    'matplotlib.backends',
    'matplotlib.backends.backend_qt5agg',
    'matplotlib.backends.backend_agg',
    'mplfinance',
    'mplfinance.plotting',
    'pandas',
    'pandas.io.formats.style',
    'numpy',
    'numpy.core._methods',
    'talib',
    'talib.stream',
    'talib.abstract',
    'requests',
    'urllib3',
    'slacker',
    'openpyxl',
    'openpyxl.styles',
    'pyautogui',
    'pygetwindow',
    'psutil',
    'sqlite3',
    'configparser',
    'logging',
    'json',
    'threading',
    'queue',
    'collections',
    'traceback',
    'copy',
    'html',
    'html.parser',
    'xml',
    'xml.etree.ElementTree',
    'email',
    'urllib',
    'urllib.parse',
    'http',
]

excludes = [
    'tkinter',
    'IPython',
    'jupyter',
    'notebook',
    'pytest',
]

a = Analysis(
    ['stock_trader.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
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
    name='stock_trader_debug',
    debug=True,           # 디버그 모드
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,           # 디버그 시 압축 비활성화
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,        # 콘솔 창 표시
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='stock_trader.ico',
    uac_admin=True,
)

