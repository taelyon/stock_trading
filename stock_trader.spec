# -*- mode: python ; coding: utf-8 -*-
"""
Stock Trader PyInstaller Spec File
대신증권 크레온 초단타 매매 프로그램 v3.0

기능:
- 실시간 자동매매
- 백테스팅 (일별 성과 포함)
- 통합 전략 시스템
- 기술적 지표 분석
"""

block_cipher = None

# ===== 데이터 파일 설정 =====
datas = [
    ('settings.ini.example', '.'),
]

# ===== 숨겨진 import 모듈 =====
hiddenimports = [
    # 자체 모듈 (리팩토링으로 추가됨)
    'strategy_utils',
    'backtester',
    
    # PyQt5
    'PyQt5.QtWidgets',
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtPrintSupport',
    
    # win32com (크레온 API)
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
    
    # matplotlib
    'matplotlib',
    'matplotlib.pyplot',
    'matplotlib.figure',
    'matplotlib.backends',
    'matplotlib.backends.backend_qt5agg',
    'matplotlib.backends.backend_agg',
    'matplotlib.dates',  # 일별 성과 차트에서 사용
    
    # 차트 라이브러리
    'mplfinance',
    'mplfinance.plotting',
    
    # 데이터 분석
    'pandas',
    'pandas.io.formats.style',
    'numpy',
    'numpy.core._methods',
    
    # 기술적 지표
    'talib',
    'talib.stream',
    'talib.abstract',
    
    # 네트워크
    'requests',
    'urllib3',
    'slacker',
    
    # Excel
    'openpyxl',
    'openpyxl.styles',
    
    # 자동화
    'pyautogui',
    'pygetwindow',
    'psutil',
    
    # 기본 모듈
    'sqlite3',
    'configparser',
    'logging',
    'json',
    'threading',
    'queue',
    'collections',
    'traceback',
    'copy',
    'datetime',
    're',
    'os',
    'sys',
    'time',
    
    # pyparsing 의존성
    'html',
    'html.parser',
    'xml',
    'xml.etree.ElementTree',
    'email',
    'urllib',
    'urllib.parse',
    'http',
]

# ===== 제외할 모듈 =====
excludes = [
    'tkinter',
    'IPython',
    'jupyter',
    'notebook',
    'pytest',
    'sphinx',
    'docutils',
]

a = Analysis(
    ['stock_trader.py', 'strategy_utils.py', 'backtester.py'],  # 메인 스크립트 + 추가 모듈
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
    name='stock_trader',
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
    icon='stock_trader.ico',
    uac_admin=True,  # 관리자 권한 필수
)

