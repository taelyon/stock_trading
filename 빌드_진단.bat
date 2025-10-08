@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo.
echo ==========================================
echo  빌드 환경 진단
echo ==========================================
echo.

REM ===== 1. Python 확인 =====
echo [1] Python 버전 확인
python --version
if %ERRORLEVEL% neq 0 (
    echo [오류] Python이 설치되어 있지 않거나 PATH에 없습니다.
    pause
    exit /b 1
)
echo.

REM ===== 2. PyInstaller 확인 =====
echo [2] PyInstaller 확인
pyinstaller --version
if %ERRORLEVEL% neq 0 (
    echo [오류] PyInstaller가 설치되어 있지 않습니다.
    echo [해결] pip install pyinstaller
    pause
    exit /b 1
)
echo.

REM ===== 3. 필수 파일 확인 =====
echo [3] 필수 파일 확인
set ALL_FILES_OK=1

if not exist "stock_trader.py" (
    echo [X] stock_trader.py - 없음
    set ALL_FILES_OK=0
) else (
    echo [O] stock_trader.py
)

if not exist "strategy_utils.py" (
    echo [X] strategy_utils.py - 없음
    set ALL_FILES_OK=0
) else (
    echo [O] strategy_utils.py
)

if not exist "backtester.py" (
    echo [X] backtester.py - 없음
    set ALL_FILES_OK=0
) else (
    echo [O] backtester.py
)

if not exist "stock_trader.spec" (
    echo [X] stock_trader.spec - 없음
    set ALL_FILES_OK=0
) else (
    echo [O] stock_trader.spec
)

if not exist "stock_trader.ico" (
    echo [경고] stock_trader.ico - 없음 (빌드는 가능하지만 아이콘 없음)
) else (
    echo [O] stock_trader.ico
)

if not exist "settings.ini.example" (
    echo [경고] settings.ini.example - 없음 (빌드는 가능)
) else (
    echo [O] settings.ini.example
)

echo.

if %ALL_FILES_OK%==0 (
    echo [오류] 필수 파일이 누락되었습니다.
    pause
    exit /b 1
)

REM ===== 4. Python 패키지 확인 =====
echo [4] 주요 Python 패키지 확인
python -c "import PyQt5; print('PyQt5:', PyQt5.__version__)" 2>nul
if %ERRORLEVEL% neq 0 echo [X] PyQt5 - 없음

python -c "import win32com.client; print('pywin32: OK')" 2>nul
if %ERRORLEVEL% neq 0 echo [X] pywin32 - 없음

python -c "import pandas; print('pandas:', pandas.__version__)" 2>nul
if %ERRORLEVEL% neq 0 echo [X] pandas - 없음

python -c "import numpy; print('numpy:', numpy.__version__)" 2>nul
if %ERRORLEVEL% neq 0 echo [X] numpy - 없음

python -c "import matplotlib; print('matplotlib:', matplotlib.__version__)" 2>nul
if %ERRORLEVEL% neq 0 echo [X] matplotlib - 없음

python -c "import talib; print('TA-Lib: OK')" 2>nul
if %ERRORLEVEL% neq 0 echo [경고] talib - 없음 (빌드 실패 가능성 높음)

echo.

REM ===== 5. import 테스트 =====
echo [5] 자체 모듈 import 테스트
python -c "import strategy_utils; print('[O] strategy_utils - import 성공')" 2>nul
if %ERRORLEVEL% neq 0 (
    echo [X] strategy_utils - import 실패
    echo [상세] Python 파일에 문법 오류가 있을 수 있습니다.
)

python -c "import backtester; print('[O] backtester - import 성공')" 2>nul
if %ERRORLEVEL% neq 0 (
    echo [X] backtester - import 실패
    echo [상세] Python 파일에 문법 오류가 있을 수 있습니다.
)

echo.

REM ===== 6. 디스크 공간 확인 =====
echo [6] 디스크 공간 확인
for /f "tokens=3" %%a in ('dir /-c ^| find "남은 바이트"') do set FREE_SPACE=%%a
echo 남은 공간: %FREE_SPACE% bytes
echo (빌드에는 최소 500MB 권장)
echo.

REM ===== 7. 권한 확인 =====
echo [7] 관리자 권한 확인
net session >nul 2>&1
if %ERRORLEVEL% == 0 (
    echo [O] 관리자 권한으로 실행 중
) else (
    echo [경고] 관리자 권한 없음 (일부 빌드 실패 가능성)
)
echo.

REM ===== 8. 이전 빌드 확인 =====
echo [8] 이전 빌드 파일 확인
if exist "build\" (
    echo [O] build 폴더 존재 (정리 권장: 정리.bat)
) else (
    echo [O] build 폴더 없음 (깨끗한 상태)
)

if exist "dist\stock_trader.exe" (
    echo [O] 이전 실행 파일 존재
    for %%A in ("dist\stock_trader.exe") do (
        set size=%%~zA
        set /a sizeMB=%%~zA/1024/1024
    )
    echo    크기: !sizeMB! MB
) else (
    echo [O] 이전 실행 파일 없음
)
echo.

REM ===== 결과 요약 =====
echo ==========================================
echo  진단 완료
echo ==========================================
echo.
echo 위 결과를 확인하여 문제가 있다면 해결 후 빌드하세요.
echo.
echo 다음 단계:
echo 1. 정리.bat (선택사항)
echo 2. 빌드.bat
echo.

pause
