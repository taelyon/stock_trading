@echo off
chcp 65001 >nul 2>&1

cd /d "%~dp0"

echo.
echo ==========================================
echo  Stock Trader 빌드
echo ==========================================
echo.

if not exist "stock_trader.py" (
    echo [오류] stock_trader.py를 찾을 수 없습니다.
    pause
    exit /b 1
)

if not exist "stock_trader.spec" (
    echo [오류] stock_trader.spec를 찾을 수 없습니다.
    pause
    exit /b 1
)

echo 이전 빌드 정리 중...
if exist "build" rmdir /s /q "build" 2>nul
if exist "dist\stock_trader.exe" del /q "dist\stock_trader.exe" 2>nul
echo.

echo 빌드 시작...
echo.
pyinstaller stock_trader.spec

echo.
if exist "dist\stock_trader.exe" (
    echo ==========================================
    echo  빌드 성공!
    echo ==========================================
    echo.
    echo 실행 파일: dist\stock_trader.exe
    echo.
    
    for %%A in ("dist\stock_trader.exe") do (
        set size=%%~zA
        set /a sizeMB=%%~zA/1024/1024
    )
    echo 파일 크기: !sizeMB! MB
    echo.
    
    explorer "dist"
) else (
    echo ==========================================
    echo  빌드 실패
    echo ==========================================
    echo.
)

pause

