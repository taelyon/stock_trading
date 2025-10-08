@echo off
chcp 65001 >nul 2>&1

cd /d "%~dp0"

echo.
echo ==========================================
echo  Stock Trader 디버그 빌드
echo ==========================================
echo.
echo 콘솔 창이 표시되어 오류를 확인할 수 있습니다.
echo.

if not exist "stock_trader_debug.spec" (
    echo [오류] stock_trader_debug.spec를 찾을 수 없습니다.
    pause
    exit /b 1
)

echo 이전 빌드 정리 중...
if exist "build" rmdir /s /q "build" 2>nul
if exist "dist\stock_trader_debug.exe" del /q "dist\stock_trader_debug.exe" 2>nul
echo.

echo 디버그 빌드 시작...
echo.
pyinstaller stock_trader_debug.spec

echo.
if exist "dist\stock_trader_debug.exe" (
    echo ==========================================
    echo  빌드 성공!
    echo ==========================================
    echo.
    echo 실행 파일: dist\stock_trader_debug.exe
    echo.
    
    choice /C YN /M "지금 실행하시겠습니까?"
    if errorlevel 2 goto END
    if errorlevel 1 (
        echo.
        echo 실행 중... 콘솔 창에서 오류를 확인하세요.
        echo.
        cd dist
        stock_trader_debug.exe
        cd ..
    )
) else (
    echo ==========================================
    echo  빌드 실패
    echo ==========================================
)

:END
pause

