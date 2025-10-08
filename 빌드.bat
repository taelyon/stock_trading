@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

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

if not exist "strategy_utils.py" (
    echo [오류] strategy_utils.py를 찾을 수 없습니다.
    pause
    exit /b 1
)

if not exist "backtester.py" (
    echo [오류] backtester.py를 찾을 수 없습니다.
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
echo [INFO] pyinstaller stock_trader.spec 실행 중...
echo.

pyinstaller stock_trader.spec
set BUILD_ERROR=%ERRORLEVEL%

echo.
echo [DEBUG] pyinstaller 종료 코드: %BUILD_ERROR%
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
    echo [오류] 실행 파일이 생성되지 않았습니다.
    echo.
    echo 가능한 원인:
    echo 1. pyinstaller 종료 코드: %BUILD_ERROR%
    echo 2. 위 로그에서 ERROR 또는 CRITICAL 메시지 확인
    echo 3. 누락된 모듈이나 의존성 문제
    echo.
    echo 해결 방법:
    echo - 수동 실행: pyinstaller stock_trader.spec
    echo - 로그 확인: 위 출력 내용 검토
    echo - BUILD.md 문서의 문제 해결 섹션 참조
    echo.
)

pause

