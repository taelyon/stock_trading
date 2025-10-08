@echo off
chcp 65001 >nul 2>&1

cd /d "%~dp0"

echo.
echo ==========================================
echo  빌드 파일 정리
echo ==========================================
echo.
echo 삭제할 폴더/파일:
echo - build\
echo - dist\
echo - __pycache__\
echo - *.pyc
echo.

choice /C YN /M "계속하시겠습니까?"
if errorlevel 2 goto END
if errorlevel 1 goto CLEAN

:CLEAN
echo.
echo 정리 중...

if exist "build" (
    rmdir /s /q "build"
    echo [완료] build 폴더 삭제
)

if exist "dist" (
    rmdir /s /q "dist"
    echo [완료] dist 폴더 삭제
)

if exist "__pycache__" (
    rmdir /s /q "__pycache__"
    echo [완료] __pycache__ 폴더 삭제
)

del /s /q *.pyc >nul 2>&1
echo [완료] .pyc 파일 삭제

echo.
echo ==========================================
echo  정리 완료!
echo ==========================================

:END
echo.
pause

