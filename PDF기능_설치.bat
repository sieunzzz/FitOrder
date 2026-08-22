@echo off
title FitOrder PDF Setup
cd /d "%~dp0"

if not exist "python\python.exe" (
    echo [ERROR] python\python.exe not found.
    echo Copy this file next to FitOrder.exe and run it again.
    pause
    exit /b 1
)

echo Installing PDF support...
"python\python.exe" -m pip install pymupdf
if errorlevel 1 (
    echo.
    echo [FAILED] Check the internet connection and try again.
    pause
    exit /b 1
)

echo.
echo PDF support installed successfully.
pause
