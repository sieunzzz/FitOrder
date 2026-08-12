@echo off
title FitOrder - Setup
cd /d "%~dp0"

echo ============================================
echo  FitOrder portable setup (run once)
echo ============================================
echo.

if not exist "python\python.exe" (
    echo [ERROR] python folder not found.
    echo   Expected: %~dp0python\python.exe
    echo   Download cpython-3.12.x-x86_64-pc-windows-msvc-install_only.tar.gz
    echo   from https://github.com/astral-sh/python-build-standalone/releases
    echo   and extract the python folder here.
    pause
    exit /b 1
)

echo [1/3] ensurepip
"python\python.exe" -m ensurepip --upgrade
if errorlevel 1 goto fail

echo.
echo [2/3] installing packages (takes a few minutes)
"python\python.exe" -m pip install --upgrade pip
"python\python.exe" -m pip install streamlit pandas openpyxl pillow openai python-dotenv
if errorlevel 1 goto fail

echo.
echo [3/3] verify
"python\python.exe" -c "import streamlit, pandas, openpyxl, PIL, openai"
if errorlevel 1 goto fail
echo   all packages OK

echo.
echo ============================================
echo  Setup complete. Use run.bat to start.
echo ============================================
pause
exit /b 0

:fail
echo.
echo [FAILED] See the error message above.
pause
exit /b 1