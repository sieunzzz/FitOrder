@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
title FitOrder v81 - PySide6

set "PYEXE="
if exist "%~dp0.fitorder_python.txt" (
  set /p PYEXE=<"%~dp0.fitorder_python.txt"
  if defined PYEXE (
    "%PYEXE%" -c "import sys" >nul 2>nul
    if errorlevel 1 set "PYEXE="
  )
)

if not defined PYEXE if exist "%LocalAppData%\Programs\Python\Python314\python.exe" set "PYEXE=%LocalAppData%\Programs\Python\Python314\python.exe"
if not defined PYEXE if exist "%LocalAppData%\Programs\Python\Python313\python.exe" set "PYEXE=%LocalAppData%\Programs\Python\Python313\python.exe"
if not defined PYEXE if exist "%LocalAppData%\Programs\Python\Python312\python.exe" set "PYEXE=%LocalAppData%\Programs\Python\Python312\python.exe"
if not defined PYEXE if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set "PYEXE=%LocalAppData%\Programs\Python\Python311\python.exe"

if not defined PYEXE (
  echo [SETUP] Python setup is required.
  call "%~dp0install_FitOrder.bat"
  if errorlevel 1 exit /b 1
  if exist "%~dp0.fitorder_python.txt" set /p PYEXE=<"%~dp0.fitorder_python.txt"
)

"%PYEXE%" "%~dp0setup_fitorder.py" --check >nul 2>nul
if errorlevel 1 (
  echo [SETUP] Required packages are missing.
  call "%~dp0install_FitOrder.bat"
  if errorlevel 1 exit /b 1
)

"%PYEXE%" "%~dp0src\qt_app.py"
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo.
  echo [ERROR] FitOrder stopped with error code %RC%.
  echo Please send a screenshot of the error above.
  echo.
  pause
)
exit /b %RC%
