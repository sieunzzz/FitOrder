@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
title FitOrder v81 - New PC Setup

echo ==================================================
echo FitOrder v81 new-PC setup
echo ==================================================
echo.

call :find_python
if defined PYEXE goto :have_python

echo [INFO] A usable Python installation was not found.
echo [INFO] Trying to install Python 3.12 with winget...
echo.
where winget >nul 2>nul
if errorlevel 1 goto :no_winget

winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
if errorlevel 1 goto :python_install_failed

call :find_python
if not defined PYEXE goto :python_install_failed

:have_python
echo [OK] Python found:
echo %PYEXE%
echo.
"%PYEXE%" "%~dp0setup_fitorder.py"
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo.
  echo [FAILED] Dependency setup failed. Error code: %RC%
  echo Please send a screenshot of this window.
  echo.
  pause
  exit /b %RC%
)

>"%~dp0.fitorder_python.txt" echo %PYEXE%
echo.
echo [OK] FitOrder v81 setup completed.
echo Run 01_RUN_FITORDER.bat next.
echo.
pause
exit /b 0

:no_winget
echo.
echo [ERROR] Python is not installed and winget is unavailable.
echo Install 64-bit Python 3.12 from python.org and check "Add python.exe to PATH".
echo Then run 00_INSTALL_FIRST.bat again.
echo.
pause
exit /b 1

:python_install_failed
echo.
echo [ERROR] Automatic Python installation failed.
echo Install 64-bit Python 3.12 manually, then run 00_INSTALL_FIRST.bat again.
echo.
pause
exit /b 1

:find_python
set "PYEXE="
if exist "%~dp0python\python.exe" (
  "%~dp0python\python.exe" -c "import sys; print(sys.executable)" >nul 2>nul
  if not errorlevel 1 set "PYEXE=%~dp0python\python.exe"
)
if defined PYEXE exit /b 0

for %%V in (314 313 312 311) do (
  if not defined PYEXE if exist "%LocalAppData%\Programs\Python\Python%%V\python.exe" (
    "%LocalAppData%\Programs\Python\Python%%V\python.exe" -c "import sys" >nul 2>nul
    if not errorlevel 1 set "PYEXE=%LocalAppData%\Programs\Python\Python%%V\python.exe"
  )
)
if defined PYEXE exit /b 0

where py >nul 2>nul
if not errorlevel 1 (
  for /f "delims=" %%P in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
)
if defined PYEXE exit /b 0

where python >nul 2>nul
if not errorlevel 1 (
  for /f "delims=" %%P in ('python -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
)
if defined PYEXE (
  "%PYEXE%" -c "import sys" >nul 2>nul
  if errorlevel 1 set "PYEXE="
)
exit /b 0
