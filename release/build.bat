@echo off
title FitOrder - Build
cd /d "%~dp0"

echo ============================================
echo  FitOrder exe build
echo ============================================
echo.

if not exist "..\python\python.exe" (
    echo [ERROR] portable python not found at ..\python\python.exe
    pause
    exit /b 1
)

echo [1/3] installing pyinstaller
"..\python\python.exe" -m pip install pyinstaller
if errorlevel 1 goto fail

echo.
echo [2/3] building (this takes several minutes)
"..\python\python.exe" -m PyInstaller FitOrder.spec --noconfirm --clean
if errorlevel 1 goto fail

echo.
echo [3/3] copying app files
xcopy /E /I /Y "..\src" "dist\FitOrder\src" >nul
xcopy /E /I /Y "..\data" "dist\FitOrder\data" >nul
if exist "..\.env" copy /Y "..\.env" "dist\FitOrder\.env" >nul
if exist "..\THEHappyfruit.ttf" copy /Y "..\THEHappyfruit.ttf" "dist\FitOrder\" >nul

echo.
echo ============================================
echo  Done.  dist\FitOrder\FitOrder.exe
echo ============================================
pause
exit /b 0

:fail
echo.
echo [FAILED] See the error above.
pause
exit /b 1
