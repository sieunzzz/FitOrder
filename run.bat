@echo off
title FitOrder
cd /d "%~dp0"

if not exist "python\python.exe" (
    echo python folder not found. Run setup_portable.bat first.
    pause
    exit /b 1
)

if not exist ".font_ok" (
    if exist "THEHappyfruit.ttf" (
        echo.
        echo  [NOTICE] Install THEHappyfruit.ttf first.
        echo    Right-click the file - "Install for all users"
        echo    Press any key when done.
        echo.
        pause
        echo ok > .font_ok
    )
)

start "" chrome --app=http://localhost:8501 --window-size=1600,950
if errorlevel 1 start "" http://localhost:8501

cd src
"..\python\python.exe" -m streamlit run app.py --server.headless true --server.port 8501
pause