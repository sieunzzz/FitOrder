@echo off
cd /d "%~dp0"
set PYTHONPATH=%CD%\src
python run_fitorder.py
if errorlevel 1 pause
