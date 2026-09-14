@echo off
cd /d "%~dp0src"
python -m streamlit run app.py
if errorlevel 1 pause
