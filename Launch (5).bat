@echo off
title DocuStyle Launcher
cd /d "%~dp0"

echo.
echo  =========================================
echo   DocuStyle - Starting up...
echo  =========================================
echo.

set DS_BAT=%USERPROFILE%\ds\tools\dev-shell-cli\2.4.2\bin\ds.bat

echo  Checking Python installation...
call "%DS_BAT%" tool install python3.12 >nul 2>&1

echo  Checking required libraries...
pip install -r requirements.txt --quiet >nul 2>&1

echo  Starting DocuStyle...
echo  Keep this window open while using the app.
echo  Close this window to shut the app down.
echo.

streamlit run app.py

pause >nul
