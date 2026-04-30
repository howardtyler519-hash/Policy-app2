@echo off
title DocuStyle Launcher

set APPDIR=%~dp0
set DS_BAT=%USERPROFILE%\ds\tools\dev-shell-cli\2.4.2\bin\ds.bat

echo.
echo  =========================================
echo   DocuStyle - Starting up...
echo  =========================================
echo.

echo  Checking Python installation...
call "%DS_BAT%" tool install python3.12 >nul 2>&1

echo  Checking required libraries...
pip install -r "%APPDIR%requirements.txt" --quiet >nul 2>&1

echo  Starting DocuStyle...
echo  Keep this window open while using the app.
echo  Close this window to shut the app down.
echo.

streamlit run "%APPDIR%app.py"

pause >nul
