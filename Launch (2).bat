@echo off
title DocuStyle Launcher
cd /d "%~dp0"

echo.
echo  =========================================
echo   DocuStyle - Starting up...
echo  =========================================
echo.

:: Install Python 3.12 via ds tool
set DS_BAT=%USERPROFILE%\ds\tools\dev-shell-cli\2.4.2\bin\ds.bat
echo  Checking Python installation...
call "%DS_BAT%" tool install python3.12

:: Set Python path
set PYTHON_EXE=%USERPROFILE%\ds\tools\python3.12\latest\python.exe

:: Install required libraries
echo  Checking required libraries...
"%PYTHON_EXE%" -m pip install -r requirements.txt --quiet

:: Launch the app
echo  Launching DocuStyle...
echo  (A browser window will open automatically)
echo  (To close the app, close this window)
echo.
start "" /min "%PYTHON_EXE%" -m streamlit run app.py --server.headless=true --browser.gatherUsageStats=false
timeout /t 3 >nul
start "" http://localhost:8501
exit
