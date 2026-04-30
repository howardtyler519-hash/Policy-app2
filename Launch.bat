@echo off
title DocuStyle Launcher

echo.
echo  =========================================
echo   DocuStyle - Starting up...
echo  =========================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  Python not found. Opening download page...
    echo  Please install Python from python.org
    echo  IMPORTANT: Check "Add Python to PATH" during install.
    echo  Then double-click this file again.
    start https://www.python.org/downloads/
    pause
    exit
)

echo  Checking required libraries...

:: Install dependencies silently if not already installed
pip show streamlit >nul 2>&1
if %errorlevel% neq 0 (
    echo  Installing Streamlit - this may take a minute...
    pip install streamlit --quiet
)

pip show python-docx >nul 2>&1
if %errorlevel% neq 0 (
    echo  Installing python-docx...
    pip install python-docx --quiet
)

echo  All good! Launching DocuStyle...
echo  (A browser window will open automatically)
echo  (To close the app, close this window)
echo.

:: Launch streamlit and open browser
start "" http://localhost:8501
streamlit run "%~dp0app.py" --server.headless true --browser.gatherUsageStats false

pause
