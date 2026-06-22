@echo off
REM Fund & Portfolio Tool - one-time installer (Windows)
setlocal
cd /d "%~dp0"

echo(
echo === Fund ^& Portfolio Tool installer ===
echo(

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found on PATH.
    echo Install Python 3.10+ from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during setup, then re-run this file.
    pause
    exit /b 1
)

echo Creating local virtual environment (.venv)...
python -m venv .venv
if errorlevel 1 ( echo [ERROR] Could not create the virtual environment. & pause & exit /b 1 )

call .venv\Scripts\activate.bat

echo Upgrading pip...
python -m pip install --upgrade pip

echo Installing dependencies (pandas, openpyxl, xlwings)...
pip install -r requirements.txt
if errorlevel 1 ( echo [ERROR] Dependency install failed. & pause & exit /b 1 )

echo Installing the xlwings Excel add-in...
xlwings addin install

echo(
echo === Done ===
echo Next:
echo   1) (optional) double-click run_headless.bat to verify the engine.
echo   2) Open PortfolioTool.xlsx and follow SETUP.md to add the buttons,
echo      then Save As PortfolioTool.xlsm.
echo(
pause
