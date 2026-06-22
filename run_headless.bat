@echo off
REM Verify the engine end-to-end WITHOUT Excel (uses the sample templates).
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] Not installed yet. Double-click install.bat first.
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat
python test_pipeline.py
echo(
echo If you see "ALL CHECKS PASSED" above, the engine works.
echo Outputs: sample_results.xlsx and sample_evolution.xlsx (open in Excel).
pause
