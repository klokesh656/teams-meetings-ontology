@echo off
REM ============================================
REM VA Check-in Pipeline - Daily Run
REM ============================================

echo Starting Daily VA Check-in Pipeline...
echo.

cd /d "%~dp0"

REM Activate virtual environment if it exists
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

REM Run the pipeline (default: 7 days lookback)
python src/daily_pipeline.py --days 7

echo.
echo Pipeline completed. Press any key to exit...
pause >nul
