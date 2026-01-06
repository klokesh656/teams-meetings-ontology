@echo off
REM ============================================================================
REM AUTO OUTLIER ANALYZER - Daily Run Script
REM ============================================================================
REM This script replaces weekly Outlier meetings by automatically analyzing
REM check-in meetings and generating AI suggestions.
REM 
REM Run this daily after transcript sync to get automated outlier-level insights.
REM ============================================================================

echo.
echo ============================================================================
echo   AUTO OUTLIER ANALYZER - Daily Run
echo   Analyzing Check-in Meetings for Outlier-Level Insights
echo ============================================================================
echo.

cd /d "%~dp0"

REM Activate virtual environment and run analyzer
call .venv\Scripts\activate.bat

REM Process transcripts from last 3 days (catches weekend backlog)
python src/auto_outlier_analyzer.py --days 3

echo.
echo ============================================================================
echo   Analysis Complete!
echo   
echo   Next Steps:
echo   1. Review reports in output\ folder
echo   2. Review pending suggestions: python src/outlier_insights_engine.py pending
echo   3. Approve/reject suggestions with your name
echo ============================================================================
echo.

pause
