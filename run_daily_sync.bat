@echo off
REM Daily Sync Script - Run this to sync transcripts and analyze with AI
REM Can be scheduled with Windows Task Scheduler

cd /d "c:\Users\ASUS\Desktop\Downloads\Downloads\New folder\Upwork\Issac - Copilot agent"

REM Activate virtual environment if exists
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

REM Run the daily sync
python src/daily_sync.py

REM Pause only if run manually (not from scheduler)
if "%1"=="" pause
