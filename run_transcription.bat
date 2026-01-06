@echo off
REM Transcribe recordings in batches with automatic recovery
REM This script processes recordings in small batches and automatically restarts on errors

setlocal EnableDelayedExpansion

REM Add FFmpeg to PATH
set "FFMPEG_PATH=C:\Users\ASUS\AppData\Local\Microsoft\WinGet\Packages\Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe\ffmpeg-8.0.1-full_build\bin"
set "PATH=%FFMPEG_PATH%;%PATH%"

cd /d "C:\Users\ASUS\Desktop\Downloads\Downloads\New folder\Upwork\Issac - Copilot agent"

echo ============================================================
echo TRANSCRIPTION BATCH RUNNER
echo ============================================================
echo Started at: %date% %time%
echo FFmpeg path: %FFMPEG_PATH%
echo.

:loop
echo.
echo [%date% %time%] Processing next batch of 5 recordings...
echo.

.\.venv\Scripts\python.exe src/transcribe_one_by_one.py --count 5

if errorlevel 1 (
    echo.
    echo [%date% %time%] Script exited with error. Waiting 30 seconds before retry...
    timeout /t 30 /nobreak
)

REM Check if there are more recordings to process
.\.venv\Scripts\python.exe -c "import json; p=json.load(open('transcription_progress.json')); exit(0 if p.get('pending_recordings') else 1)"

if errorlevel 1 (
    echo.
    echo ============================================================
    echo ALL RECORDINGS PROCESSED!
    echo Completed at: %date% %time%
    echo ============================================================
    goto :end
)

REM Small delay between batches
echo.
echo [%date% %time%] Batch complete. Starting next batch in 10 seconds...
timeout /t 10 /nobreak
goto :loop

:end
echo.
echo Press any key to exit...
pause >nul
