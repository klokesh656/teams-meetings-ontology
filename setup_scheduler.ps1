# Setup Windows Task Scheduler for Daily Sync
# Run this script as Administrator to create the scheduled task

$TaskName = "TeamsTranscriptDailySync"
$TaskPath = "\TranscriptAnalysis\"
$ScriptPath = "c:\Users\ASUS\Desktop\Downloads\Downloads\New folder\Upwork\Issac - Copilot agent\run_daily_sync.bat"
$WorkingDir = "c:\Users\ASUS\Desktop\Downloads\Downloads\New folder\Upwork\Issac - Copilot agent"

# Create the action
$Action = New-ScheduledTaskAction -Execute $ScriptPath -WorkingDirectory $WorkingDir -Argument "scheduled"

# Create the trigger - Run daily at 6:00 AM
$Trigger = New-ScheduledTaskTrigger -Daily -At 6:00AM

# Create settings
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries

# Register the task (runs as current user)
Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Action $Action -Trigger $Trigger -Settings $Settings -Description "Daily sync of Teams meeting transcripts and AI analysis"

Write-Host ""
Write-Host "✅ Scheduled task created successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Task Details:" -ForegroundColor Cyan
Write-Host "  Name: $TaskName"
Write-Host "  Schedule: Daily at 6:00 AM"
Write-Host "  Script: $ScriptPath"
Write-Host ""
Write-Host "To manage the task:" -ForegroundColor Yellow
Write-Host "  - Open Task Scheduler (taskschd.msc)"
Write-Host "  - Navigate to: Task Scheduler Library > TranscriptAnalysis"
Write-Host "  - Right-click to Run, Disable, or modify"
Write-Host ""
Write-Host "To run manually now:" -ForegroundColor Yellow
Write-Host "  schtasks /run /tn '\TranscriptAnalysis\TeamsTranscriptDailySync'"
