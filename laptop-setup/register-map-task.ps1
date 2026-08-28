# register-map-task.ps1
# Run ONCE on the laptop (as your normal user, no admin needed).
# Creates a scheduled task that maps E: every 2 minutes and at logon.

$TaskName   = "MapVideoDrive"
$ScriptPath = "$PSScriptRoot\map-video-drive.ps1"

if (-not (Test-Path $ScriptPath)) {
    Write-Error "map-video-drive.ps1 not found at: $ScriptPath"
    exit 1
}

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`""

$logon = New-ScheduledTaskTrigger -AtLogOn

$repeating = New-ScheduledTaskTrigger -Once -At (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Minutes 2) `
    -RepetitionDuration (New-TimeSpan -Days 3650)

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable

# S4U principal: run under Omri's account with no interactive session, so
# PowerShell does not flash a console window every 2 minutes. (Lesson from the
# Remote-HDD project - the default "run only when the user is logged on"
# principal pops a visible window despite -WindowStyle Hidden.)
$principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType S4U `
    -RunLevel Limited

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger @($logon, $repeating) `
    -Settings $settings `
    -Principal $principal `
    -Force | Out-Null

Write-Host "Task registered: 'MapVideoDrive'"
Write-Host "It will run at logon and every 2 minutes."
Write-Host ""
Write-Host "NEXT STEP - store home PC credentials (run once in Command Prompt):"
Write-Host ""
Write-Host "    cmdkey /add:100.111.186.101 /user:DESKTOP-7HQM8GO\netshare /pass:YOUR_PASSWORD"
Write-Host ""
Write-Host "'netshare' is the limited account the home PC's Remote-HDD setup created"
Write-Host "for the DriveE / DriveD shares; its password lives in that project's notes."
