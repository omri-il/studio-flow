# register-map-task.ps1
# Run ONCE on the laptop (as your normal user, no admin needed).
# Creates a scheduled task that maps E: every 2 minutes and at logon.

$TaskName   = "Map Video Drive (E:)"
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

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger @($logon, $repeating) `
    -Settings $settings `
    -Force | Out-Null

Write-Host "Task registered: '$TaskName'"
Write-Host "It will run at logon and every 2 minutes."
Write-Host ""
Write-Host "NEXT STEP - store home PC credentials (run once in Command Prompt):"
Write-Host ""
Write-Host "  If home PC uses a Microsoft account (most Windows 11):"
Write-Host "    cmdkey /add:100.111.186.101 /user:MicrosoftAccount\your.email@example.com /pass:YOUR_PASSWORD"
Write-Host ""
Write-Host "  If home PC uses a local account:"
Write-Host "    cmdkey /add:100.111.186.101 /user:DESKTOP-7HQM8GO\omrii /pass:YOUR_PASSWORD"
