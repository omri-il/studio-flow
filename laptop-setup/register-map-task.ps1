# register-map-task.ps1
# Run ONCE on the laptop. No admin needed - see the elevation note below.
# Creates a scheduled task that maps E: every 2 minutes, plus a Startup-folder
# shortcut so E: is there immediately after logon.

$TaskName   = "MapVideoDrive"
$ScriptPath = "$PSScriptRoot\map-video-drive.ps1"
$VbsPath    = "$PSScriptRoot\map-video-drive-hidden.vbs"

foreach ($required in @($ScriptPath, $VbsPath)) {
    if (-not (Test-Path $required)) {
        Write-Error "Not found: $required"
        exit 1
    }
}

# Launch through wscript, not powershell.exe directly: powershell flashes a
# console for a moment every 2 minutes even with -WindowStyle Hidden, while
# wscript has no console of its own. See map-video-drive-hidden.vbs for why the
# task deliberately keeps an ordinary interactive logon instead of S4U.
$TaskCommand = "wscript.exe //B //Nologo `"$VbsPath`""

$isAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($isAdmin) {
    # An AtLogOn trigger is only registrable from an elevated shell, so it is
    # a bonus, not a requirement - the Startup shortcut below covers logon for
    # the ordinary non-elevated case.
    $action = New-ScheduledTaskAction -Execute "wscript.exe" `
        -Argument "//B //Nologo `"$VbsPath`""

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

    Write-Host "Task '$TaskName' registered (at logon + every 2 minutes)."
} else {
    # schtasks accepts a MINUTE schedule from a standard user; only ONLOGON
    # demands elevation.
    schtasks /create /tn $TaskName /tr $TaskCommand /sc minute /mo 2 /f | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Could not register '$TaskName' (schtasks exit $LASTEXITCODE)."
        exit 1
    }
    Write-Host "Task '$TaskName' registered (every 2 minutes)."
}

# Startup shortcut - runs the same VBS at logon without needing elevation.
$startup  = [Environment]::GetFolderPath('Startup')
$lnkPath  = Join-Path $startup "MapVideoDrive.lnk"
$shell    = New-Object -ComObject WScript.Shell
$lnk      = $shell.CreateShortcut($lnkPath)
$lnk.TargetPath       = "wscript.exe"
$lnk.Arguments        = "//B //Nologo `"$VbsPath`""
$lnk.WorkingDirectory = $PSScriptRoot
$lnk.WindowStyle      = 7          # minimized; wscript shows nothing anyway
$lnk.Description      = "Map the home PC's E: drive as E: on this laptop"
$lnk.Save()
Write-Host "Startup shortcut written: $lnkPath"

Write-Host ""
Write-Host "NEXT STEP - store home PC credentials (run once in Command Prompt):"
Write-Host ""
Write-Host "    cmdkey /add:100.111.186.101 /user:DESKTOP-7HQM8GO\netshare /pass:YOUR_PASSWORD"
Write-Host ""
Write-Host "'netshare' is the limited account the home PC's Remote-HDD setup created"
Write-Host "for the DriveE / DriveD shares; its password lives in that project's notes."
