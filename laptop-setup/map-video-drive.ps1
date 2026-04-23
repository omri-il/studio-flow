# map-video-drive.ps1
# Run automatically by scheduled task — maps home PC E: drive as E: on this laptop.
# Log: C:\Users\omrii\Scripts\map-video-drive.log

$HomePcIp   = "100.111.186.101"
$SharePath  = "\\$HomePcIp\e"
$DriveLetter = "E"
$LogFile    = "$PSScriptRoot\map-video-drive.log"
$MaxLogLines = 50

function Write-Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Add-Content -Path $LogFile -Value $line
}

# Keep log trim
if (Test-Path $LogFile) {
    $lines = Get-Content $LogFile
    if ($lines.Count -gt $MaxLogLines) {
        $lines | Select-Object -Last $MaxLogLines | Set-Content $LogFile
    }
}

# Check if E: is already mapped to our share
$existing = Get-PSDrive -Name $DriveLetter -ErrorAction SilentlyContinue
if ($existing) {
    if ($existing.DisplayRoot -eq $SharePath) {
        # Already mapped correctly - exit silently
        exit 0
    } else {
        # E: exists but points somewhere else (e.g. local USB drive)
        Write-Log "WARNING: ${DriveLetter}: exists but points to '$($existing.DisplayRoot)' - not overwriting."
        exit 0
    }
}

# Ping home PC (1 attempt, 1-second timeout)
$ping = Test-Connection -ComputerName $HomePcIp -Count 1 -Quiet -ErrorAction SilentlyContinue
if (-not $ping) {
    Write-Log "Home PC ($HomePcIp) unreachable - skipping."
    exit 0
}

# Map the drive
try {
    $result = net use "${DriveLetter}:" $SharePath /persistent:no 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Log "Mapped ${DriveLetter}: -> $SharePath"
    } else {
        Write-Log "ERROR mapping ${DriveLetter}: -> $SharePath : $result"
    }
} catch {
    Write-Log "EXCEPTION: $_"
}
