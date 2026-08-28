' map-video-drive-hidden.vbs
' Launches map-video-drive.ps1 with no visible window.
'
' Why a VBS wrapper at all: the Remote-HDD project hides its own every-2-minutes
' mapper with an S4U task, but registering one requires an elevated shell (so
' does any at-logon trigger). This task is registered by an ordinary user, so it
' runs on an interactive logon - and powershell.exe flashes a console for a
' moment every cycle even with -WindowStyle Hidden. wscript.exe has no console
' of its own, so Run(..., 0, False) is truly silent without any elevation.
' Same pattern as the DaVinci dashboard's start_dashboard_hidden.vbs.

Dim shell, fso, scriptDir, ps1
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ps1 = fso.BuildPath(scriptDir, "map-video-drive.ps1")

shell.Run "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & ps1 & """", 0, False
