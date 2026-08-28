' map-video-drive-hidden.vbs
' Launches map-video-drive.ps1 with no visible window.
'
' Why a VBS wrapper at all: the scheduled task has to run under an ordinary
' interactive logon, because (a) registering an S4U task needs an elevated
' shell, and (b) an S4U token carries no network credentials and cannot unlock
' the user's DPAPI-protected Credential Manager - so `net use` would not see the
' stored netshare credential, which is the whole point of the task.
' An interactive task keeps the credential, but powershell.exe still flashes a
' console for a moment every 2 minutes even with -WindowStyle Hidden.
' wscript.exe has no console of its own, so Run(..., 0, False) is truly silent.
' Same pattern as the DaVinci dashboard's start_dashboard_hidden.vbs.

Dim shell, fso, scriptDir, ps1
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ps1 = fso.BuildPath(scriptDir, "map-video-drive.ps1")

shell.Run "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & ps1 & """", 0, False
