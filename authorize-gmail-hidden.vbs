Option Explicit

' Starts the one-time Gmail consent flow without leaving a PowerShell window open.
' The default browser still opens so the account owner can approve the request.
Dim shell, folder, command, result
Set shell = CreateObject("WScript.Shell")
folder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & folder & "\run.ps1"" authorize-gmail"
result = shell.Run(command, 0, True)
WScript.Quit result
