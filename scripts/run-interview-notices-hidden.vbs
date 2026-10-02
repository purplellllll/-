Option Explicit

' Runs only the interview-link notification scan without opening a console.
Dim shell, folder, command, result
Set shell = CreateObject("WScript.Shell")
folder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & folder & "\run.ps1"" send-interview-notices"
result = shell.Run(command, 0, True)
WScript.Quit result
