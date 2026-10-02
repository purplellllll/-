Option Explicit

' Runs the scheduled sync without creating a console window.
Dim shell, folder, command, result
Set shell = CreateObject("WScript.Shell")
folder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & folder & "\run.ps1"" sync"
result = shell.Run(command, 0, True)
WScript.Quit result
