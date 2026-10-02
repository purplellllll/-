Option Explicit

' Starts the persistent Feishu callback listener without a console window.
Dim shell, folder, command, result
Set shell = CreateObject("WScript.Shell")
folder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & folder & "\run-interview-listener.ps1"""
result = shell.Run(command, 0, True)
WScript.Quit result
