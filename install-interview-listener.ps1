$ErrorActionPreference = 'Stop'
$TaskName = 'Recruitment Feishu Interview Listener'
$HiddenRunner = Join-Path $PSScriptRoot 'run-interview-listener-hidden.vbs'
if (-not (Test-Path -LiteralPath $HiddenRunner)) {
    throw "Missing hidden listener runner: $HiddenRunner"
}

try {
    $Action = New-ScheduledTaskAction -Execute "$env:WINDIR\System32\wscript.exe" -Argument "`"$HiddenRunner`""
    $Trigger = New-ScheduledTaskTrigger -AtLogOn
    $Principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
    $Settings = New-ScheduledTaskSettingsSet -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit (New-TimeSpan -Days 3650)
    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings -Description 'Receives Feishu interview scheduling replies through a long connection.' -Force | Out-Null
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "Interview listener installed as a scheduled task: $TaskName"
} catch {
    # Some Windows installations prohibit registering tasks for standard users.
    # A per-user Run entry keeps the listener hidden and starts it after logon.
    $RunKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
    $RunValue = "`"$env:WINDIR\System32\wscript.exe`" `"$HiddenRunner`""
    if (-not (Test-Path -LiteralPath $RunKey)) {
        New-Item -Path $RunKey -Force | Out-Null
    }
    Set-ItemProperty -Path $RunKey -Name $TaskName -Value $RunValue -ErrorAction Stop
    Write-Host "Scheduled-task access was denied. Installed hidden per-user startup entry: $TaskName"
}
