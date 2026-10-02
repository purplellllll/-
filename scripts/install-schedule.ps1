param(
    [ValidateRange(1, 60)]
    [int]$Minutes = 2
)

$ErrorActionPreference = 'Stop'
$TaskName = 'Recruitment Gmail to Feishu Sync'
$HiddenRunner = Join-Path $PSScriptRoot 'run-hidden.vbs'
if (-not (Test-Path -LiteralPath $HiddenRunner)) {
    throw "Missing hidden task runner: $HiddenRunner"
}
$Action = New-ScheduledTaskAction -Execute "$env:WINDIR\System32\wscript.exe" -Argument "`"$HiddenRunner`""
$Trigger = New-ScheduledTaskTrigger `
    -Once `
    -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes $Minutes) `
    -RepetitionDuration (New-TimeSpan -Days 3650)
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Description 'Reads Gmail label and writes candidate records to Feishu.' -Force | Out-Null
Write-Host "Scheduled task installed: $TaskName (every $Minutes minutes)."
