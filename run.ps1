param(
    [Parameter(Position = 0)]
    [ValidateSet('authorize-gmail', 'check', 'sync', 'dispatch-interviews', 'send-interview-notices', 'queue-existing-interviews', 'status', 'self-test')]
    [string]$Command = 'sync'
)

$ErrorActionPreference = 'Stop'
$ProjectRoot = $PSScriptRoot
$PythonCandidates = @(
    (Join-Path $ProjectRoot '.venv\Scripts\python.exe'),
    $env:PYTHON_EXE,
    'python'
) | Where-Object { $_ }

$Python = $null
foreach ($Candidate in $PythonCandidates) {
    if ($Candidate -eq 'python') {
        $Found = Get-Command python -ErrorAction SilentlyContinue
        if ($Found) { $Python = $Found.Source; break }
    } elseif (Test-Path -LiteralPath $Candidate) {
        $Python = $Candidate
        break
    }
}

if (-not $Python) {
    throw 'Python 3.10+ was not found. Set PYTHON_EXE or install Python, then run again.'
}

& $Python (Join-Path $ProjectRoot 'recruitment_sync.py') $Command --config (Join-Path $ProjectRoot 'config.json')
exit $LASTEXITCODE
