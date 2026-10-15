[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('export', 'import')]
    [string]$Mode
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
$logDirectory = Join-Path $projectRoot 'logs'
New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
$logPath = Join-Path $logDirectory "$Mode-scheduler.log"
Set-Location -LiteralPath $projectRoot
"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting $Mode" | Out-File -FilePath $logPath -Append -Encoding utf8
$previousErrorActionPreference = $ErrorActionPreference
try {
    # Native stderr can contain warnings; use Python's exit code for success.
    $ErrorActionPreference = 'Continue'
    if ($Mode -eq 'export') {
        & $python -u '.\hrmi.py' --auto-login --report-only --scheduled 2>&1 | Out-File -FilePath $logPath -Append -Encoding utf8
    }
    else {
        & $python -u '.\rocket75_import.py' 2>&1 | Out-File -FilePath $logPath -Append -Encoding utf8
    }
    $result = $LASTEXITCODE
}
finally {
    $ErrorActionPreference = $previousErrorActionPreference
}
"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Finished $Mode; exit=$result" | Out-File -FilePath $logPath -Append -Encoding utf8
exit $result
