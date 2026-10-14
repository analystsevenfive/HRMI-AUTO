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
if ($Mode -eq 'export') {
    & $python -u '.\hrmi.py' --auto-login --report-only --scheduled *>> $logPath
}
else {
    & $python -u '.\rocket75_import.py' *>> $logPath
}
$result = $LASTEXITCODE
"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Finished $Mode; exit=$result" | Out-File -FilePath $logPath -Append -Encoding utf8
exit $result
