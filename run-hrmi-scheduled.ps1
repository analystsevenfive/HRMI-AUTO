[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $projectRoot ".venv\Scripts\python.exe"
$logDirectory = Join-Path $projectRoot "logs"
$schedulerLog = Join-Path $logDirectory "scheduler.log"

if (-not (Test-Path -LiteralPath $python)) {
    throw "Virtual environment not found: $python. Run setup-daily-task.ps1 first."
}

New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
Set-Location -LiteralPath $projectRoot

"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting HRMI" | Out-File -FilePath $schedulerLog -Append
# Windows PowerShell treats native stderr (including Python warnings) as errors.
# Let Python finish and use its exit code to determine success.
$previousErrorActionPreference = $ErrorActionPreference
try {
    $ErrorActionPreference = 'Continue'
    & $python -u ".\hrmi.py" --scheduled *>> $schedulerLog
    $result = $LASTEXITCODE
}
finally {
    $ErrorActionPreference = $previousErrorActionPreference
}
"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Finished HRMI; exit=$result" | Out-File -FilePath $schedulerLog -Append
exit $result
