[CmdletBinding()]
param(
    [string]$LeaveRoot = 'T:\HR_Leave_Approved'
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
$logDirectory = Join-Path $projectRoot 'logs'
$logPath = Join-Path $logDirectory 'leave-import-scheduler.log'
New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
Set-Location -LiteralPath $projectRoot
$result = 1
try {
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting leave import; root=$LeaveRoot" | Out-File -FilePath $logPath -Append -Encoding utf8
    if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
        throw 'Virtual environment not found. Run install-new-machine.ps1 first.'
    }
    if (-not (Test-Path -LiteralPath $LeaveRoot -PathType Container)) {
        throw "Leave root is unavailable: $LeaveRoot"
    }
    & $python -u (Join-Path $projectRoot 'rocket75_import.py') --types leave --leave-root $LeaveRoot *>> $logPath
    $result = $LASTEXITCODE
}
catch {
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $($_.Exception.Message)" | Out-File -FilePath $logPath -Append -Encoding utf8
    Write-Error $_ -ErrorAction Continue
}
finally {
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Finished leave import; exit=$result" | Out-File -FilePath $logPath -Append -Encoding utf8
}
exit $result
