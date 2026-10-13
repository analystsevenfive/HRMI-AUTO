[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$TaskName = "HRMI AUTO Daily",
    [datetime]$DailyAt = "06:00"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $projectRoot ".venv\Scripts\python.exe"
$requirements = Join-Path $projectRoot "requirements.txt"
$runner = Join-Path $projectRoot "run-hrmi-scheduled.ps1"
$navigation = Join-Path $projectRoot "hrmi-navigation.json"

if (-not (Test-Path -LiteralPath $navigation)) {
    throw "hrmi-navigation.json not found. Calibrate Time Attendance before scheduling."
}

if (-not (Test-Path -LiteralPath $python)) {
    & py -m venv (Join-Path $projectRoot ".venv")
    if ($LASTEXITCODE -ne 0) { throw "Failed to create the virtual environment." }
}

& $python -m pip install -r $requirements
if ($LASTEXITCODE -ne 0) { throw "Failed to install Python dependencies." }

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$runner`"" `
    -WorkingDirectory $projectRoot
$trigger = New-ScheduledTaskTrigger -Daily -At $DailyAt
$settings = New-ScheduledTaskSettingsSet `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 30) `
    -ExecutionTimeLimit (New-TimeSpan -Hours 6) `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable
$principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Highest

if ($PSCmdlet.ShouldProcess($TaskName, "Create daily Scheduled Task at $($DailyAt.ToString('HH:mm'))")) {
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -Principal $principal `
        -Description "Process HRMI for all organizations using the date two days ago" `
        -Force | Out-Null
    Write-Host "Scheduled Task '$TaskName' created for $($DailyAt.ToString('HH:mm'))."
}
