[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$TaskName = "HRMI AUTO Daily",
    [datetime]$DailyAt = "10:00"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $projectRoot ".venv\Scripts\python.exe"
$requirements = Join-Path $projectRoot "requirements.txt"
$runner = Join-Path $projectRoot "run-hrmi-scheduled.ps1"
$navigation = Join-Path $projectRoot "hrmi-navigation.json"

if (-not (Test-Path -LiteralPath $python)) {
    throw "Virtual environment not found. Run install-new-machine.ps1 first."
}

$previousErrorActionPreference = $ErrorActionPreference
try {
    # Python warnings on stderr must not terminate Windows PowerShell.
    $ErrorActionPreference = 'Continue'
    & $python -m pip install -r $requirements
}
finally {
    $ErrorActionPreference = $previousErrorActionPreference
}
if ($LASTEXITCODE -ne 0) { throw "Failed to install Python dependencies." }

if (-not (Test-Path -LiteralPath $navigation)) {
    Write-Host "No saved coordinates; HRMI will use UI Automation to find Time Attendance."
}

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$runner`"" `
    -WorkingDirectory $projectRoot
$trigger = New-ScheduledTaskTrigger -Daily -At ((Get-Date).Date.Add($DailyAt.TimeOfDay))
$settings = New-ScheduledTaskSettingsSet `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 30) `
    -ExecutionTimeLimit (New-TimeSpan -Hours 6) `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$windowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($identity)
$runLevel = if ($windowsPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) { 'Highest' } else { 'Limited' }
$principal = New-ScheduledTaskPrincipal `
    -UserId $identity.Name `
    -LogonType Interactive `
    -RunLevel $runLevel

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
