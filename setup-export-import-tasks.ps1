[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$runner = Join-Path $projectRoot 'run-export-import-scheduled.ps1'
if (-not (Test-Path -LiteralPath (Join-Path $projectRoot '.venv\Scripts\python.exe'))) {
    throw 'Run install-new-machine.ps1 first.'
}
if ((Get-TimeZone).Id -ne 'SE Asia Standard Time') {
    throw 'Set the Windows time zone to Bangkok before installing these tasks.'
}
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$windowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($identity)
$runLevel = if ($windowsPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) { 'Highest' } else { 'Limited' }
$principal = New-ScheduledTaskPrincipal -UserId $identity.Name -LogonType Interactive -RunLevel $runLevel
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 6) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
foreach ($job in @(
    @{ Name = 'HRMI AUTO Export'; Mode = 'export'; At = '14:00' },
    @{ Name = 'HRMI AUTO Rocket75 Import'; Mode = 'import'; At = '14:30' },
    @{ Name = 'HRMI AUTO Leave Import'; Mode = 'leave'; At = '15:00' }
)) {
    if ($job.Mode -eq 'leave') {
        $leaveRunner = Join-Path $projectRoot 'run-rocket75-leave-scheduled.ps1'
        $taskArguments = "-NoProfile -ExecutionPolicy Bypass -File `"$leaveRunner`""
    }
    else {
        $taskArguments = "-NoProfile -ExecutionPolicy Bypass -File `"$runner`" -Mode $($job.Mode)"
    }
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $taskArguments -WorkingDirectory $projectRoot
    $timeOfDay = [TimeSpan]::Parse($job.At)
    $trigger = New-ScheduledTaskTrigger -Daily -At ((Get-Date).Date.Add($timeOfDay))
    Register-ScheduledTask -TaskName $job.Name -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Daily HRMI $($job.Mode) at $($job.At) Bangkok time" -Force | Out-Null
    Write-Host "Created $($job.Name) at $($job.At); privileges=$runLevel"
}
