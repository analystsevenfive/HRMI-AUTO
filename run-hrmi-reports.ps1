[CmdletBinding()]
param(
    [string]$StartDate,
    [string]$EndDate,
    [ValidateSet("[BB]", "[BK]", "[NTS]", "[PK]", "[SM]")]
    [string[]]$Organizations = @("[BB]", "[BK]", "[NTS]", "[PK]", "[SM]"),
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $projectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $python)) {
    throw "Virtual environment not found: $python. Run setup-daily-task.ps1 first."
}
if ([bool]$StartDate -xor [bool]$EndDate) {
    throw "Specify both StartDate and EndDate, or omit both to use the automatic range."
}
if (-not $StartDate) {
    $today = Get-Date
    if ($today.DayOfWeek -eq [System.DayOfWeek]::Tuesday) {
        $StartDate = $today.AddDays(-3).ToString("yyyy-MM-dd")
        $EndDate = $today.AddDays(-2).ToString("yyyy-MM-dd")
    }
    else {
        $StartDate = $today.AddDays(-2).ToString("yyyy-MM-dd")
        $EndDate = $StartDate
    }
    Write-Host "Automatic report range: $StartDate to $EndDate"
}

Set-Location -LiteralPath $projectRoot
$arguments = @(
    ".\hrmi.py",
    "--auto-login",
    "--report-only",
    "--start-date", $StartDate,
    "--end-date", $EndDate,
    "--organizations"
) + $Organizations
if ($Force) { $arguments += "--force" }

$previousErrorActionPreference = $ErrorActionPreference
try {
    # Python warnings on stderr must not terminate Windows PowerShell.
    $ErrorActionPreference = 'Continue'
    & $python @arguments
    $result = $LASTEXITCODE
}
finally {
    $ErrorActionPreference = $previousErrorActionPreference
}
exit $result
