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
if (-not $StartDate) { $StartDate = Read-Host "Start date (YYYY-MM-DD or DD/MM/YYYY)" }
if (-not $EndDate) { $EndDate = Read-Host "End date (YYYY-MM-DD or DD/MM/YYYY)" }

Set-Location -LiteralPath $projectRoot
$arguments = @(
    ".\hrmi.py",
    "--auto-login",
    "--start-date", $StartDate,
    "--end-date", $EndDate,
    "--organizations"
) + $Organizations
if ($Force) { $arguments += "--force" }

& $python @arguments
exit $LASTEXITCODE
