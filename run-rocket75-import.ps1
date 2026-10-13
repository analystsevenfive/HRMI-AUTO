[CmdletBinding()]
param(
    [string]$StartDate,
    [string]$EndDate,
    [ValidateSet("BB", "BK", "NTS", "PK", "SM")]
    [string[]]$Organizations = @("BB", "BK", "NTS", "PK", "SM"),
    [ValidateSet("absence", "late")]
    [string[]]$Types = @("absence", "late")
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $projectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    throw "Virtual environment not found: $python. Run setup-daily-task.ps1 first."
}
if (([bool]$StartDate) -xor ([bool]$EndDate)) {
    throw "Specify both StartDate and EndDate, or omit both to use the automatic range."
}

$arguments = @(".\rocket75_import.py", "--organizations") + $Organizations + @("--types") + $Types
if ($StartDate) {
    $arguments += @("--start-date", $StartDate, "--end-date", $EndDate)
}
Set-Location -LiteralPath $projectRoot
& $python @arguments
exit $LASTEXITCODE
