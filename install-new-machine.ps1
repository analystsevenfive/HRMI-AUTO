[CmdletBinding()]
param(
    [switch]$InstallDailyTask,
    [datetime]$DailyAt = "09:00",
    [string]$HrmiExe = "C:\Program Files (x86)\Prosoft\Prosoft HRMI\Bin\HRMI.exe"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
$envFile = Join-Path $projectRoot ".env"
$envExample = Join-Path $projectRoot ".env.example"
$navigation = Join-Path $projectRoot "hrmi-navigation.json"

function Invoke-SystemPython {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3 @Arguments
    }
    elseif (Get-Command python -ErrorAction SilentlyContinue) {
        & python @Arguments
    }
    else {
        throw "Python 3.10 or newer was not found. Install Python and enable 'Add python.exe to PATH'."
    }
    if ($LASTEXITCODE -ne 0) { throw "Python command failed: $($Arguments -join ' ')" }
}

Write-Host "[1/6] Checking required software..."
Invoke-SystemPython -Arguments @("-c", "import sys; assert sys.version_info >= (3, 10), 'Python 3.10 or newer is required'")
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    throw "Node.js was not found. Install the current Node.js LTS release and run this script again."
}
if (-not (Test-Path -LiteralPath $HrmiExe -PathType Leaf)) {
    throw "Prosoft HRMI was not found at: $HrmiExe"
}

Write-Host "[2/6] Creating Python virtual environment..."
if (-not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
    Invoke-SystemPython -Arguments @("-m", "venv", (Join-Path $projectRoot ".venv"))
}

Write-Host "[3/6] Installing Python dependencies..."
& $venvPython -m pip install --disable-pip-version-check -r (Join-Path $projectRoot "requirements.txt")
if ($LASTEXITCODE -ne 0) { throw "Failed to install Python dependencies." }

Write-Host "[4/6] Preparing local configuration..."
if (-not (Test-Path -LiteralPath $envFile)) {
    Copy-Item -LiteralPath $envExample -Destination $envFile
    Write-Host "Created .env from .env.example. Add the four account values before running automation." -ForegroundColor Yellow
}
else {
    Write-Host ".env already exists; it was left unchanged."
}

Write-Host "[5/6] Running installation checks..."
& $venvPython -c "import psutil, requests, win32api; from pywinauto import Application; print('Python dependencies: OK')"
if ($LASTEXITCODE -ne 0) { throw "Python dependency check failed." }
& node --version
if ($LASTEXITCODE -ne 0) { throw "Node.js check failed." }

Write-Host "[6/6] Scheduled Task..."
if ($InstallDailyTask) {
    if (Test-Path -LiteralPath $navigation) {
        & (Join-Path $projectRoot "setup-daily-task.ps1") -DailyAt $DailyAt
    }
    else {
        Write-Host "Scheduled Task was not created because this computer has not been calibrated." -ForegroundColor Yellow
    }
}
else {
    Write-Host "Skipped. Add -InstallDailyTask after calibration when daily automation is needed."
}

Write-Host ""
Write-Host "Installation files are ready." -ForegroundColor Green
if (-not (Test-Path -LiteralPath $navigation)) {
    Write-Host "Next:" -ForegroundColor Cyan
    Write-Host "  1. Fill in .env"
    Write-Host "  2. Calibrate this computer:"
    Write-Host "     .\.venv\Scripts\python.exe .\hrmi.py --auto-login --stop-before-next --calibrate-time-attendance"
    Write-Host "  3. Install the 09:00 task:"
    Write-Host "     .\install-new-machine.ps1 -InstallDailyTask"
}
elseif (-not $InstallDailyTask) {
    Write-Host "To install the 09:00 task: .\install-new-machine.ps1 -InstallDailyTask"
}
