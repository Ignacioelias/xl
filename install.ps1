# xl installer. Right-click > "Run with PowerShell", or from a terminal:
#   powershell -ExecutionPolicy Bypass -File install.ps1 [-NoPause]
# Steps: (0) make sure Python 3.14 is present (installs it with winget if not),
#        (1-7) setup.py installs the Python packages, puts the xl package in %USERPROFILE%\tools\xl,
#        copies the two Claude Code skills and the hook, registers the hook, adds the `xl` command
#        and runs six checks with Excel.
param([switch]$NoPause)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$pause = -not $NoPause -and -not $env:CLAUDECODE -and -not $env:CI -and $Host.Name -eq "ConsoleHost"

function Have-Py314 {
    try { $v = & py -3.14 -c "import sys;print(sys.version_info[:2])" 2>$null; return ($LASTEXITCODE -eq 0 -and $v -match "3, 14") }
    catch { return $false }
}

Write-Host ""
Write-Host "xl installer" -ForegroundColor Cyan
Write-Host "Source: $here"
Write-Host ""

if (-not (Have-Py314)) {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "[0/7] Python 3.14 not found, and winget is not available to install it." -ForegroundColor Red
        Write-Host "Install Python 3.14 from https://www.python.org/downloads/windows/ (tick 'py launcher'), then run install.ps1 again."
        if ($pause) { Read-Host "Press Enter to close" | Out-Null }
        exit 1
    }
    Write-Host "[0/7] Python 3.14 not found. Installing with winget (2-3 minutes)..." -ForegroundColor Yellow
    winget install --id Python.Python.3.14 --scope user --silent --accept-package-agreements --accept-source-agreements
    # the py launcher lands in %LOCALAPPDATA%\Programs\Python\Launcher; make it visible to this process
    $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [Environment]::GetEnvironmentVariable("Path", "User")
    if (-not (Have-Py314)) {
        Write-Host "Python 3.14 still not visible in this terminal. Close it, open a new one, and run install.ps1 again." -ForegroundColor Red
        Write-Host "If winget is blocked on your laptop, install Python 3.14 from https://www.python.org/downloads/windows/ instead."
        if ($pause) { Read-Host "Press Enter to close" | Out-Null }
        exit 1
    }
} else {
    Write-Host "[0/7] Python 3.14 found."
}

& py -3.14 (Join-Path $here "setup.py") install
$code = $LASTEXITCODE
Write-Host ""
if ($code -eq 0) { Write-Host "xl is installed." -ForegroundColor Green }
elseif ($code -eq 2) { Write-Host "Installed, but a check failed (see above)." -ForegroundColor Yellow }
else { Write-Host "Install failed (see above)." -ForegroundColor Red }
if ($pause) { Read-Host "Press Enter to close" | Out-Null }
exit $code
