# Removes xl: package, command shims, Claude Code skill and hook, settings.json entry.
# Leaves Python and its packages in place.
param([switch]$NoPause)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
& py -3.14 (Join-Path $here "setup.py") uninstall
$code = $LASTEXITCODE
if (-not $NoPause -and -not $env:CLAUDECODE -and -not $env:CI -and $Host.Name -eq "ConsoleHost") { Read-Host "Press Enter to close" | Out-Null }
exit $code
