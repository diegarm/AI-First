# Timesheet daily hour registration — PowerShell wrapper
# Usage:
#   .\run-timesheet.ps1              # register today
#   .\run-timesheet.ps1 -Date 20260508   # register specific date
#   .\run-timesheet.ps1 -Config      # reconfigure user/settings
#   .\run-timesheet.ps1 -DryRun      # preview without writing

param(
    [string]$Date = "",
    [switch]$Config,
    [switch]$DryRun
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonScript = Join-Path $ScriptDir "timesheet.py"

# Find python executable
$PythonCmd = $null
foreach ($candidate in @("python", "python3", "py")) {
    if (Get-Command $candidate -ErrorAction SilentlyContinue) {
        $PythonCmd = $candidate
        break
    }
}

if (-not $PythonCmd) {
    Write-Error "Python não encontrado. Instala Python 3 e tenta novamente."
    exit 1
}

$args_list = @()

if ($Config) {
    $args_list += "--config"
}
if ($DryRun) {
    $args_list += "--dry-run"
}
if ($Date) {
    $args_list += "--date"
    $args_list += $Date
}

& $PythonCmd $PythonScript @args_list
