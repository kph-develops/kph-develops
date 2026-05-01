# =============================================================================
# HP Printer Usage Collector - Windows Task Scheduler Setup
# =============================================================================
# Compatible with ALL Windows versions (uses schtasks.exe only).
#
# Run ONCE from an elevated PowerShell prompt:
#   cd C:\path\to\hp_printer_collector
#   .\schedule_task.ps1
#
# To remove the task later:
#   schtasks /delete /tn "HP Printer Monthly Report" /f
# =============================================================================

#Requires -RunAsAdministrator

# Use Continue so a non-zero exit from schtasks.exe does not throw;
# we check $LASTEXITCODE ourselves where it matters.
$ErrorActionPreference = "Continue"

$TaskName = "HP Printer Monthly Report"

# ---------------------------------------------------------------------------
# Locate Python
# ---------------------------------------------------------------------------

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$VenvPython = Join-Path $ProjectDir "venv\Scripts\python.exe"

$SystemPythonCmd = Get-Command python -ErrorAction SilentlyContinue
if ($SystemPythonCmd) {
    $SystemPython = $SystemPythonCmd.Source
} else {
    $SystemPython = $null
}

if (Test-Path $VenvPython) {
    $PythonExe = $VenvPython
    Write-Host "Using venv Python : $PythonExe" -ForegroundColor Cyan
} elseif ($SystemPython) {
    $PythonExe = $SystemPython
    Write-Warning "venv not found. Falling back to system Python: $PythonExe"
    Write-Warning "Recommended: create a venv first with:"
    Write-Warning "  python -m venv venv"
    Write-Warning "  venv\Scripts\pip install -r requirements.txt"
} else {
    Write-Error "Python not found. Install Python 3 and re-run this script."
    exit 1
}

$MainScript = Join-Path $ProjectDir "main.py"
if (-not (Test-Path $MainScript)) {
    Write-Error "main.py not found at '$MainScript'. Run this script from the project root."
    exit 1
}

# ---------------------------------------------------------------------------
# Create a .bat launcher (schtasks.exe has no working-directory flag)
# ---------------------------------------------------------------------------

$LauncherPath = Join-Path $ProjectDir "run_report.bat"
$LauncherContent = "@echo off`r`ncd /d `"$ProjectDir`"`r`n`"$PythonExe`" `"$MainScript`"`r`n"
Set-Content -Path $LauncherPath -Value $LauncherContent -Encoding ASCII
Write-Host "Created launcher  : $LauncherPath" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Remove existing task silently (ignore error when task does not exist)
# ---------------------------------------------------------------------------

Write-Host "Removing existing task (if any)..." -ForegroundColor Yellow
schtasks /delete /tn $TaskName /f 2>$null | Out-Null
# Reset exit code - a non-zero here just means the task didn't exist, which is fine
$global:LASTEXITCODE = 0

# ---------------------------------------------------------------------------
# Register the task via schtasks.exe
# Use an array so quoting and spaces are handled correctly
# ---------------------------------------------------------------------------

$schtasksArgs = @(
    "/create",
    "/tn",  $TaskName,
    "/tr",  "`"$LauncherPath`"",
    "/sc",  "MONTHLY",
    "/d",   "1",
    "/st",  "09:00",
    "/rl",  "HIGHEST",
    "/f"
)

Write-Host "Registering scheduled task..." -ForegroundColor Cyan
& schtasks.exe @schtasksArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Error "schtasks.exe failed (exit code $LASTEXITCODE). See the error above."
    exit 1
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Green
Write-Host " Task registered successfully!"                        -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
Write-Host "  Task name   : $TaskName"
Write-Host "  Launcher    : $LauncherPath"
Write-Host "  Python      : $PythonExe"
Write-Host "  Script      : $MainScript"
Write-Host "  Working dir : $ProjectDir"
Write-Host "  Schedule    : 09:00 AM on the 1st of every month"
Write-Host ""
Write-Host "Verify in Task Scheduler GUI:" -ForegroundColor Cyan
Write-Host "  taskschd.msc"
Write-Host ""
Write-Host "Run immediately for a quick test:" -ForegroundColor Cyan
Write-Host "  schtasks /run /tn `"$TaskName`""
Write-Host ""
Write-Host "Remove the task:" -ForegroundColor Cyan
Write-Host "  schtasks /delete /tn `"$TaskName`" /f"
Write-Host ""
