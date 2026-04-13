# =============================================================================
# HP Printer Usage Collector – Windows Task Scheduler Setup
# =============================================================================
# Run this script ONCE from an elevated PowerShell prompt to register a
# monthly scheduled task that executes main.py at 09:00 on the 1st of
# every month.
#
# Usage (elevated PowerShell):
#   cd C:\path\to\hp_printer_collector
#   .\schedule_task.ps1
#
# To remove the task later:
#   Unregister-ScheduledTask -TaskName "HP Printer Monthly Report" -Confirm:$false
# =============================================================================

#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Configuration – edit these if your layout differs
# ---------------------------------------------------------------------------

$TaskName        = "HP Printer Monthly Report"
$TaskDescription = "Collect HP printer usage statistics and email a monthly report."

# Resolve the directory containing this script – works regardless of where
# you call it from, as long as you haven't dot-sourced it from another path.
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Prefer the virtual-environment Python; fall back to system Python.
$VenvPython  = Join-Path $ProjectDir "venv\Scripts\python.exe"
$SystemPython = (Get-Command python -ErrorAction SilentlyContinue)?.Source

if (Test-Path $VenvPython) {
    $PythonExe = $VenvPython
    Write-Host "Using venv Python : $PythonExe" -ForegroundColor Cyan
} elseif ($SystemPython) {
    $PythonExe = $SystemPython
    Write-Warning "venv not found at '$VenvPython'. Falling back to system Python: $PythonExe"
    Write-Warning "It is strongly recommended to create a venv first:"
    Write-Warning "  python -m venv venv && venv\Scripts\pip install -r requirements.txt"
} else {
    Write-Error "Python not found. Install Python 3 and re-run this script."
    exit 1
}

$MainScript  = Join-Path $ProjectDir "main.py"

if (-not (Test-Path $MainScript)) {
    Write-Error "main.py not found at '$MainScript'. Run this script from the project root."
    exit 1
}

# ---------------------------------------------------------------------------
# Build the scheduled task components
# ---------------------------------------------------------------------------

# Action: python main.py
$Action = New-ScheduledTaskAction `
    -Execute    $PythonExe `
    -Argument   "`"$MainScript`"" `
    -WorkingDirectory $ProjectDir

# Trigger: 09:00 on the 1st of every month
$Trigger = New-ScheduledTaskTrigger `
    -Monthly `
    -DaysOfMonth 1 `
    -At "09:00"

# Settings: run missed tasks, don't stop on AC/battery, allow on demand
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -MultipleInstances IgnoreNew

# Principal: run as the current user with highest available privileges
$Principal = New-ScheduledTaskPrincipal `
    -UserId    ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel  Highest

# ---------------------------------------------------------------------------
# Register (or update) the task
# ---------------------------------------------------------------------------

$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($existingTask) {
    Write-Host "Task '$TaskName' already exists – updating it..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

Register-ScheduledTask `
    -TaskName   $TaskName `
    -Description $TaskDescription `
    -Action     $Action `
    -Trigger    $Trigger `
    -Settings   $Settings `
    -Principal  $Principal | Out-Null

# ---------------------------------------------------------------------------
# Verify and summarise
# ---------------------------------------------------------------------------

$task = Get-ScheduledTask -TaskName $TaskName
$taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Green
Write-Host " Task registered successfully!" -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
Write-Host "  Name        : $($task.TaskName)"
Write-Host "  State       : $($task.State)"
Write-Host "  Python      : $PythonExe"
Write-Host "  Script      : $MainScript"
Write-Host "  Working dir : $ProjectDir"
Write-Host "  Schedule    : 09:00 on the 1st of every month"
Write-Host "  Next run    : $($taskInfo.NextRunTime)"
Write-Host ""
Write-Host "To run the task immediately (for testing):" -ForegroundColor Cyan
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host ""
Write-Host "To remove the task:" -ForegroundColor Cyan
Write-Host "  Unregister-ScheduledTask -TaskName '$TaskName' -Confirm:`$false"
Write-Host ""
