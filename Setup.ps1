<#
.SYNOPSIS
Initial setup script for Auto Theme.
.DESCRIPTION
Sets up the Auto Theme script and scheduled task if not already configured. Automatically requests admin privileges if not run as admin.
#>

# Function to check if the script is running as admin
function Test-AdminRights {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Relaunch script as admin if not already running as admin
if (-not (Test-AdminRights)) {
    Write-Host "This script requires administrative privileges. Requesting elevation..." -ForegroundColor Yellow
    Start-Process -FilePath "powershell.exe" `
                  -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
                  -Verb RunAs
    Exit 0
}

# Define the path to the AutoTheme.ps1 script and task XML file
$ScriptFile = Join-Path -Path $PSScriptRoot -ChildPath "AutoTheme.ps1"
$TaskFile = Join-Path -Path $PSScriptRoot -ChildPath "AutoTheme.xml"
$TaskName = "AutoTheme2"

# Check if AutoTheme.ps1 exists
if (!(Test-Path $ScriptFile)) {
    Write-Host "Error: Required script file '$ScriptFile' not found. Exiting setup..." -ForegroundColor Red
    Pause
    Exit 1
}

# Check if the scheduled task already exists
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Write-Host "Task '$TaskName' already exists. Exiting setup..." -ForegroundColor Yellow
    Pause
    Exit 0
}

# Check if the task XML exists
if (!(Test-Path $TaskFile)) {
    Write-Host "Error: Task configuration file '$TaskFile' not found. Exiting setup..." -ForegroundColor Red
    Pause
    Exit 1
}

# Modify the XML to update the path to the AutoTheme.ps1 script dynamically
$taskXml = Get-Content -Path $TaskFile -Raw
$taskXml = $taskXml -replace "AutoTheme.ps1", $ScriptFile

# Register the scheduled task using Register-ScheduledTask
Write-Host "Importing scheduled task '$TaskName' from XML..." -ForegroundColor Green
try {
    Register-ScheduledTask -Xml $taskXml -TaskName $TaskName -User "SYSTEM" -Force | Out-Null
    Write-Host "Scheduled task '$TaskName' created successfully!" -ForegroundColor Cyan
} catch {
    Write-Host "Failed to import scheduled task: $_" -ForegroundColor Red
    Pause
    Exit 1
}

Write-Host "First-run setup completed successfully."
Pause
