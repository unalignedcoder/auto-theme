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

Write-Host "===== Task Setup script started ====="

# Relaunch script as admin if not already running as admin
if (-not (Test-AdminRights)) {
    Write-Host "This script requires administrative privileges. Requesting elevation..." -ForegroundColor Yellow
    Start-Process -FilePath "powershell.exe" `
                  -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
                  -Verb RunAs
    Pause
    Exit 0
}

# Define the path to the AutoTheme.ps1 script and task XML file
$AutoThemeScript = Join-Path -Path $PSScriptRoot -ChildPath "AutoTheme.ps1"
$TaskName = "Auto Theme"

# Check if AutoTheme.ps1 exists
if (!(Test-Path $AutoThemeScript)) {
    Write-Host "Error: Required script file '$AutoThemeScript' not found. Exiting setup..." -ForegroundColor Red
    Pause
    Exit 1
}

# Check if the scheduled task already exists
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Write-Host "Task '$TaskName' already exists." -ForegroundColor Yellow
    $runExistingTask = Read-Host "Would you like to run the existing task now? (Yes/No)"
    if ($runExistingTask -match '^(Yes|Y)$') {
        try {
            # Run the task using Task Scheduler
            Write-Host "Running the existing task via Task Scheduler..." -ForegroundColor Cyan
            Start-ScheduledTask -TaskName $TaskName
            Write-Host "Task '$TaskName' has been triggered successfully." -ForegroundColor Cyan
        } catch {
            Write-Host "Failed to run the task: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "You chose not to run the task. Exiting setup..." -ForegroundColor Yellow
    }
    Pause
    Exit 0
}
# Create the triggers
$LogonTrigger = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERNAME"
$startupTrigger = New-ScheduledTaskTrigger -AtStartup -User "$env:USERNAME"

<# Create the Unlock Trigger using CIM, Thanks to
https://stackoverflow.com/questions/53704188/syntax-for-execute-on-workstation-unlock #>
$StateChangeTrigger = Get-CimClass `
    -Namespace ROOT\Microsoft\Windows\TaskScheduler `
    -ClassName MSFT_TaskSessionStateChangeTrigger

$UnlockTrigger = New-CimInstance `
    -CimClass $StateChangeTrigger `
    -Property @{
        StateChange = 8  # 8 = TASK_SESSION_UNLOCK
		UserId = "$env:USERNAME"
    } `
    -ClientOnly

$Triggers = @($LogonTrigger, $startupTrigger, $UnlockTrigger)

# Create the action
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -NoProfile -File `"$AutoThemeScript`""

# Register the task
Register-ScheduledTask -TaskName $TaskName -Trigger $Triggers -User "$env:USERNAME" -Action $Action -RunLevel Highest -Compatibility Win8 -Force | Out-Null
Write-Host "Scheduled task '$TaskName' created successfully!" -ForegroundColor Cyan

# Prompt the user to run the task immediately
$runNow = Read-Host "Would you like to run the task now? (Yes/No)"

if ($runNow -match '^(Yes|Y)$') {
    try {
        # Run the task using Task Scheduler (this simulates the task running as scheduled)
        Write-Host "Running the task via Task Scheduler..." -ForegroundColor Cyan
        Start-ScheduledTask -TaskName $TaskName
        Write-Host "Task '$TaskName' has been triggered successfully." -ForegroundColor Cyan
    } catch {
        Write-Host "Failed to run the task: $_" -ForegroundColor Red
    }
} else {
    Write-Host "You chose not to run the task. Setup is complete." -ForegroundColor Yellow
}

