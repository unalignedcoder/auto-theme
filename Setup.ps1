<#
.SYNOPSIS
    Initial Setup script for Auto Theme.

.DESCRIPTION
    Sets up the Auto Theme script and scheduled task if not already configured. 
    Automatically requests admin privileges if not run as admin.
#>

# ============= Config file ==============

	$ConfigPath = Join-Path $PSScriptRoot "Config.ps1"

# ============= FUNCTIONS  ==============

    # Function to check if the script is running as admin
    function Test-AdminRights {
        $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    # Function to check if the OS is Windows 10 or Windows 11
    function Get-WindowsVersion {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        if ($os.Caption -match "Windows 10") {
            return "Windows 10"
        } elseif ($os.Caption -match "Windows 11") {
            return "Windows 11"
        } else {
            return "Other"
        }
    }

    # Create the logging system
    function LogThis {
	    param (
		    [string]$message,
		    [bool]$verboseMessage = $false   # Default to false if not specified
	    )

	    try {

		    # Only proceed if in debug mode
		    if ($log) {

			    <# Check for verbosity:
			    If the message is verbose, but verbose is false, end the Function
			    If the message is verbose, but verbose is true, continue
			    If the message is not verbose, continue #>
			    if ($verboseMessage -and -not $verbose) {

				    return  # Skip logging if message is verbose and $verbose is set to false
			    }
				    Add-Content -Path $logFile -Value "$message"  # Log to file
			    }
		    }

	    } catch {

		    Write-Output "Error in LogThis: $_"
	    }		
    }

    Write-Host "===== Task Setup script started ====="
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogThis ""
    LogThis "$timestamp === Setup started (Version: $scriptVersion)"

    # Relaunch script as admin if not already running as admin
    if (-not (Test-AdminRights)) {
        Write-Host "This script requires administrative privileges. Requesting elevation..." -ForegroundColor Yellow
        LogThis "This script requires administrative privileges. Requesting elevation."
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
        LogThis "Error: Required script file '$AutoThemeScript' not found. Exiting setup."
        Pause
        Exit 1
    }

    # Check if the scheduled task already exists
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Write-Host "Task '$TaskName' already exists." -ForegroundColor Yellow
        LogThis "Task '$TaskName' already exists."
        $runExistingTask = Read-Host "Would you like to run the existing task now? (Yes/No)"
        if ($runExistingTask -match '^(Yes|Y)$') {
            try {
                # Run the task using Task Scheduler
                Write-Host "Running the existing task via Task Scheduler..." -ForegroundColor Cyan
                LogThis "User requested to run the existing task via Task Scheduler."
                Start-ScheduledTask -TaskName $TaskName
                Write-Host "Task '$TaskName' has been triggered successfully." -ForegroundColor Cyan
                LogThis "Task '$TaskName' has been triggered successfully."
            } catch {
                Write-Host "Failed to run the task: $_" -ForegroundColor Red
                LogThis "Failed to run the task: $_"
            }
        } else {
            Write-Host "You chose not to run the task. Exiting setup..." -ForegroundColor Yellow
            LogThis "User chose not to run the task. Exiting setup."
        }
        Pause
        Exit 0
    }

    LogThis "Creating scheduled task '$TaskName'."

    # Create the triggers
    $LogonTrigger = New-ScheduledTaskTrigger -AtLogOn
    $startupTrigger = New-ScheduledTaskTrigger -AtStartup

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
    $windowsVersion = Get-WindowsVersion
    if ($windowsVersion -eq "Windows 10") {
        LogThis "Creating scheduled task for Windows 10."
        Register-ScheduledTask -TaskName $TaskName -Trigger $Triggers -User "$env:USERNAME" -Action $Action -RunLevel Highest -Compatibility Win8 -Force | Out-Null
    } else {
        LogThis "Creating scheduled task for Windows 11."
        Register-ScheduledTask -TaskName $TaskName -Trigger $Triggers -Action $Action -RunLevel Highest -Force | Out-Null
    }

    Write-Host "Scheduled task '$TaskName' created successfully!" -ForegroundColor Cyan
    LogThis "Scheduled task '$TaskName' created successfully."

    # Prompt the user to run the task immediately
    $runNow = Read-Host "Would you like to run the task now? (Yes/No)"

    if ($runNow -match '^(Yes|Y)$') {
        try {
            # Run the task using Task Scheduler (this simulates the task running as scheduled)
            Write-Host "Running the task via Task Scheduler..." -ForegroundColor Cyan
            LogThis "User requested to run the task via Task Scheduler."
            Start-ScheduledTask -TaskName $TaskName
            Write-Host "Task '$TaskName' has been triggered successfully." -ForegroundColor Cyan
            LogThis "Task '$TaskName' has been triggered successfully."
        } catch {
            Write-Host "Failed to run the task: $_" -ForegroundColor Red
            LogThis "Failed to run the task: $_"
        }
    } else {
        Write-Host "You chose not to run the task. Setup is complete." -ForegroundColor Yellow
        LogThis "User chose not to run the task. Setup is complete."
    }
