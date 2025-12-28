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
            [ValidateSet('Info', 'Success', 'Warning', 'Error')]
            [string]$Level = 'Info',
		    [bool]$verboseMessage = $false
	    )

	    try {
            # Verbosity check
            if ($verboseMessage -and -not $verbose) { return }

            # Console Output using semantic streams
            switch ($Level) {
                'Error'   { Write-Error -Message $message -ErrorAction Continue }
                'Warning' { Write-Warning -Message $message }
                'Success' { Write-Information -MessageData "SUCCESS: $message" -InformationAction Continue }
                'Info'    { Write-Information -MessageData $message -InformationAction Continue }
            }

		    # File Logging
		    if ($log) {
                $logEntry = if ($Level -ne 'Info') { "$($Level.ToUpper()): $message" } else { $message }
				Add-Content -Path $logFile -Value $logEntry
			}
	    } catch {
		    Write-Warning "Error in LogThis: $_"
	    }		
    }


# ============= RUNTIME  ==============

    # Include config variables
    if (-Not (Test-Path $ConfigPath)) {
        Write-Error "Configuration file not found: $ConfigPath"
        Pause
        Exit 1
    }
    . $ConfigPath

    LogThis "===== Task Setup script started ====="
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogThis ""
    LogThis "$timestamp === Setup started"

    # Relaunch script as admin if not already running as admin
    if (-not (Test-AdminRights)) {
        LogThis "This script requires administrative privileges. Requesting elevation..." -Level Warning
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
        LogThis "Required script file '$AutoThemeScript' not found. Exiting setup..." -Level Error
        Pause
        Exit 1
    }

    # Check if the scheduled task already exists
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        LogThis "Task '$TaskName' already exists." -Level Warning
        $runExistingTask = Read-Host "Would you like to run the existing task now? (Yes/No)"
        if ($runExistingTask -match '^(Yes|Y)$') {
            try {
                # Run the task using Task Scheduler
                LogThis "Running the existing task via Task Scheduler..." -Level Success
                Start-ScheduledTask -TaskName $TaskName
                LogThis "Task '$TaskName' has been triggered successfully." -Level Success
            } catch {
                LogThis "Failed to run the task: $_" -Level Error
            }
        } else {
            LogThis "You chose not to run the task. Exiting setup..." -Level Warning
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
        LogThis "Creating scheduled task for Windows 10." -verboseMessage $true
        Register-ScheduledTask -TaskName $TaskName -Trigger $Triggers -User "$env:USERNAME" -Action $Action -RunLevel Highest -Compatibility Win8 -Force | Out-Null
    } else {
        LogThis "Creating scheduled task for Windows 11." -verboseMessage $true
        Register-ScheduledTask -TaskName $TaskName -Trigger $Triggers -Action $Action -RunLevel Highest -Force | Out-Null
    }

    LogThis "Scheduled task '$TaskName' created successfully!" -Level Success

    # Prompt the user to run the task immediately
    $runNow = Read-Host "Would you like to run the task now? (Yes/No)"

    if ($runNow -match '^(Yes|Y)$') {
        try {
            # Run the task using Task Scheduler (this simulates the task running as scheduled)
            LogThis "Running the task via Task Scheduler..." -Level Success
            Start-ScheduledTask -TaskName $TaskName
            LogThis "Task '$TaskName' has been triggered successfully." -Level Success
        } catch {
            LogThis "Failed to run the task: $_" -Level Error
        }
    } else {
        LogThis "You chose not to run the task. Setup is complete." -Level Warning
    }
