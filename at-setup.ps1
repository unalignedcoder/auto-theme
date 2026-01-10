<#
.SYNOPSIS
    Initial Setup script for Auto Theme.

.DESCRIPTION
    Creates the initial Scheduled Task for the script AutoTheme. 
    Automatically requests admin privileges if not run as admin.

.LINK
	https://github.com/unalignedcoder/auto-theme/

.NOTES
    - See main script file for latest changes.
#>

# ============= Config file ==============

    $ConfigPath = Join-Path $PSScriptRoot "at-config.ps1"

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
        # -NoProfile to ensure a clean environment during elevation
        $proc = Start-Process -FilePath "powershell.exe" `
                    -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
                    -Verb RunAs -PassThru
        # Exit the non-elevated instance
        Exit 0
    }

    # Define the path to the at.ps1 script and task XML file
    $AutoThemeScript = Join-Path -Path $PSScriptRoot -ChildPath "at.ps1"
    $TaskName = "Auto Theme"

    # Check if at.ps1 exists
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

    # IMPROVEMENT: Use the Current User's SID for the Unlock Trigger to prevent "User not found" errors
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $userSid = $currentUser.User.Value

    $StateChangeTrigger = Get-CimClass `
        -Namespace ROOT\Microsoft\Windows\TaskScheduler `
        -ClassName MSFT_TaskSessionStateChangeTrigger

    $UnlockTrigger = New-CimInstance `
        -CimClass $StateChangeTrigger `
        -Property @{
            StateChange = 8  # 8 = TASK_SESSION_UNLOCK
            UserId = $userSid # Using SID instead of Username string
        } `
        -ClientOnly

    $Triggers = @($LogonTrigger, $startupTrigger, $UnlockTrigger)

    # --- Terminal Visibility Logic ---
    $psArgs = "-WindowStyle Hidden -ExecutionPolicy Bypass -NoProfile -File `"$AutoThemeScript`""
    
    switch ($terminalVisibility) {
        "ch" {
            LogThis "Visibility mode: Invisible (conhost --headless)"
            $exe = "conhost.exe"
            $arguments = "--headless PowerShell.exe $psArgs"
        }
        "wt" {
            if (Get-Command "wt.exe" -ErrorAction SilentlyContinue) {
                LogThis "Visibility mode: Windows Terminal"
                $exe = "wt.exe"
                $arguments = "-w 0 nt PowerShell.exe $psArgs"
            } else {
                LogThis "Warning: wt.exe not found. Falling back to default (ps) mode." -Level Warning
                $exe = "PowerShell.exe"
                $arguments = $psArgs
            }
        }
        default {
            LogThis "Visibility mode: PowerShell Console"
            $exe = "PowerShell.exe"
            $arguments = $psArgs
        }
    }

    # 1. Define common settings used by both windows 10 and 11
    $CommonSettings = @{
        AllowStartIfOnBatteries      = $true
        DontStopIfGoingOnBatteries   = $true
        StartWhenAvailable           = $true
    }

    # Create the action
    $Action = New-ScheduledTaskAction -Execute $exe -Argument $arguments

    # Register the task
    $windowsVersion = Get-WindowsVersion
    $TaskDescription = "Main Auto Theme task for scheduling sunrise/sunset events."
    
    try {

        if ($windowsVersion -eq "Windows 10") {

            LogThis "Creating scheduled task for Windows 10 (Win8 Compatibility)." -verboseMessage $true
            
            $Settings = New-ScheduledTaskSettingsSet @CommonSettings -Compatibility Win8
            
            Register-ScheduledTask -TaskName $TaskName -Trigger $Triggers -User $userSid `
                -Action $Action -Settings $Settings -Description $TaskDescription `
                -RunLevel Highest -Force | Out-Null

        } else {

            LogThis "Creating scheduled task for Windows 11 (Native Compatibility)." -verboseMessage $true
            
            # Windows 11 uses the most modern task engine by default
            $Settings = New-ScheduledTaskSettingsSet @CommonSettings
            
            Register-ScheduledTask -TaskName $TaskName -Trigger $Triggers -User $userSid `
                -Action $Action -Settings $Settings -Description $TaskDescription `
                -RunLevel Highest -Force | Out-Null
        }

        LogThis "Scheduled task '$TaskName' created successfully!" -Level Success

    } catch {

        LogThis "Critical error registering task: $_" -Level Error
        Pause
        Exit 1
    }

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