<#
.SYNOPSIS
    Initial Setup script for Auto Theme.
.DESCRIPTION
    Validates configuration and creates the main Scheduled Task for Auto Theme.
    Run with -Uninstall to remove every task and Registry entry the script creates.
#>

# ============ Script Parameters ==============

param (
    [switch]$Uninstall,   # Remove everything Auto Theme installed, then exit

    <# Internal. The SID of the user who started the un-elevated instance, handed to the
    elevated one so that we clean up the right HKCU hive. See the elevation check below. #>
    [string]$UserSid
)

# ============= Path Variables ==============
$ConfigPath = Join-Path $PSScriptRoot "at-config.ps1"
$AutoThemeScript = Join-Path $PSScriptRoot "at.ps1"
$TaskName = "Auto Theme"

# ============= FUNCTIONS ==============

function Test-AdminRights {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-WindowsVersion {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    if ($os.Caption -match "Windows 10") { return "Windows 10" }
    elseif ($os.Caption -match "Windows 11") { return "Windows 11" }
    return "Other"
}

function LogThis {
    param (
        [string]$message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info',
        [bool]$verboseMessage = $false
    )
    if ($verboseMessage -and -not $verbose) { return }
    switch ($Level) {
        'Error'   { Write-Host "[-] ERROR: $message" -ForegroundColor Red }
        'Warning' { Write-Host "[!] WARNING: $message" -ForegroundColor Yellow }
        'Success' { Write-Host "[+] SUCCESS: $message" -ForegroundColor Green }
        'Info'    { Write-Host "[*] $message" }
    }
    if ($log -and (Test-Path $logFile)) {
        $logEntry = "$(Get-Date -Format 'HH:mm:ss') [$Level] $message"
        Add-Content -Path $logFile -Value $logEntry
    }
}

function Invoke-ConfigValidation {
    LogThis "Checking configuration file..."
    $errors = @()

    # 1. Check Coordinates if not using fixed hours
    if (-not $useFixedHours) {
        # Check if user left the New York defaults
        if ($userLat -eq "40.7128" -and $userLng -eq "-74.0060") {
            $errors += "Coordinates are set to defaults. Please set your Lat/Lng in at-config.ps1 for offline accuracy."
        }
    }

    # 2. Check Theme Files
    if ($useThemeFiles) {
        if (-not (Test-Path $lightPath)) { $errors += "Light Theme missing: $lightPath" }
        if (-not (Test-Path $darkPath)) { $errors += "Dark Theme missing: $darkPath" }
    }

    # 3. Check Wallpaper Folders
    if (-not $noWallpaperChange) {
        if (-not (Test-Path $wallLightPath)) { $errors += "Light Wallpaper path missing: $wallLightPath" }
        if (-not (Test-Path $wallDarkPath)) { $errors += "Dark Wallpaper path missing: $wallDarkPath" }
    }

    # 4. Check Lock Screen
    # These are warnings, not errors: they do not stop the setup, but the feature will misbehave.
    if ($changeLockScreen) {

        # The WinRT API used to set the lock screen does not exist in PowerShell 7
        if ($PSVersionTable.PSEdition -eq 'Core') {
            LogThis "Lock screen changes require Windows PowerShell 5.1. Run the scripts with 'powershell.exe', not 'pwsh.exe'." -Level Warning
        }

        <# If Spotlight is left on, Windows quietly puts its own pictures back,
        which makes it look as though the script is not working. #>
        if (-not $disableSpotlight) {

            $cdmPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
            $rotating = (Get-ItemProperty -Path $cdmPath -Name "RotatingLockScreenEnabled" -ErrorAction SilentlyContinue).RotatingLockScreenEnabled

            if ($rotating -eq 1) {
                LogThis "Windows Spotlight is on while `$disableSpotlight is `$false. Windows will overwrite the lock screen pictures set by this script." -Level Warning
            }
        }
    }

    # 5. Check Extra Apps
    if ($updateTClockColor -and -not (Test-Path $tClockPath)) {
        $errors += "T-Clock enabled but executable not found: $tClockPath"
    }

    if ($errors.Count -gt 0) {
        foreach ($e in $errors) { LogThis $e -Level Error }
        return $false
    }
    return $true
}

function Register-AutoThemeTask {
    LogThis "Preparing Scheduled Task triggers via COM API..."

    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $userSid = $currentUser.User.Value

    # 1. Initialize the Task Scheduler COM object
    $service = New-Object -ComObject Schedule.Service
    $service.Connect()
    $rootFolder = $service.GetFolder("\")

    # 2. Create a new Task Definition
    $taskDefinition = $service.NewTask(0)
    $taskDefinition.RegistrationInfo.Description = "Main Auto Theme management task for scheduling sunrise/sunset events."
    $taskDefinition.Principal.UserId = $userSid
    $taskDefinition.Principal.LogonType = 3 # TASK_LOGON_INTERACTIVE_TOKEN
    $taskDefinition.Principal.RunLevel = 1  # TASK_RUNLEVEL_HIGHEST

    # 3. Define Settings (COM API Property Names)
    # Note: COM uses 'Disallow' logic (set to false to ALLOW)
    $taskDefinition.Settings.DisallowStartIfOnBatteries = $false
    $taskDefinition.Settings.StopIfGoingOnBatteries = $false
    $taskDefinition.Settings.StartWhenAvailable = $true
    $taskDefinition.Settings.Enabled = $true
    $taskDefinition.Settings.Hidden = $false
    $taskDefinition.Settings.Compatibility = 4 # 4 = Windows 8/10

    # Optional: Prevents Windows from stopping the task after 3 days
    $taskDefinition.Settings.ExecutionTimeLimit = "PT0S"

    # 4. Add Triggers
    # Trigger 1: Logon
    $logonTrigger = $taskDefinition.Triggers.Create(9) # TASK_TRIGGER_LOGON
    $logonTrigger.UserId = $userSid
    $logonTrigger.Enabled = $true

    # Trigger 2: Startup
    $startupTrigger = $taskDefinition.Triggers.Create(8) # TASK_TRIGGER_BOOT
    $startupTrigger.Enabled = $true

    # Trigger 3: Session Unlock
    $unlockTrigger = $taskDefinition.Triggers.Create(11) # TASK_TRIGGER_SESSION_STATE_CHANGE
    $unlockTrigger.StateChange = 8 # TASK_SESSION_UNLOCK
    $unlockTrigger.UserId = $userSid
    $unlockTrigger.Enabled = $true

    # 5. Define Action
    # Visibility / Executable Logic
    $psArgs = "-WindowStyle Hidden -ExecutionPolicy Bypass -NoProfile -File `"$AutoThemeScript`" -Schedule" # -Schedule is a script parameter
    switch ($terminalVisibility) {
        "ch" { $exe = "conhost.exe"; $taskArgs = "--headless PowerShell.exe $psArgs" }
        "wt" {
            if (Get-Command "wt.exe" -ErrorAction SilentlyContinue) {
                $exe = "wt.exe"; $taskArgs = "-w 0 nt PowerShell.exe $psArgs"
            } else {
                $exe = "PowerShell.exe"; $taskArgs = $psArgs
            }
        }
        default { $exe = "PowerShell.exe"; $taskArgs = $psArgs }
    }

    $action = $taskDefinition.Actions.Create(0) # TASK_ACTION_EXEC
    $action.Path = $exe
    $action.Arguments = $taskArgs

    # 6. Save (Register) the Task
    try {
        # 6 = TASK_CREATE_OR_UPDATE
        $rootFolder.RegisterTaskDefinition($TaskName, $taskDefinition, 6, $null, $null, 3) | Out-Null
        LogThis "Main Scheduled Task registered successfully via COM." -Level Success
        return $true
    } catch {
        LogThis "COM Registration failed: $_" -Level Error
        return $false
    }
}

<# Delete one Registry key and everything under it, but only if it is actually there.
Takes a 'Registry::' path, so that HKCR and a hive addressed by SID both work. #>
function Remove-RegKeyIfPresent {
    param (
        [string]$Path,
        [string]$Description
    )

    try {

        if (-not (Test-Path $Path)) {
            LogThis "$Description : not present."
            return $false
        }

        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        LogThis "Removed $Description" -Level Success
        return $true

    } catch {

        LogThis "Could not remove $Description : $_" -Level Error
        return $false
    }
}

<# Take out everything Auto Theme has ever put on the machine.
We look for each item rather than trusting the current configuration: someone who has since
turned `$writeRegistry` or `$addContextMenu` back off still has the entries an earlier run
created, and those are exactly the ones they are asking us to clear away. #>
function Invoke-AutoThemeUninstall {

    # Anchored, unlike the reinstall prompt below: this one throws things away
    $choice = Read-Host "Remove all Auto Theme scheduled tasks and Registry entries? (Y/N)"
    if ($choice -notmatch '^\s*y') {
        LogThis "Uninstall aborted by user." -Level Info
        return
    }

    $removed = 0

    # ---- 1. Scheduled Tasks ----
    <# The main task comes from this script, the rest are created by at.ps1 as it runs.
    All six use fixed names, so we can just walk the list. #>
    $taskNames = @(
        "Auto Theme",
        "Auto Theme wallpaper changer",
        "Auto Theme sunrise",
        "Auto Theme sunset",
        "Auto Theme fixed sunrise",
        "Auto Theme fixed sunset"
    )

    foreach ($name in $taskNames) {

        try {

            if (-not (Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue)) {
                LogThis "Task '$name' : not present."
                continue
            }

            Unregister-ScheduledTask -TaskName $name -Confirm:$false -ErrorAction Stop
            LogThis "Removed task '$name'" -Level Success
            $removed++

        } catch {

            # One task we cannot touch must not stop us clearing the others
            LogThis "Could not remove task '$name' : $_" -Level Error
        }
    }

    # ---- 2. Context menu ----
    if (Remove-RegKeyIfPresent -Path "Registry::HKEY_CLASSES_ROOT\DesktopBackground\Shell\AutoTheme" `
                               -Description "desktop context menu") { $removed++ }

    # Left behind by the old context-menu.reg, from before the cascading menu
    if (Remove-RegKeyIfPresent -Path "Registry::HKEY_CLASSES_ROOT\DesktopBackground\Shell\AutoThemeNext" `
                               -Description "legacy 'Next wallpaper' menu entry") { $removed++ }

    # ---- 3. Per-user entries ----
    <# Normally this is simply HKCU. But when the UAC prompt was answered with a different
    Administrator account, our HKCU is that account's, and the entries we want belong to the
    user who launched us, so we reach their hive through the SID they passed us. #>
    $userHives = @("Registry::HKEY_CURRENT_USER")

    if ($UserSid -and $UserSid -ne ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value)) {
        $userHives += "Registry::HKEY_USERS\$UserSid"
    }

    foreach ($hive in $userHives) {

        # The state key, whatever `$writeRegistry` happens to say today
        $tuscoss = "$hive\SOFTWARE\Tuscoss"

        if (Remove-RegKeyIfPresent -Path "$tuscoss\AutoTheme" -Description "state key ($tuscoss\AutoTheme)") { $removed++ }

        <# Tidy away our own parent key, but only once it holds nothing else:
        another Tuscoss product may well be keeping its settings in there. #>
        try {

            if (Test-Path $tuscoss) {

                $key = Get-Item -Path $tuscoss
                if ($key.SubKeyCount -eq 0 -and $key.ValueCount -eq 0) {
                    Remove-Item -Path $tuscoss -Force -ErrorAction Stop
                    LogThis "Removed empty parent key ($tuscoss)" -Level Success
                }
            }

        } catch {

            LogThis "Could not remove parent key ($tuscoss) : $_" -Level Error
        }

        <# Where the wallpaper name used to live, before it moved into the state key.
        Windows does not put this value there itself, so it is safe to clear. #>
        try {

            $desktop = "$hive\Control Panel\Desktop"

            if ((Test-Path $desktop) -and (Get-ItemProperty -Path $desktop -Name "WallpaperName" -ErrorAction SilentlyContinue)) {

                Remove-ItemProperty -Path $desktop -Name "WallpaperName" -Force -ErrorAction Stop
                LogThis "Removed legacy WallpaperName value ($desktop)" -Level Success
                $removed++
            }

        } catch {

            LogThis "Could not remove legacy WallpaperName value : $_" -Level Error
        }
    }

    # ---- 4. Summary ----
    if ($removed -eq 0) {
        LogThis "Nothing left to remove: Auto Theme was already uninstalled." -Level Info
    } else {
        LogThis "Uninstall complete. $removed item(s) removed." -Level Success
    }

    LogThis "The script folder is left as it is. Delete it, along with 'at-state.ini' and the log files, to finish." -Level Info
}

# ============= RUNTIME ==============

# 1. Elevation Check
if (-not (Test-AdminRights)) {
    Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow

    <# Pass our own SID along. If the user answers the UAC prompt with a different
    Administrator account, the elevated instance gets that account's HKCU, and the state key
    we are asked to remove lives in this one. Uninstall reads the hive by SID instead. #>
    $callerSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value

    # Arguments do not survive a plain relaunch, so -Uninstall has to be forwarded by hand
    $relaunchArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -UserSid $callerSid"
    if ($Uninstall) { $relaunchArgs += " -Uninstall" }

    Start-Process "powershell.exe" -ArgumentList $relaunchArgs -Verb RunAs
    Exit 0
}

# 2. Load Config
if (-not (Test-Path $ConfigPath)) {
    Write-Error "at-config.ps1 not found. Please create it first."
    Pause; Exit 1
}

<# Loading these files normally creates the state store as a side effect. When uninstalling we
want the settings only: creating the key or the ini here, seconds before removing them, would
leave a brand new 'at-state.ini' sitting in a folder we have just cleaned out. #>
$skipStateProvision = $Uninstall.IsPresent

. $ConfigPath

<# The state store lives in the worker, which we dot-source for its definitions: with no
`-Path` it stops after them rather than changing any wallpaper. Provisioning the store is
part of setting Auto Theme up, so a missing worker is worth a warning, but not a refusal:
everything else this script installs is still worth having. #>
if (Test-Path $workerPath) {
    . $workerPath
} else {
    LogThis "Worker script not found at $workerPath. The state store will be created on the first run instead." -Level Warning
}

<# Uninstall runs before validation on purpose. A half-edited config, with theme or wallpaper
paths that no longer exist, is precisely when someone reaches for uninstall, and the validation
below would turn them away. #>
if ($Uninstall) {

    LogThis "===== Auto Theme Uninstall Started ====="
    Invoke-AutoThemeUninstall
    LogThis "===== Uninstall Finished ====="

    Pause; Exit 0
}

LogThis "===== Auto Theme Setup Started ====="

# 3. Validation
if (-not (Invoke-ConfigValidation)) {
    LogThis "Validation failed. Please fix errors or missing values in at-config.ps1 and try again." -Level Warning
    Pause; Exit 1
}

# 4. Check Script Presence
if (-not (Test-Path $AutoThemeScript)) {
    LogThis "Core script 'at.ps1' missing from folder." -Level Error
    Pause; Exit 1
}

# ============ Context Menu Setup ============
if ($addContextMenu) {

    LogThis "Installing 'Auto Theme' cascading desktop menu..."

    $parentKey = "HKEY_CLASSES_ROOT\DesktopBackground\Shell\AutoTheme"
    $subKey = "$parentKey\shell"

    # Commands
    # Next Wallpaper (no console window)
    $psNext = 'conhost.exe --headless PowerShell.exe -ExecutionPolicy Bypass -NoProfile -Command "& {0} -Next"' -f "'$AutoThemeScript'"
    # Toggle Theme
    $psToggle = 'PowerShell.exe -ExecutionPolicy Bypass -NoProfile -Command "& {0} -Toggle"' -f "'$AutoThemeScript'"
    # Refresh Schedule
    $psSched = 'PowerShell.exe -ExecutionPolicy Bypass -NoProfile -Command "& {0} -Schedule"' -f "'$AutoThemeScript'"

    try {
        # 1. Parent Entry (No changes here)
        New-Item -Path "Registry::$parentKey" -Force | Out-Null
        Set-ItemProperty -Path "Registry::$parentKey" -Name "MUIVerb" -Value "Auto Theme"
        Set-ItemProperty -Path "Registry::$parentKey" -Name "Icon" -Value "imageres.dll,-114"
        Set-ItemProperty -Path "Registry::$parentKey" -Name "SubCommands" -Value ""
        Set-ItemProperty -Path "Registry::$parentKey" -Name "SeparatorBefore" -Value ""
        Set-ItemProperty -Path "Registry::$parentKey" -Name "SeparatorAfter" -Value ""

        # 2. Next Wallpaper (Top)
        $nextPath = "$subKey\Next"
        New-Item -Path "Registry::$nextPath\command" -Force | Out-Null
        Set-ItemProperty -Path "Registry::$nextPath" -Name "(Default)" -Value "Next background picture"
        Set-ItemProperty -Path "Registry::$nextPath" -Name "Icon" -Value "imageres.dll,-21"
        # $psNext = "conhost.exe --headless PowerShell.exe -ExecutionPolicy Bypass -NoProfile -Command `\"& '$AutoThemeScript' -Next`\""
        Set-ItemProperty -Path "Registry::$nextPath\command" -Name "(Default)" -Value $psNext

        # 3. Toggle Theme (Middle)
        $togglePath = "$subKey\Toggle"
        New-Item -Path "Registry::$togglePath\command" -Force | Out-Null
        Set-ItemProperty -Path "Registry::$togglePath" -Name "(Default)" -Value "Toggle Theme Now"
        Set-ItemProperty -Path "Registry::$togglePath" -Name "Icon" -Value "themecpl.dll,-1"
        # $psToggle = "PowerShell.exe -ExecutionPolicy Bypass -NoProfile -Command `\"& '$AutoThemeScript' -Toggle`\""
        Set-ItemProperty -Path "Registry::$togglePath\command" -Name "(Default)" -Value $psToggle

        # 4. Refresh Schedule (Pinned to Bottom)
        $refreshPath = "$subKey\TRefresh"
        New-Item -Path "Registry::$refreshPath\command" -Force | Out-Null
        Set-ItemProperty -Path "Registry::$refreshPath" -Name "(Default)" -Value "Refresh Schedule"
        Set-ItemProperty -Path "Registry::$refreshPath" -Name "Icon" -Value "Shell32.dll,-16741"
        Set-ItemProperty -Path "Registry::$parentKey" -Name "SeparatorBefore" -Value ""
        # This is the trick to move it to the end:
        Set-ItemProperty -Path "Registry::$refreshPath" -Name "Position" -Value "Bottom"

        # $psSched = "PowerShell.exe -ExecutionPolicy Bypass -NoProfile -Command `\"& '$AutoThemeScript' -Schedule`\""
        Set-ItemProperty -Path "Registry::$refreshPath\command" -Name "(Default)" -Value $psSched

        LogThis "Cascading menu reordered: Refresh is now at the bottom."
    } catch {
        LogThis "Failed to install ordered context menu: $_" -Level Error
    }

} else {
    LogThis "Context menu installation skipped as per configuration." -Level Info
}

# 5. Task Registration
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    $choice = Read-Host "Task already exists. Reinstall? (Y/N)"
    if ($choice -notmatch 'y') {
        LogThis "Setup aborted by user." -Level Info
        Exit 0
    }
}

if (Register-AutoThemeTask) {
    LogThis "Setup complete. Launching 'Auto Theme' task to apply initial settings..." -Level Success

    # We trigger it automatically to ensure the system state is synced immediately
    Start-ScheduledTask -TaskName $TaskName

    LogThis "Task is now running in the background." -Level Info
} else {
    LogThis "Setup failed during task registration." -Level Error
}

LogThis "===== Setup Finished ====="
# Pause