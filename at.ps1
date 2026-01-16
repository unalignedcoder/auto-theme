<#
.SYNOPSIS
	Changes the active Windows theme based on a predefined/daylight schedule. Works in Windows 10/11.

.DESCRIPTION
	This highly-sophisticated Powershell script automatically switches the Windows Theme depending on Sunrise and Sunset, or hours set by the user.
	It can activate Windows Dark and Light mode directly, also handling wallpaper changes natively.
	It can also activate given `.theme` files, which may allow for a higher degree of customization and compatibility.
	The script is designed to run in the background as a scheduled task, ensuring that the system theme is updated without user intervention.
	It will automatically create the next temporary task for the next daylight event.
	Such tasks ("Auto Theme sunrise" and "Auto Theme sunset") will be overwritten as a matter of course to avoid clutter.
	It only connects to the internet to verify Location, if it cannot be retrieved from the system.
	It can stay completely offline operating on fixed hours provided by the user.
	When ran as the command `./at.ps1` from terminal or desktop shortcut, the script will only toggle between themes.
	IMPORTANT: Edit `at-config.ps1` to configure this script. The file contains all necessary explanations.
	OPTIONALLY: Run `./at-setup.ps1` to create the main Scheduled Task and verify the configuration is in place.
	For more information, refer to the README file.

.LINK
	https://github.com/unalignedcoder/auto-theme/

.NOTES
	- Fixed a parsing error in the Setup script
#>

# ============ Script Parameters ==============

param (
    [switch]$Toggle,   # Force a theme flip
    [switch]$Schedule, # Force the daylight calculation
    [switch]$Next      # Force a wallpaper skip
)

# ============= Script Version ==============

	# This is automatically updated
	$scriptVersion = "1.0.46"

# ============= Config file ==============

	$ConfigPath = Join-Path $PSScriptRoot "at-config.ps1"

# ============= FUNCTIONS  ==============

	# Determine if the script runs interactively
    function Test-TerminalSession {

        # Get the current process ID
        $proc = Get-CimInstance Win32_Process -Filter "ProcessId = $pid"
        # Get the parent process ID
        $parent = Get-CimInstance Win32_Process -Filter "ProcessId = $($proc.ParentProcessId)"

        # Check for svchost or Schedule (Standard Task Scheduler)
        if ($parent.Name -eq "svchost.exe" -or $parent.CommandLine -like "*Schedule*") {
            return $false
        }

        # Use the config variable to detect the specific launcher used
        switch ($terminalVisibility) {
            "ch" {
                # If we expect headless, check if parent is conhost with the headless flag
                if ($parent.Name -eq "conhost.exe" -and $parent.CommandLine -like "*--headless*") { return $false }
            }
            "wt" {
                # If we expect Windows Terminal, check if parent is wt.exe
                if ($parent.Name -eq "wt.exe") { return $false }
            }
        }

        # Default to true (Interactive) if no scheduled parent is detected
        return $true
    }

	# Create the logging system
	function Write-Log {
		param (
			[string]$message,
			[bool]$verboseMessage = $false   # Default to false if not specified
		)

		try {

			# Only proceed if in debug mode
			if (-not $log) { return }

			<# Check for verbosity:
			If the message is verbose, but verbose is false, end the Function
			If the message is verbose, but verbose is true, continue
			If the message is not verbose, continue #>
			if ($verboseMessage -and -not $verbose) { return }

			# Interactive: print to console (no prefix). Verbose messages are shown only when $verbose is true.
			if (Test-TerminalSession) {

				Write-Information -MessageData $message -InformationAction Continue

				# Optionally also append to the log file when requested from interactive sessions
				if ($logFromTerminal) { Add-Content -Path $logFile -Value $message }
			}
			else {
				# Non-interactive: always append to log file
				Add-Content -Path $logFile -Value $message
			}

		} catch {

			# Write-Host avoids emitting pipeline output from the logger
			Write-Information -MessageData "Error in Write-Log: $_" -InformationAction Continue
		}
	}

	# Trim old log entries
	function Limit-LogHistory {
		param (
			[string]$logFilePath,    # Path to the log file
			[int]$maxSessions = 5  # Maximum number of log sessions to keep
		)

		if (-Not ($trimLog)) {
			return
		}

		if (-Not (Test-Path $logFilePath)) {
			# Log file doesn't exist, no need to trim
			Write-Output "Log file doesn't exist, no need to trim"
			return
		}

		# Read all lines from the log file
		Write-Output "Reading all lines from the log file"
		$logLines = Get-Content -Path $logFilePath

		# Find the indices of all session start lines
		$sessionStartIndices = @()
		for ($i = 0; $i -lt $logLines.Count; $i++) {
			if ($logLines[$i] -match '=== Auto-Theme script started \(Version: .*?\)') {
				$sessionStartIndices += $i
			}
		}

		# Check if the number of sessions exceeds the maximum allowed
		if ($sessionStartIndices.Count -le $maxSessions) {
			# No need to trim the log
			Write-Output "Log file is small, no need to trim"
			return
		}

		# Calculate how many sessions to remove
		$sessionsToRemove = $sessionStartIndices.Count - $maxSessions

		# Identify the range of lines to keep
		$startIndexToKeep = $sessionStartIndices[$sessionsToRemove]

		# Extract the lines to keep and overwrite the log file
		Write-Output "Extracting the log lines to keep, and overwriting the log file"
		$linesToKeep = $logLines[$startIndexToKeep..($logLines.Count - 1)]
		Set-Content -Path $logFilePath -Value $linesToKeep
	}

	# Handle BurntToast Notifications
	function Show-Notification {
		param(
			# Accept either a single string or an array of strings.
			[object]$Text,
			[string]$AppLogo
		)

		# Install the BurntToast module if not already installed
		if (-not (Get-Module -Name BurntToast -ListAvailable)) {

			try {

				Write-Log "Installing the BurnToast Notifications module"
				Install-Module -Name BurntToast -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -Confirm:$false

			} catch {

				Write-Log "Failed to install BurntToast module: $_"
				return
			}
		}

		# for when the above is commented out or fails, we double-check.
		if (Get-Module -Name BurntToast -ListAvailable) {

			try {

				Write-Log "Creating BurnToast notification"  -verboseMessage $true

				# create Burntoast header
				$BurnHeader = New-BTHeader -Title 'Auto Theme'

				# show notification
				New-BurntToastNotification -Header $BurnHeader -Text $Text -AppLogo $AppLogo

				Write-Log "Displayed BurntToast notification with text: $(
					if ($Text -is [System.Array]) { ($Text -join ' | ') } else { $Text }
				)"  -verboseMessage $true

			} catch {

				Write-Log "Error displaying BurntToast notification: $_"  -verboseMessage $true
			}

		} else {

			Write-Log "BurntToast module is not installed. Cannot display system notifications."
		}
	}

	# Prepare BurntToast notification content
	function Send-ThemeNotification {
		param(
			[string]$MainLine,
			[string]$SelectedWallpaper = ""
		)

		# If user disabled showing wallpaper names, show only main line
		if ($SelectedWallpaper -and $showWallName) {

			# Blank line for spacing, then wallpaper line
			Show-Notification -Text @($MainLine, "", "Wallpaper: $SelectedWallpaper") -AppLogo $appLogo

		} else {

			Show-Notification -Text $MainLine -AppLogo $appLogo
		}
	}

	# Check if the script has been run in the last interval
	function Test-LastRunTime {

		Write-Log "Checking if script was run in the last $lastRunInterval minutes"  -verboseMessage $true

		if (Test-Path $lastRunFile) {

			$lastRun = Get-Content $lastRunFile | Out-String
			$lastRun = [DateTime]::Parse($lastRun)
			$now = Get-Date

			$timeSinceLastRun = $now - $lastRun

			if ($timeSinceLastRun.TotalMinutes -lt $lastRunInterval) {

				Write-Log "Script was run within the last $lastRunInterval minutes. Exiting."
				exit
			}
		}
	}

	# Update the last run time
	function Update-LastRunTime {

		$now = Get-Date
		$now | Out-File -FilePath $lastRunFile -Force
	}

	# Check whether wallpaper shuffle is enabled in .theme file
	function Test-DoWeShuffle {
		param (
			[string]$themeFilePath
		)

		Write-Log "Checking if the .theme file shuffles wallpapers" -verboseMessage $true

		# Read the content of the theme file
		$themeContent = Get-Content -Path $themeFilePath

		# Flag to indicate if we are inside the [Slideshow] section
		$inSlideshowSection = $false

		foreach ($line in $themeContent) {
			# Write-Log "Processing line: $line" -verboseMessage $true

			# Check for the start of the [Slideshow] section
			if ($line -match '^\[Slideshow\]') {
				Write-Log "Found Slideshow section" -verboseMessage $true
				$inSlideshowSection = $true
				continue
			}

			# If we are inside the [Slideshow] section, look for 'shuffle' setting
			if ($inSlideshowSection) {

				if ($line -match '(?i)shuffle=(\d)') { # Case-insensitive match for 'shuffle'
					Write-Log "Found shuffle setting: $line" -verboseMessage $true
					return $matches[1] -eq '1'
				}

				# If we encounter the next section or end of file, break out of the loop
				if ($line -match '^\[.*\]') {
					Write-Log "Leaving [Slideshow] section" -verboseMessage $true
					break
				}
			}
		}

		# If no shuffle setting is found, return false
		Write-Log "No, the .theme file does not." -verboseMessage $true
		return $false
	}

	<# Prepend the substring '_0_at_' to one randomly chosen
	wallpaper filename, so as to make it first pick. #>
	function Get-RandomFirstWall {
		param (
			[string]$wallpaperDirectory
		)

		# Will return basename (no path, no extension) or empty string
		[string]$SelectedWallpaperBasename = ""

		# Removed the unhelpful "Get-RandomFirstWall: entry" log line

		if (-Not ($randomFirst)) {
			Write-Log "The first wallpaper will not be randomized."  -verboseMessage $true
			return $SelectedWallpaperBasename
		}

		Write-Log "Randomizing first wallpaper."  -verboseMessage $true
		Write-Log "Looking in $wallpaperDirectory"  -verboseMessage $true

		# Build list of folders to sanitize (remove any existing _0_at_ prefixes)
		$dirsToSanitize = @($wallpaperDirectory)

		# If global light/dark wallpaper paths exist and are different, include them
		if ($null -ne $wallLightPath -and $wallLightPath -ne "" -and $wallLightPath -ne $wallpaperDirectory) {
			$dirsToSanitize += $wallLightPath
		}
		if ($null -ne $wallDarkPath -and $wallDarkPath -ne "" -and $wallDarkPath -ne $wallpaperDirectory -and $wallDarkPath -ne $wallLightPath) {
			$dirsToSanitize += $wallDarkPath
		}

		# Deduplicate and filter to existing directories
		$dirsToSanitize = $dirsToSanitize | Where-Object { $_ } | Get-Unique

		foreach ($dir in $dirsToSanitize) {

			if (-Not (Test-Path $dir)) {
				Write-Log "Wallpaper folder not found: $dir" -verboseMessage $true
				continue
			}

			# Retrieve all wallpaper files in this directory
			$wallpapers = Get-ChildItem -Path $dir -File -ErrorAction SilentlyContinue
			if (-not $wallpapers) {
				Write-Log "No wallpapers in $dir" -verboseMessage $true
				continue
			}

			# Find renamed files with prefix and restore original names
			$existingRenamedWallpapers = $wallpapers | Where-Object { $_.Name -match '^_0_at_' }

			if ($existingRenamedWallpapers.Count -gt 0) {

				foreach ($wallpaper in $existingRenamedWallpapers) {
					try {
						$originalName = $wallpaper.Name -replace '^_0_at_', ''
						Write-Log "Restoring original name: $($wallpaper.FullName) Ã¢â€ â€™ $originalName" -verboseMessage $true
						# Use only the name in -NewName to avoid path issues
						Rename-Item -Path $wallpaper.FullName -NewName $originalName -Force -ErrorAction Stop | Out-Null
					} catch {
						Write-Log "Failed restoring $($wallpaper.FullName): $_" -verboseMessage $true
					}
				}
			}
		}

		# Now operate on the requested target directory to pick and prefix a random wallpaper.
		if (-Not (Test-Path $wallpaperDirectory)) {
			Write-Log "Target wallpaper directory not found: $wallpaperDirectory" -verboseMessage $true
			return $SelectedWallpaperBasename
		}

		# Refresh the list of wallpapers in the target folder and exclude already-prefixed files
		$wallpapers = Get-ChildItem -Path $wallpaperDirectory -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '^_0_at_' }

		# Ensure there are wallpapers available
		if (-Not $wallpapers -or $wallpapers.Count -eq 0) {
			Write-Log "No wallpapers available in target folder: $wallpaperDirectory" -verboseMessage $true
			return $SelectedWallpaperBasename
		}

		# Select a random wallpaper and rename it with the prefix
		$randomFirstWallpaper = $wallpapers | Get-Random
		$newWallpaperName = "_0_at_" + $randomFirstWallpaper.Name

		try {
			Rename-Item -Path $randomFirstWallpaper.FullName -NewName $newWallpaperName -Force -ErrorAction Stop | Out-Null
			$newWallpaperNameFull = Join-Path $wallpaperDirectory $newWallpaperName
			Write-Log "Renamed $($randomFirstWallpaper.FullName) to $newWallpaperNameFull" -verboseMessage $true

			# Prepare basename (no path, no extension) to return
			$SelectedWallpaperBasename = [System.IO.Path]::GetFileNameWithoutExtension($randomFirstWallpaper.Name)

		} catch {
			Write-Log "Failed to rename $($randomFirstWallpaper.FullName): $_" -verboseMessage $true
			$SelectedWallpaperBasename = ""
		}

		# Return only the basename string
		return $SelectedWallpaperBasename
	}

	# Get the currently selected wallpaper's basename (no path, no extension)
	function Get-WallpaperName {
        param (
            [string]$wallpaperDirectory,
            [string]$themeFilePath
        )

		if ([string]::IsNullOrWhiteSpace($wallpaperDirectory)) {
            Write-Log "Get-WallpaperName: wallpaperDirectory is null or empty." -verboseMessage $true
            return ""
        }

        [string]$selected = ""

        try {

			# return empty string if wallpaper path is not found
            if (-not (Test-Path $wallpaperDirectory)) { return $selected }

            # If theme uses shuffle, use the "Rename Trick" to force a random start
            if ($useThemeFiles -and (Test-DoWeShuffle -themeFilePath $themeFilePath)) {

                $selected = Get-RandomFirstWall -wallpaperDirectory $wallpaperDirectory

            }

            # Fallback (or if static): Just get the first file alphabetically
            if (-not $selected) {

                $first = Get-ChildItem -Path $wallpaperDirectory -File | Sort-Object Name | Select-Object -First 1
                if ($first) { $selected = [System.IO.Path]::GetFileNameWithoutExtension($first.Name) }
            }

        } catch {
            Write-Log "Get-WallpaperName error: $_"
        }

        if ($selected) { $selected = $selected -replace '^_0_at_', '' }

        return $selected
    }

	# Handles the creation and management of the native wallpaper slideshow task
	function Set-WallpaperRotationTask {
		param (
			[string]$Mode # "light" or "dark"
		)

		# 1. Setup Variables
		$TaskName = "Auto Theme wallpaper changer"
		$wallPathToPass = if ($Mode -eq "light") { $wallLightPath } else { $wallDarkPath }
		$wallpaperDesc = "Triggers wallpaper change based on user-defined intervals. Managed by the main Auto Theme task."

		# --- THE LOGIC CHECK ---
		$isFolder = Test-Path $wallPathToPass -PathType Container
		if ($useThemeFiles -or $slideShowInterval -le 0 -or $noWallpaperChange -or -not $isFolder) {
			if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
				Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
				Write-Log "Wallpaper rotation not needed (Fixed image or disabled). Removed task." -verboseMessage $true
			}
			return
		}

		# 2. Verify worker script exists
		$workerScript = Join-Path $PSScriptRoot "at-wallpaper.ps1"
		if (-not (Test-Path $workerScript)) {
			Write-Log "Error: Worker script not found at $workerScript. Cannot schedule rotation."
			return
		}

		# 3. Define the Start Time (1 interval in the future)
		$triggerTime = (Get-Date).AddMinutes([int]$slideShowInterval)

		# 4. Prepare specific arguments for the worker script
		$workerArgs = "-Path `"$wallPathToPass`""

		try {
			# 5. Call the Helper with the CUSTOM script and arguments
			# This tells the task to run 'at-wallpaper.ps1' instead of 'at.ps1'
			Register-Task -Name $TaskName `
						-NextTriggerTime $triggerTime `
						-Description $wallpaperDesc `
						-CustomFile $workerScript `
						-Arguments $workerArgs

			# 6. Apply the Repetition Patch
			# Register-Task creates a standard trigger; we must modify it for repeating intervals.
			$task = Get-ScheduledTask -TaskName $TaskName

			# We switch to a Daily trigger so it persists across reboots
			$task.Triggers[0] = New-ScheduledTaskTrigger -Daily -At ($triggerTime.ToString("HH:mm"))

			# ISO 8601 format (PT#M) prevents the XML formatting error
			$task.Triggers[0].Repetition.Interval = "PT$($slideShowInterval)M"
			$task.Triggers[0].Repetition.Duration = "" # Indefinite repetition

			# Ensure it runs even if the PC was off during a scheduled tick
			$task.Settings.StartWhenAvailable = $true
			$task.Settings.MultipleInstances = "IgnoreNew"

			Set-ScheduledTask -InputObject $task | Out-Null

			Write-Log "Native Slideshow scheduled: Every $slideShowInterval minutes using $Mode wallpapers."
			Write-Log "Worker: $workerScript" -verboseMessage $true


		} catch {

			Write-Log "Failed to register wallpaper rotation task: $_"
		}
	}

	# Helper to launch processes without admin privileges
    function Start-ProcessUnelevated {
        param (
            [string]$FilePath,
            [string]$ArgumentList = ""
        )
        try {
            <# Use the Shell COM object to ask the standard user desktop shell
            to launch the process for us, bypassing the script's Admin token. #>
            $shell = New-Object -ComObject Shell.Application
            $shell.ShellExecute($FilePath, $ArgumentList, "", "open", 1)
            Write-Log "Launched unelevated: $FilePath" -verboseMessage $true
        } catch {
            Write-Log "Failed to launch unelevated ($FilePath): $_"
        }
    }

	# Restart the 'Themes' Service
	function Restart-ThemeService {

		[bool]$IsAdmin = Test-IsAdmin
		if ($restartThemeService -and $IsAdmin) {

			try {

				Write-Log "Restarting the Themes service." -verboseMessage $true

				Restart-Service -Name "Themes" -Force -ErrorAction SilentlyContinue

				Write-Log "Themes service restarted successfully." -verboseMessage $true

			} catch {

				Write-Log "Failed to restart Themes service: $_"  -verboseMessage $true
			}
		}
	}

	# Modify TrueLaunchBar default colors
	function Update-TrueLaunch {

		param (
			[string]$themeMode  # Expected values: "dark" or "light"
		)

		# Check if TrueLaunch modification is enabled
		if (-Not $customizeTrueLaunch) {
			Write-Log "TrueLaunchBar modification is disabled in config.ps1. Skipping." -verboseMessage $true
			return
		}

		# Validate if the file exists
		if (-Not (Test-Path $trueLaunchIniFilePath)) {
			Write-Log "True Launch Bar settings file not found: $trueLaunchIniFilePath" -verboseMessage $true
			return
		}

		Write-Log "Modifying True Launch Bar settings for $themeMode theme." -verboseMessage $true
		Write-Log "Using $trueLaunchIniFilePath." -verboseMessage $true

		<# Define settings for dark and light themes
		Study TLB Setup.ini for more customizations #>
		$settingsDark = @{
			"MenuActiveColor2"      = "10053120"
			"MenuActiveColor"       = "10053120"
			"MenuActiveTextColor"   = "16777215"
			"MenuBackgroundColor"   = "2960685"
			"menuSeparatorColor1"   = "0"
			"menuSeparatorColor2"   = "6908265"
			"MenuTextColor"         = "16777215"
		}

		$settingsLight = @{
			"MenuActiveColor2"      = "-1"
			"MenuActiveColor"       = "-1"
			"MenuActiveTextColor"   = "-1"
			"MenuBackgroundColor"   = "-1"
			"menuSeparatorColor1"   = "-1"
			"menuSeparatorColor2"   = "-1"
			"MenuTextColor"         = "-1"
		}

		# Select settings based on the theme mode
		$settingsToApply = if ($themeMode -eq "dark") { $settingsDark } else { $settingsLight }

		# Read existing INI file content
		$iniContent = Get-Content -Path $trueLaunchIniFilePath -Raw
		$updatedContent = $iniContent

		# Modify settings under [settings] section
		foreach ($key in $settingsToApply.Keys) {

			$regex = "(?<=\b$key=)[\-\d]+"

			if ($updatedContent -match "\b$key=") {

				# Update existing key
				$updatedContent = $updatedContent -replace $regex, $settingsToApply[$key]
				Write-Log "Updated $key to $($settingsToApply[$key])" -verboseMessage $true

			} else {

				# If key is missing, append it (shouldn't happen, but just in case)
				$updatedContent = $updatedContent -replace "\[settings\]", "[settings]`r`n$key=$($settingsToApply[$key])"
				Write-Log "Added missing key: $key=$($settingsToApply[$key])" -verboseMessage $true
			}
		}

		# Save the updated content back to the INI file
		Set-Content -Path $trueLaunchIniFilePath -Value $updatedContent -Encoding UTF8

		Write-Log "True Launch Bar settings updated." -verboseMessage $true
	}

	# Change T-Clock Text Color
	function Set-TClockColor {
		param (
			[string]$ThemeMode  # "dark" or "light"
		)

		# Define Colors (Decimal format for DWORD)
		# White = 16777215 (0xFFFFFF) Visible on Dark Taskbar
		# Black = 0 (0x000000) Visible on Light Taskbar
		$newColor = if ($ThemeMode -eq "dark") { 16777215 } else { 0 }

		# Portable Mode Method

		if ($tClockPath -and (Test-Path $tClockPath)) {

			$iniPath = Join-Path (Split-Path -Parent $tClockPath) "T-Clock.ini"

			# if T-Clock is running from a portable folder with an INI file
			if (Test-Path $iniPath) {
				try {
					$content = Get-Content -Path $iniPath -Raw

					# Regex replace ForeColor
					if ($content -match "(?m)^ForeColor=[-0-9]+") {
						$content = $content -replace "(?m)^ForeColor=[-0-9]+", "ForeColor=$newColor"
					} else {
						$content = $content -replace "\[Clock\]", "[Clock]`r`nForeColor=$newColor"
					}

					# Regex remove BackColor (Crash prevention)
					$content = $content -replace "(?m)^BackColor=[-0-9]+\r?\n?", ""

					Set-Content -Path $iniPath -Value $content -NoNewline
					Write-Log "INI: Updated portable settings." -verboseMessage $true
				} catch {
					Write-Log "INI update failed: $_"
				}

			# If there is no portable .ini file, check registry
			} else {

				# Registry Method
				$regKey = "HKCU:\SOFTWARE\Stoic Joker's\T-Clock 2010\Clock"

				if (Test-Path $regKey) {
					try {
						# Set Text Color
						Set-ItemProperty -Path $regKey -Name "ForeColor" -Value $newColor -Type DWORD -Force
						Write-Log "Registry: Set T-Clock ForeColor to $newColor ($ThemeMode)" -verboseMessage $true

						# SAFETY MEASURE: Remove BackColor to prevent Shell Crashes
						# We silently continue if it's already gone.
						Remove-ItemProperty -Path $regKey -Name "BackColor" -ErrorAction SilentlyContinue
						Write-Log "Registry: Ensured BackColor is removed (Crash prevention)." -verboseMessage $true

					} catch {
						Write-Log "Registry update failed: $_"
					}
				}
			}
		}

		# Restart T-Clock
		# This applies the changes immediately
		if ($restartTClockColor) {

			Write-Log "Restarting T-Clock to apply new color." -verboseMessage $true
			$proc = Get-Process -Name "Clock64" -ErrorAction SilentlyContinue
			if (-not $proc) { $proc = Get-Process -Name "Clock" -ErrorAction SilentlyContinue }
			if (-not $proc) { $proc = Get-Process -Name "T-Clock" -ErrorAction SilentlyContinue }

			if ($proc) {

				Stop-Process -Id $proc.Id -Force
				Start-Sleep -Milliseconds 250

				if ($tClockPath) {

					Start-ProcessUnelevated -FilePath $tClockPath

				} else {

					# Try to restart from the path of the process we just killed
					Start-ProcessUnelevated -FilePath $proc.Path
				}

			} elseif ($tClockPath) {

				Start-ProcessUnelevated -FilePath $tClockPath
			}
		}
	}

	# Restart Sysinternals Process Explorer
	function Restart-ProcessExplorer {

		# Check if procexp.exe or procexp64.exe are running
		$proc = Get-Process | Where-Object { $_.ProcessName -match "procexp(64)?" }

		if ($proc) {

			# Retrieve the executable path safely
			$exePath = ($proc | Select-Object -First 1).Path

			if (-not $exePath) {
				Write-Log "Error: Could not retrieve Process Explorer's path." -verboseMessage $true
				return
			}

			Write-Log "Restarting Process Explorer: $exePath" -verboseMessage $true

			# Stop Process Explorer
			Stop-Process -Id $proc.Id -Force

			Start-Sleep -Seconds 2  # Ensure it has fully closed

			# Restart minimized

			if ($restartProcexpElevated) {

				Start-Process -FilePath $exePath -ArgumentList "-t" -WindowStyle Minimized
				Write-Log "Started Process Explorer with elevated rights." -verboseMessage $true

			} else {

				<# If not elevated, we close and restart via COM.
				Note: We manually handle the pathing here since Restart-ProcessExplorer uses Start-Process. #>
				$proc = Get-Process | Where-Object { $_.ProcessName -match "procexp(64)?" }
				if ($proc) {
					$exePath = ($proc | Select-Object -First 1).Path
					Stop-Process -Id $proc.Id -Force
					Start-Sleep -Seconds 2
					Start-ProcessUnelevated -FilePath $exePath -ArgumentList "-t"
				}

				Write-Log "Started Process Explorer without elevated rights." -verboseMessage $true
			}

		} else {
			Write-Log "Process Explorer is not running. No restart needed." -verboseMessage $true
		}
	}

	# Restart Windows Explorer
	function Restart-Explorer {

		Write-Log "Waiting $waitExplorer seconds before restarting Windows Explorer..." -verboseMessage $true

		# Delay so that it doesn't mess with Windows startup programs
		Start-Sleep -Seconds $waitExplorer

		Write-Log "Restarting Windows Explorer." -verboseMessage $true

		# Attempt 1: The "Polite" Request (CloseMainWindow)
        # This sends a close signal (like clicking 'X'), giving Explorer time to save state
        # and notify background tasks (preventing the surge).
        # $explorerProc = Get-Process -Name explorer -ErrorAction SilentlyContinue

        # if ($explorerProc) {
        #     foreach ($proc in $explorerProc) {
        #         # This is the magic command missing from the previous snippet
        #         $proc.CloseMainWindow() | Out-Null
        #     }

        #     # Wait up to 5 seconds for it to close gracefully
        #     $timeout = 0
        #     while ((Get-Process -Name explorer -ErrorAction SilentlyContinue) -and ($timeout -lt 5)) {
        #         Start-Sleep -Seconds 1
        #         $timeout++
        #     }
        # }

		# Attempt 2: The "Hard" Kill (Only if it stuck)
        # If it's still running after 5 seconds, it's hung, so we force kill it.
        if (Get-Process -Name explorer -ErrorAction SilentlyContinue) {
            #Write-Log "Explorer did not close gracefully. Forcing close..." -verboseMessage $true
            Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        }

        # Delay to ensure the system registers the closure
        Start-Sleep -Seconds 2

		# Delay to ensure it's fully closed
		Start-Sleep -Seconds 3

		$explorer = Get-Process | Where-Object { $_.ProcessName -eq "explorer" } -ErrorAction SilentlyContinue

		# Start if it hasn't already started (avoids new window)
		if (-Not ($explorer)) {Start-ProcessUnelevated "explorer.exe" -ErrorAction SilentlyContinue}

		Write-Log "Windows Explorer restarted." -verboseMessage $true
	}

	# Combined function to configure/restart apps
	function Update-Apps {

		param (
			[string]$themeMode  # "light" or "dark"
		)

		try {

			# extra apps
			if ($restartProcexp) {Restart-ProcessExplorer}
			if ($customizeTrueLaunch) {Update-TrueLaunch -themeMode $themeMode }
			if ($RestartMusicBee) {Restart-MusicBee}
			if ($updateTClockColor) {Set-TClockColor -ThemeMode $themeMode}

			# Restart Explorer
			Restart-Explorer

		} catch {

			Write-Log "Update-Apps error: $_" -verboseMessage $true
		}
	}

	# Function to check if the script is running as admin
	function Test-IsAdmin {

		$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
		$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
		return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
	}

	# Run script as Administrator
	function Invoke-AsAdmin {
		if (-Not (Test-IsAdmin)) {
			Write-Log "Requesting elevation. Forwarding parameters..." -verboseMessage $true

			# We use $script:PSBoundParameters to reach outside the function's scope
			$paramsToPass = $script:PSBoundParameters.Keys | ForEach-Object { "-$_" }

			$argList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"") + $paramsToPass

			try {
				# Added -WindowStyle Hidden here to keep the secondary window from "flashing"
				Start-Process -FilePath "powershell.exe" -ArgumentList $argList -Verb RunAs -WindowStyle Hidden
				exit 0
			} catch {
				Write-Log "Elevation failed: $_" -Level Error
				exit 1
			}
		}
	}

	# Return coordinates for Sunrise API (may require Internet connectivity)
	function Get-Geolocation {
        param (
            [double]$FallbackLatitude = $userLat,
            [double]$FallbackLongitude = $userLng
        )

        Write-Log "Getting location coordinates." -verboseMessage $true
        $finalCoords = @{ Latitude = 0; Longitude = 0 }

        # 1. User-defined Override
        if ($useUserLoc) {
            Write-Log "Using user-defined coordinates from config."
            $finalCoords = @{ Latitude = $FallbackLatitude; Longitude = $FallbackLongitude }
        }
        # 2. Windows Sensor
        elseif ($null -ne (Get-Service -Name "lfsvc" -ErrorAction SilentlyContinue) -and (Get-Service -Name "lfsvc").StartType -ne 'Disabled') {
            try {
                Add-Type -AssemblyName 'Windows.Devices.Geolocation'
                $geolocator = New-Object Windows.Devices.Geolocation.Geolocator
                $asyncOp = $geolocator.GetGeopositionAsync()
                $counter = 0
                while ($asyncOp.Status -eq 'Started' -and $counter -lt 50) { Start-Sleep -Milliseconds 100; $counter++ }
                if ($asyncOp.Status -eq 'Completed') {
                    $pos = $asyncOp.GetResults().Coordinate.Point.Position
                    $finalCoords = @{ Latitude = [double]$pos.Latitude; Longitude = [double]$pos.Longitude }
                }
            } catch { Write-Log "Device sensor unavailable." -verboseMessage $true }
        }

        # 3. IP Fallback if sensor failed
        if ($finalCoords.Latitude -eq 0) {
            try {
                $res = Invoke-RestMethod -Uri "http://ip-api.com/json" -TimeoutSec 5
                if ($res.status -eq "success") {
                    $finalCoords = @{ Latitude = [double]$res.lat; Longitude = [double]$res.lon }
                }
            } catch { Write-Log "Online IP location unreachable." -verboseMessage $true }
        }

        # 4. Final Default
        if ($finalCoords.Latitude -eq 0) {
            $finalCoords = @{ Latitude = $FallbackLatitude; Longitude = $FallbackLongitude }
        }

        # --- Check for possible incongruity between time zone and longitude ---
        $systemOffset = [System.TimeZoneInfo]::Local.GetUtcOffset((Get-Date)).TotalHours
        $expectedOffset = [Math]::Round($finalCoords.Longitude / 15)
        if ([Math]::Abs($systemOffset - $expectedOffset) -gt 1) {
            Write-Log "Warning: System Timezone ($systemOffset) differs from Longitude-based timezone ($expectedOffset). Check coordinates in config." -verboseMessage $true
        }

        return $finalCoords
    }

	# Calculate Sunrise and Sunset times based on coordinates and date
	function Get-SolarTimes {
		param (
			[double]$lat,
			[double]$lng,
			[DateTime]$date = (Get-Date)
		)

		# --- 1. Basic Solar Geometry ---
		$dayOfYear = $date.DayOfYear
		# Fractional year in radians
		$gamma = 2 * [Math]::PI / 365 * ($dayOfYear - 1)

		# Equation of time (minutes) - accounts for Earth's elliptical orbit
		$eqtime = 229.18 * (0.000075 + 0.001868 * [Math]::Cos($gamma) - 0.032077 * [Math]::Sin($gamma) - 0.014615 * [Math]::Cos(2 * $gamma) - 0.040849 * [Math]::Sin(2 * $gamma))

		# Solar declination (radians) - accounts for Earth's axial tilt
		$decl = 0.006918 - 0.399912 * [Math]::Cos($gamma) + 0.070257 * [Math]::Sin($gamma) - 0.006758 * [Math]::Cos(2 * $gamma) + 0.000907 * [Math]::Sin(2 * $gamma)

		# --- 2. Calculate Hour Angle ---
		# -0.833 degrees is the standard for sunrise/sunset (accounting for atmospheric refraction)
		$zenithRad = 90.833 * [Math]::PI / 180
		$latRad = $lat * [Math]::PI / 180

		$cosHA = ([Math]::Cos($zenithRad) / ([Math]::Cos($latRad) * [Math]::Cos($decl))) - ([Math]::Tan($latRad) * [Math]::Tan($decl))

		# Handle edge cases (Polar Day/Night)
		if ($cosHA -ge 1) { return @{ sunrise = $null; sunset = $null; status = "Polar Night" } }
		if ($cosHA -le -1) { return @{ sunrise = $null; sunset = $null; status = "Polar Day" } }

		$ha = [Math]::Acos($cosHA) * 180 / [Math]::PI

		# --- 3. Final UTC Calculation ---
		$sunriseUTC = 720 - 4 * ($lng + $ha) - $eqtime
		$sunsetUTC = 720 - 4 * ($lng - $ha) - $eqtime

		# --- 4. Local Conversion ---
		# We grab the current system's timezone offset automatically
		$offset = [System.TimeZoneInfo]::Local.GetUtcOffset($date).TotalHours

		return @{
			sunrise = $date.Date.AddMinutes($sunriseUTC).AddHours($offset)
			sunset  = $date.Date.AddMinutes($sunsetUTC).AddHours($offset)
			status  = "Success"
		}
	}

	# Debug function. Turn off accent color in taskbar. This is called by the Set-ThemeFile function.
	function Disable-TaskbarAccent {
		param (
			[ValidateSet("On","Off")]
			[string]$State
		)

		$value = if ($State -eq "On") { 1 } else { 0 }

		# This controls "Show accent color on Start and taskbar"
		Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "ColorPrevalence" -Value $value -Force
	}

	# Sets the system color mode directly via registry and refreshes the UI
	function Set-ThemeMode {

        param ([string]$Mode)

        $value = if ($Mode -eq "light") { 1 } else { 0 }
        $path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"

        Write-Log "Setting native $Mode mode via registry." -verboseMessage $true

		# Applying Registry values
        Set-ItemProperty -Path $path -Name "SystemUsesLightTheme" -Value $value -Force | Out-Null
        Set-ItemProperty -Path $path -Name "AppsUseLightTheme" -Value $value -Force | Out-Null

		# Telling Windows to make the changes
        Write-Log "Broadcasting WM_SETTINGCHANGE to all windows." -verboseMessage $true
        $result = [IntPtr]::Zero

        # Catch the return value in $null to prevent "1" from appearing in logs
        $null = [WinAPI]::SendMessageTimeout(0xffff, 0x001A, [IntPtr]::Zero, "ImmersiveColorSet", 0x0002, 5000, [ref]$result)
    }

	# Run the .theme file
	function Set-ThemeFile {
		param (
			[string]$ThemePath
			)

		# Check if the theme file exists
		if (Test-Path $ThemePath) {

			Write-Log "Activating the .theme file" -verboseMessage $true

			# Apply the theme; equivalent of double-clicking on a .theme file
			Start-Process $ThemePath

			# Wait a bit for the theme to apply and the Settings window to appear
			Start-Sleep -Seconds 4

			# Turn off or on the accent color in taskbar usage. Debug feature.
			if ($turnOffAccentColor) { Disable-TaskbarAccent -State "Off" } else { Disable-TaskbarAccent -State "On" }

			Write-Log "Closing the Settings window" -verboseMessage $true

			# Close the Settings window by stopping the "ApplicationFrameHost" process
			$settingsProcess = Get-Process -Name "ApplicationFrameHost" -ErrorAction SilentlyContinue
			if ($settingsProcess) {
				Stop-Process -Id $settingsProcess.Id
			}

		} else {

			Write-Log "Theme file not found: $ThemePath"
		}
	}

	# Check if a Scheduled task exists
	function Test-IfTaskExists {
		param (
			[string]$taskName
		)

		$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
		return $null -ne $task
	}

	<# Create a scheduled task for next daylight events
    This task will be created or overwritten by the main task. #>
	function Register-Task {
        param (
            [DateTime]$NextTriggerTime,
            [String]$Name,
            [String]$Description = "Managed by the Auto Theme script.",
            [String]$CustomFile = $PSCommandPath, # Default to at.ps1
            [String]$Arguments = ""               # Optional extra arguments
        )

        # Schedule next run
        Write-Log "Setting scheduled task: $Name"

        # Define base PowerShell arguments
        $psArgs = "-WindowStyle Hidden -ExecutionPolicy Bypass -NoProfile -NonInteractive -File `"$CustomFile`" $Arguments"

        # Logic to determine the executable and arguments based on $terminalVisibility
        switch ($terminalVisibility) {
            "ch" {
                $exe = "conhost.exe"
                $arguments = "--headless PowerShell.exe $psArgs"
            }
            "wt" {
                if (Get-Command "wt.exe" -ErrorAction SilentlyContinue) {
                    $exe = "wt.exe"
                    $arguments = "-w 0 nt PowerShell.exe -NoLogo $psArgs"
                } else {
                    $exe = "PowerShell.exe"
                    $arguments = $psArgs
                }
            }
            Default {
                $exe = "PowerShell.exe"
                $arguments = $psArgs
            }
        }

        # Create Action and Trigger
        $action = New-ScheduledTaskAction -Execute $exe -Argument $arguments

        if ($useFixedHours -and ($Name -eq "Auto Theme fixed sunrise" -or $Name -eq "Auto Theme fixed sunset")) {
            $timeOfDay = $NextTriggerTime.ToString("HH:mm")
            $trigger = New-ScheduledTaskTrigger -Daily -At $timeOfDay
        } else {
            $trigger = New-ScheduledTaskTrigger -Once -At $NextTriggerTime
        }

        # Define Principal and Settings
        $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        $userSid = $currentUser.User.Value
        $principal = New-ScheduledTaskPrincipal -UserId $userSid -LogonType Interactive -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -Compatibility Win8 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

        # Build the Task Object with the Description
        $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description $Description

        # Unregister old and register new
        if (Test-IfTaskExists $Name) {
            Unregister-ScheduledTask -TaskName $Name -Confirm:$false
        }

        try {
            Register-ScheduledTask -TaskName $Name -InputObject $task | Out-Null
            Write-Log "Registered new task: $Name"
        } catch {
            Write-Log "Error registering task: $_"
        }
    }

# ============= MAIN FUNCTIONS  ==============

	<# Toggle the theme when running from Terminal
	using the command `./at.ps1` #>
	function Switch-Theme {

		# Determine current state
        if ($useThemeFiles) {

			# Check the active .theme file name
            $CurrentTheme = (Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes" -Name CurrentTheme).CurrentTheme
            $isCurrentlyDark = $CurrentTheme -match "dark"

        } else {

            # Check the "AppsUseLightTheme" value (0 = Dark, 1 = Light)
            $RegPath = "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            $isCurrentlyDark = (Get-ItemProperty -Path $RegPath -Name "AppsUseLightTheme").AppsUseLightTheme -eq 0

        }

        # Toggle labels
        if ($isCurrentlyDark) {

            $mode = "light"
            $label = $themeLight # Or "Light Mode" if $themeLight is null
            $targetPath = $lightPath
            $wallPath = $wallLightPath

        } else {

            $mode = "dark"
            $label = $themeDark
            $targetPath = $darkPath
            $wallPath = $wallDarkPath

        }

		# Apply the change
        if ($useThemeFiles) {

            Write-Log "Selected $label" -verboseMessage $true

			# Set the wallpaper, grab its name
			$selectedWall = Get-WallpaperName -wallpaperDirectory $wallPath -themeFilePath $targetPath

            Set-ThemeFile $targetPath

        } else {

            Write-Log "Applying $mode mode" -verboseMessage $true

			# Set the wallpaper, grab its name
			$selectedWall = & $workerPath -Path $wallPath -InformationAction Continue

            Set-ThemeMode -Mode $Mode

			# Start rotating wallpapers via task
			Set-WallpaperRotationTask -Mode $mode
        }

        # Update programs
        Update-Apps -themeMode $mode

		# Send notification
        Send-ThemeNotification -MainLine "Theme toggled. $label activated." # -SelectedWallpaper $selectedWall

    }

	<# Calculate daylight events or pick fixed hours
    then select the Theme depending on daylight #>
    function Invoke-ThemeScheduling {

        $Now = Get-Date
        $NowDate = $Now.ToString("yyyy-MM-dd")

        # Picked times mode
        if ($useFixedHours) {

            # Parse fixed times as proper DateTime objects
            try {

                # Try parsing with various common formats
                if ($lightThemeTime -match '^\d{1,2}:\d{2}$') {
                    # 24-hour format (e.g., "07:00")
                    $timeComponents = $lightThemeTime.Split(':')
                    $Sunrise = Get-Date -Hour ([int]$timeComponents[0]) -Minute ([int]$timeComponents[1]) -Second 0
                }

                elseif ($lightThemeTime -match '^\d{1,2}:\d{2}\s*[AaPp][Mm]$') {
                    # 12-hour format with AM/PM (e.g., "7:00 AM")
                    $Sunrise = [DateTime]::ParseExact($lightThemeTime.Trim(), "h:mm tt", [System.Globalization.CultureInfo]::InvariantCulture)
                    $Sunrise = Get-Date -Year $Now.Year -Month $Now.Month -Day $Now.Day -Hour $Sunrise.Hour -Minute $Sunrise.Minute -Second 0
                }

                else {
                    # General fallback parsing
                    $Sunrise = [DateTime]::Parse($lightThemeTime)
                    $Sunrise = Get-Date -Year $Now.Year -Month $Now.Month -Day $Now.Day -Hour $Sunrise.Hour -Minute $Sunrise.Minute -Second 0
                }

                # Same for sunset time
                if ($darkThemeTime -match '^\d{1,2}:\d{2}$') {
                    $timeComponents = $darkThemeTime.Split(':')
                    $Sunset = Get-Date -Hour ([int]$timeComponents[0]) -Minute ([int]$timeComponents[1]) -Second 0
                }

                elseif ($darkThemeTime -match '^\d{1,2}:\d{2}\s*[AaPp][Mm]$') {
                    $Sunset = [DateTime]::ParseExact($darkThemeTime.Trim(), "h:mm tt", [System.Globalization.CultureInfo]::InvariantCulture)
                    $Sunset = Get-Date -Year $Now.Year -Month $Now.Month -Day $Now.Day -Hour $Sunset.Hour -Minute $Sunset.Minute -Second 0
                }

                else {
                    $Sunset = [DateTime]::Parse($darkThemeTime)
                    $Sunset = Get-Date -Year $Now.Year -Month $Now.Month -Day $Now.Day -Hour $Sunset.Hour -Minute $Sunset.Minute -Second 0
                }

                # Set tomorrow's sunrise for overnight calculations
                $TomorrowSunrise = $Sunrise.AddDays(1)

                Write-Log "Successfully parsed fixed times: Sunrise at $($Sunrise.ToString('HH:mm')), Sunset at $($Sunset.ToString('HH:mm'))" -verboseMessage $true
            }

            catch {

                Write-Log "Error parsing time strings. Using default values." -verboseMessage $true

                # Default fallback times if parsing fails
                $Sunrise = Get-Date -Hour 7 -Minute 0 -Second 0
                $Sunset = Get-Date -Hour 19 -Minute 0 -Second 0
                $TomorrowSunrise = $Sunrise.AddDays(1)
            }

            # In fixed hours mode, we use differently named tasks
            $SunriseTaskName = "Auto Theme fixed sunrise"
            $SunsetTaskName = "Auto Theme fixed sunset"

            # Clean up dynamic tasks if they exist
            if (Test-IfTaskExists "Auto Theme sunrise") {

                Unregister-ScheduledTask -TaskName "Auto Theme sunrise" -Confirm:$false
                Write-Log "Removed dynamic sunrise task as we're using fixed hours" -verboseMessage $true
            }
            if (Test-IfTaskExists "Auto Theme sunset") {

                Unregister-ScheduledTask -TaskName "Auto Theme sunset" -Confirm:$false
                Write-Log "Removed dynamic sunset task as we're using fixed hours" -verboseMessage $true
            }

            # Check if fixed tasks already exist - if so, we don't need to recreate them
            $sunriseTaskExists = Test-IfTaskExists $SunriseTaskName
            $sunsetTaskExists = Test-IfTaskExists $SunsetTaskName

            # Only create fixed tasks if they don't already exist
            if (-not $sunriseTaskExists) {

                Register-Task -NextTriggerTime $Sunrise -Name $SunriseTaskName -Description "Triggers transition to light mode based on a fixed schedule. Managed by the main Auto Theme task."
                Write-Log "Created fixed sunrise task for daily operation"
            }

            if (-not $sunsetTaskExists) {

                Register-Task -NextTriggerTime $Sunset -Name $SunsetTaskName -Description "Triggers transition to dark mode based on a fixed schedule. Managed by the main Auto Theme task."
                Write-Log "Created fixed sunset task for daily operation"
            }

        # Dynamic hours mode
        } else {

            # Remove fixed tasks if they exist
            if (Test-IfTaskExists "Auto Theme fixed sunrise") {

                Unregister-ScheduledTask -TaskName "Auto Theme fixed sunrise" -Confirm:$false
                Write-Log "Removed fixed sunrise task as we're using dynamic times" -verboseMessage $true
            }

            if (Test-IfTaskExists "Auto Theme fixed sunset") {
                Unregister-ScheduledTask -TaskName "Auto Theme fixed sunset" -Confirm:$false
                Write-Log "Removed fixed sunset task as we're using dynamic times" -verboseMessage $true
            }

            # Dynamic times mode
            $location = Get-Geolocation
            $lat = $location.Latitude
            $lng = $location.Longitude

            # we use online API to retrieve sunrise/sunset times
            if ($useAPIdaylight){

                Write-Log "Using online API to get daylight times." -verboseMessage $true

                # Extract latitude, longitude and timezone for API call
                $tzid = $location.Timezone

                # Either API can be used, but the first one may have faulty control over DateTime formatting
                # $APIurl1 = "https://api.sunrise-sunset.org/json?lat=$lat&lng=$lng&date=$NowDate&tzid=$tzid"
                $APIurl2 = "https://api.sunrisesunset.io/json?lat=$lat&lng=$lng&date=$NowDate&timezone=$tzid"
                $url = $APIurl2
                Write-Log "Using this API call = $url" -verboseMessage $true

                try {
                    $Daylight = (Invoke-RestMethod $url).results
                    Write-Log "Fetched daylight data string = $Daylight" -verboseMessage $true

                    # Parse and adjust the dates
                    $SunriseTimeString = $Daylight.sunrise
                    $SunriseDateString = $Daylight.date
                    $SunriseString = "$SunriseTimeString $SunriseDateString"
                    $Sunrise = [DateTime]::ParseExact($SunriseString, "h:mm:ss tt yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)

                    $SunsetTimeString = $Daylight.sunset
                    $SunsetDateString = $Daylight.date
                    $SunsetString = "$SunsetTimeString $SunsetDateString"
                    $Sunset = [DateTime]::ParseExact($SunsetString, "h:mm:ss tt yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)

                    # second query for tomorrow
                    $TomorrowDaylight = (Invoke-RestMethod "$url&date=tomorrow").results

                    $TomorrowSunriseTimeString = $TomorrowDaylight.sunrise
                    $TomorrowSunriseDateString = $TomorrowDaylight.date
                    $TomorrowSunriseString =  "$TomorrowSunriseTimeString $TomorrowSunriseDateString"
                    $TomorrowSunrise = [DateTime]::ParseExact($TomorrowSunriseString, "h:mm:ss tt yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
                } catch {
                    Write-Log "Online API failed. Falling back to internal calculation." -Level Warning
                    $useAPIdaylight = $false # Trigger internal calculation fallback
                }
            }

            # we use our own internal calculation to determine sunrise/sunset times
            if (-not $useAPIdaylight) {

                Write-Log "Using internal calculation to get daylight times." -verboseMessage $true

                $solar = Get-SolarTimes -lat $lat -lng $lng -date $Now

                # Handle Polar/Standard results
                if ($solar.status -eq "Polar Day") {
                    $Sunrise = $Now.Date; $Sunset = $Now.Date.AddDays(1).AddSeconds(-1)
                } elseif ($solar.status -eq "Polar Night") {
                    $Sunrise = $Now.Date.AddDays(1); $Sunset = $Now.Date.AddDays(-1)
                } else {
                    $Sunrise = $solar.sunrise
                    $Sunset  = $solar.sunset
                }

                # Get tomorrow for the next trigger
                $tomorrowSolar = Get-SolarTimes -lat $lat -lng $lng -date ($Now.AddDays(1))
                $TomorrowSunrise = if ($tomorrowSolar.status -eq "Success") { $tomorrowSolar.sunrise } else { $Sunrise.AddDays(1) }
            }

            # Apply offsets if defined and non-zero
            if ($null -ne $sunriseOffset -and $sunriseOffset -ne 0) {
                $Sunrise = $Sunrise.AddMinutes([int]$sunriseOffset)
                $TomorrowSunrise = $TomorrowSunrise.AddMinutes([int]$sunriseOffset)
                Write-Log "Adding sunrise offset of $sunriseOffset minutes." -verboseMessage $true
            }

            if ($null -ne $sunsetOffset -and $sunsetOffset -ne 0) {
                $Sunset = $Sunset.AddMinutes([int]$sunsetOffset)
                Write-Log "Adding sunset offset of $sunsetOffset minutes." -verboseMessage $true
            }

            Write-Log "Using dynamic hours: Sunrise at $Sunrise, Sunset at $Sunset, TomorrowSunrise at $TomorrowSunrise" -verboseMessage $true

            # In dynamic mode, we use standard task names
            $SunriseTaskName = "Auto Theme sunrise"
            $SunsetTaskName = "Auto Theme sunset"
        }

        # Determine current theme/mode state based on configuration
        if ($useThemeFiles) {

            $CurrentTheme = (Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes" -Name CurrentTheme).CurrentTheme

        } else {

            $RegPath = "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            $NativeValue = (Get-ItemProperty -Path $RegPath -Name "AppsUseLightTheme").AppsUseLightTheme
        }

        # Determine if we need to change the theme based on current time
		# It's daytime - light theme period
        if ($Now -ge $Sunrise -and $Now -lt $Sunset) {

            $NextTriggerTime = $Sunset
            $NextTaskName = $SunsetTaskName
            $mode = "light"

            # Check if mode is already correct
			# if using theme files
            if ($useThemeFiles) {

				# is the theme already active?
                if ($CurrentTheme -match $themeLight) {

					$mainLine =  "$themeLight is already active. Next trigger at: $NextTriggerTime"

                } else {

	                Write-Log "Setting the theme $themeLight" -verboseMessage $true

	                # Change the theme
	                Set-ThemeFile -ThemePath $lightPath

	                # line for log and notification
	                $mainLine = "$themeLight activated. Next trigger at: $NextTriggerTime"

				}

                # Set the wallpaper, grab its name
                $selectedWall = Get-WallpaperName -wallpaperDirectory $wallLightPath -themeFilePath $lightPath

			# if not using theme files
            } else {

                if ($NativeValue -eq 1) {

					$mainLine =  "$mode mode is already active. Next trigger at: $NextTriggerTime"

                } else {

	                Write-Log "Applying $mode mode" -verboseMessage $true

	                # Change the mode
	                Set-ThemeMode -Mode $Mode

	                # line for log and notification
	                $mainLine = "Activated $mode mode. Next trigger at: $NextTriggerTime"

				}

                # Set the wallpaper, grab its name
                $selectedWall = & $workerPath -Path $wallLightPath -InformationAction Continue

                # Start rotating wallpapers via task
                Set-WallpaperRotationTask -Mode $mode
            }

            Update-Apps -themeMode $mode
            Write-Log $mainLine

            # Send notification
            Send-ThemeNotification -MainLine $mainLine # -SelectedWallpaper $selectedWall

            # Register the task
            if (-not $useFixedHours) {

                Register-Task -NextTriggerTime $NextTriggerTime -Name $NextTaskName -Description "Triggers transition to dark mode based on daylight. Managed by the main Auto Theme task."
            }

        # if it's nighttime - dark theme period
		} else {

            $NextTriggerTime = if ($Now -ge $Sunset) { $TomorrowSunrise } else { $Sunrise }
            $NextTaskName = $SunriseTaskName
            $mode = "dark"

            # Check if mode is already correct
			# if using theme files
            if ($useThemeFiles) {

				# is the theme already active?
                if ($CurrentTheme -match $themeDark) {

                    $mainLine =  "$themeDark is already active. Next trigger at: $NextTriggerTime"

                } else {

	                Write-Log "Setting the theme $themeDark" -verboseMessage $true

	                # Change the theme
	                Set-ThemeFile -ThemePath $darkPath

	                # line for log and notification
	                $mainLine = "$themeDark activated. Next trigger at: $NextTriggerTime"

                }

                # Set the wallpaper, grab its name
                $selectedWall = Get-WallpaperName -wallpaperDirectory $wallDarkPath -themeFilePath $darkPath

			# if not using theme files
            } else {

                if ($NativeValue -eq 0) {

                    $mainLine =  "$mode mode is already active. Next trigger at: $NextTriggerTime"

                } else {

	                Write-Log "Applying $mode mode" -verboseMessage $true

	                # Change the mode
	                Set-ThemeMode -Mode $Mode

	                # line for log and notification
	                $mainLine = "Activated $mode mode. Next trigger at: $NextTriggerTime"

				}

                # Set the wallpaper, grab its name
                $selectedWall = & $workerPath -Path $wallDarkPath -InformationAction Continue

                # Start rotating wallpapers via task
                Set-WallpaperRotationTask -Mode $mode           
			}

            Update-Apps -themeMode $mode
            Write-Log $mainLine

            # Send notification
            Send-ThemeNotification -MainLine $mainLine # -SelectedWallpaper $selectedWall

            # Register the task
            if (-not $useFixedHours) {

                Register-Task -NextTriggerTime $NextTriggerTime -Name $NextTaskName -Description "Triggers transition to light mode based on daylight. Managed by the main Auto Theme task."
            }
        }
    }

# ============= RUNTIME ==============

try {

    # Load configuration
    if (-Not (Test-Path $ConfigPath)) {
        Write-Error "Configuration file not found: $ConfigPath"
        Exit 1
    }
    . $ConfigPath

    # Cleanup log file
    if ($trimLog -and -Not (Test-TerminalSession)) {
        Limit-LogHistory -logFilePath $logFile -maxSessions $maxLogEntries
    }

	# Start logging
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Log ""
    Write-Log "$timestamp === Auto-Theme script started (Version: $scriptVersion)"


	# Failsafe to avoid running the script too often
    if ($checkLastRun) { Test-LastRunTime }
    Update-LastRunTime

	# ======= Process script parameters =======

	# 1. Only Wallpaper changing
    if ($Next) {

        Write-Log "Loading 'Next Wallpaper'." -verboseMessage $true

        $RegPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $isLight = (Get-ItemProperty -Path $RegPath -Name "AppsUseLightTheme").AppsUseLightTheme -eq 1
        $mode = if ($isLight) { "light" } else { "dark" }
        $wallPath = if ($isLight) { $wallLightPath } else { $wallDarkPath }

        & $workerPath -Path $wallPath -InformationAction Continue

		# we can reset the task to restart the timer only if admin
        if (Test-IsAdmin) {
            Set-WallpaperRotationTask -Mode $mode
            Write-Log "Rotation timer reset successfully." -verboseMessage $true
        } else {
            Write-Log "Skipped timer reset (Admin rights required)." -verboseMessage $true
        }

		exit 0
	}

    # Make sure we are Admin
    if ($forceAsAdmin) {
        Write-Log "Running as Administrator." -verboseMessage $true

        Invoke-AsAdmin
    }

	# Restart Theme service
    if ($restartThemeService) { Restart-ThemeService }

	# 2. Toggle theme or color mode
    if ($Toggle) {

		Write-Log "Toggling Theme, regardless of daylight."

        Switch-Theme

	# 3. Scheduling Theme
    } elseif ($Schedule) {

        Write-Log "Script is running from Task Scheduler." -verboseMessage $true
		Write-Log "Selecting and scheduling Theme based on daylight."

        Invoke-ThemeScheduling

	# Fallback to Process Detection, if no parameters
    } else {

        Write-Log "No parameters detected. Using parent process detection." -verboseMessage $true

        if (Test-TerminalSession) {

            Write-Log "Script is running from Terminal." -verboseMessage $true
            Write-Log "Toggling Theme, regardless of daylight."

            Switch-Theme

        } else {

			Write-Log "Script is running from Task Scheduler." -verboseMessage $true
            Write-Log "Selecting and scheduling Theme based on daylight."

            Invoke-ThemeScheduling
        }
    }

    Write-Log "=== All done." -verboseMessage $true
    Write-Log "" -verboseMessage $true

} catch {

    Write-Log "Error: $_"
}