<#
.SYNOPSIS
	Changes the active Windows theme based on a predefined/daylight schedule. Works in Windows 10/11.

.DESCRIPTION
	This Powershell script automatically switches the Windows Theme depending on Sunrise and Sunset, or hours set by the user.
	Rather than relying on registry/system settings, it works by activating given `.theme` files. 
	This allows for a much higher degree of customization and compatibility.
	The script is designed to run in the background as a scheduled task, ensuring that the system theme is updated without user intervention. 
	It will automatically create the next temporary task for the next daylight event. 
	Such tasks ("Sunrise theme" and "Sunset theme") will be overwritten as a matter of course to avoid clutter.
	It only connects to the internet to verify Location and Sunrise and Sunset times.
	Alternatively, it can stay completely offline operating on fixed hours provided by the user.
	When ran as the command `./AutoTheme.ps1` from terminal or desktop shortcut, the script will only toggle between themes.
	IMPORTANT: Edit Config.ps1 to configure this script. The file contains all necessary explanations.
	OPTIONALLY: Run Setup.ps1 to create the main Scheduled Task, or create one in Task Scheduler.
	For more information, refer to the README file, on Github.

.LINK
	https://github.com/unalignedcoder/auto-theme/

.NOTES
	- Improved the release-creation system
	- Minor fixes

#>

# Script version. This is automatically updated via pre-commit hook
$scriptVersion = "1.0.17"

# ============= Config file ==============

	$ConfigPath = Join-Path $PSScriptRoot "Config.ps1"

# ============= FUNCTIONS  ==============

	# Determine if the script runs interactively
	function IsRunningFromTerminal {

		# Get the current process ID
		$proc = Get-CimInstance Win32_Process -Filter "ProcessId = $pid"

		# Get the parent process ID
		$parent = Get-CimInstance Win32_Process -Filter "ProcessId = $($proc.ParentProcessId)"

		# Check if the parent process is 'svchost.exe' and contains 'Schedule' in the command line
		if ($parent.Name -eq "svchost.exe" -or $parent.CommandLine -like "*Schedule*") {

			return $false

		} else {

			return $true
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

				# Display log output depending on session type
				if (IsRunningFromTerminal) {

					Write-Output "$message" # Output to console
					if ($logFromTerminal) {Add-Content -Path $logFile -Value "$message" }  # Log to file
					# 

				} else {

					Add-Content -Path $logFile -Value "$message"  # Log to file

				}
			}

		} catch {

			Write-Output "Error in LogThis: $_"
		}		
	}
	
	# Trim old log entries
	function TrimOldLog {
		param (
			[string]$logFilePath,    # Path to the log file
			[int]$maxSessions = 10  # Maximum number of log sessions to keep
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
			if ($logLines[$i] -match '=== Script started \(Version: .*?\)') {
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
	function ShowBurntToast {
		param(
			[string]$Text,
			[string]$AppLogo
		)

		# Install the BurntToast module if not already installed
		if (-not (Get-Module -Name BurntToast -ListAvailable)) {

			try {

				LogThis "Installing the BurnToast Notifications module"
				Install-Module -Name BurntToast -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -Confirm:$false

			} catch {

				LogThis "Failed to install BurntToast module: $_"
				return
			}
		}

		# for when the above is commented out or fails, we double-check.
		if (Get-Module -Name BurntToast -ListAvailable) {

			try {
				
				LogThis "Creating BurnToast notification"  -verboseMessage $true
				New-BurntToastNotification -Text $Text -AppLogo $AppLogo

				LogThis "Displayed BurntToast notification with text: $Text"  -verboseMessage $true

			} catch {

				LogThis "Error displaying BurntToast notification: $_"  -verboseMessage $true
			}

		} else {

			LogThis "BurntToast module is not installed. Cannot display system notifications."
		}
	}

	# Check if the script has been run in the last hour
	function LastTime {

		LogThis "Checking if script was run in the last $interval minutes"  -verboseMessage $true

		if (Test-Path $lastRunFile) {

			$lastRun = Get-Content $lastRunFile | Out-String
			$lastRun = [DateTime]::Parse($lastRun)
			$now = Get-Date

			$timeSinceLastRun = $now - $lastRun

			if ($timeSinceLastRun.TotalMinutes -lt $interval) {

				LogThis "Script was run within the last $interval minutes. Exiting."
				exit
			}
		}
	}

	# Update the last run time
	function UpdateTime {

		$now = Get-Date
		$now | Out-File -FilePath $lastRunFile -Force
	}

	# Return coordinates for Sunrise API (may require Internet connectivity)
	function LocateThis {
		param (
			[double]$FallbackLatitude = $userLat,
			[double]$FallbackLongitude = $userLng,
			[string]$FallbackTimezone = $UserTzid
		)

		LogThis "Getting location coordinates." -verboseMessage $true

		# If $useUserLoc is set to true, return user-defined coordinates and timezone
		if ($useUserLoc) {

			LogThis "Using user-defined coordinates and timezone."

			return @{
				Latitude = $FallbackLatitude
				Longitude = $FallbackLongitude
				Timezone = $FallbackTimezone
			}
		}

		# Attempt to get location and timezone from Device Geolocation
		try {

			Add-Type -AssemblyName 'Windows.Devices.Geolocation'
			$geolocator = New-Object Windows.Devices.Geolocation.Geolocator
			$position = $geolocator.GetGeopositionAsync().GetAwaiter().GetResult()
			$userLat = $position.Coordinate.Point.Position.Latitude
			$longitude = $position.Coordinate.Point.Position.Longitude
			$UserTzid = [System.TimeZone]::CurrentTimeZone.StandardName

			LogThis "Retrieved device location and system timezone." -verboseMessage $true

			return @{
				Latitude = [double]$userLat
				Longitude = [double]$longitude
				Timezone = $UserTzid
			}
		}
		catch {

			LogThis "Device location and timezone retrieval failed. Trying online service." -verboseMessage $true
		}

		# Attempt to get location and timezone from online service
		try {

			$response = Invoke-RestMethod -Uri "http://ip-api.com/json"
			if ($response.status -eq "success") {

				LogThis "Retrieved location and timezone from online service."

				return @{
					Latitude = [double]$response.lat
					Longitude = [double]$response.lon
					Timezone = $response.timezone
				}
			}
		}
		catch {

			LogThis "Online service location and timezone retrieval failed. Using fallback." -verboseMessage $true
		}

		# Fallback to user-defined coordinates and timezone if all else fails

		LogThis "Using user-defined coordinates and timezone."

		return @{
			Latitude = $FallbackLatitude
			Longitude = $FallbackLongitude
			Timezone = $FallbackTimezone
		}
	}

	# Check whether shuffle is enabled in the .theme file
	function DoWeShuffle {
		param (
			[string]$themeFilePath
		)

		LogThis "Checking if the theme shuffles wallpapers" -verboseMessage $true

		# Read the content of the theme file
		$themeContent = Get-Content -Path $themeFilePath

		# Flag to indicate if we are inside the [Slideshow] section
		$inSlideshowSection = $false

		foreach ($line in $themeContent) {
			# LogThis "Processing line: $line" -verboseMessage $true

			# Check for the start of the [Slideshow] section
			if ($line -match '^\[Slideshow\]') {
				LogThis "Found Slideshow section" -verboseMessage $true
				$inSlideshowSection = $true
				continue
			}

			# If we are inside the [Slideshow] section, look for 'shuffle' setting
			if ($inSlideshowSection) {
				if ($line -match '(?i)shuffle=(\d)') { # Case-insensitive match for 'shuffle'
					LogThis "Found shuffle setting: $line" -verboseMessage $true
					return $matches[1] -eq '1'
				}

				# If we encounter the next section or end of file, break out of the loop
				if ($line -match '^\[.*\]') {
					LogThis "Leaving [Slideshow] section" -verboseMessage $true
					break
				}
			}
		}

		# If no shuffle setting is found, return false
		LogThis "No, the theme does not shuffle wallpapers" -verboseMessage $true
		return $false
	}

	<# Prepend the substring '_0_AutoTheme_' to one randomly chosen
	wallpaper filename, so as to make it first pick. #>
	function RandomFirstWall {
		param (
			[string]$wallpaperDirectory
		)

		if (-Not ($randomFirst)) {
			LogThis "The first wallpaper will not be randomized."  -verboseMessage $true
			return
		}

		LogThis "Randomizing first wallpaper."  -verboseMessage $true
		LogThis "Looking in $wallpaperDirectory"  -verboseMessage $true

		# Retrieve all wallpaper files
		$wallpapers = Get-ChildItem -Path $wallpaperDirectory -File

		# Ensure there are wallpapers available
		if ($wallpapers.Count -eq 0) {
			LogThis "No wallpapers available." -verboseMessage $true
			return
		}

		# Find all wallpapers that have '_0_AutoTheme_' prefix
		$existingRenamedWallpapers = $wallpapers | Where-Object { $_.Name -match '^_0_AutoTheme_' }

		# If any exist, rename them back to their original names
		if ($existingRenamedWallpapers.Count -gt 0) {

			foreach ($wallpaper in $existingRenamedWallpapers) {

				$originalName = $wallpaper.Name -replace '^_0_AutoTheme_', ''
				$originalNameFull = Join-Path $wallpaperDirectory $originalName
				
				LogThis "Restoring original name: $($wallpaper.FullName) → $originalNameFull" -verboseMessage $true
				Rename-Item -Path $wallpaper.FullName -NewName $originalNameFull -Force
			}
		}

		# Refresh the list of wallpapers after renaming
		$wallpapers = Get-ChildItem -Path $wallpaperDirectory -File

		# Select a random wallpaper
		$randomFirstWallpaper = $wallpapers | Get-Random

		# Rename it with "_0_AutoTheme_" prefix
		$newWallpaperName = "_0_AutoTheme_" + $randomFirstWallpaper.Name
		$newWallpaperNameFull = Join-Path $wallpaperDirectory $newWallpaperName
		Rename-Item -Path $randomFirstWallpaper.FullName -NewName $newWallpaperNameFull -Force

		LogThis "Renamed $($randomFirstWallpaper.FullName) to $newWallpaperNameFull" -verboseMessage $true
	}

	# Check if a Scheduled task exists
	function TaskExists {
		param (
			[string]$taskName
		)
		
		$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
		return $null -ne $task
	}

	# Modify TrueLaunchBar default colors
	function UpdateTrueLaunch {

		param (
			[string]$themeMode  # Expected values: "dark" or "light"
		)

		# Check if TrueLaunch modification is enabled
		if (-Not $customizeTrueLaunch) {
			LogThis "TrueLaunchBar modification is disabled in config.ps1. Skipping." -verboseMessage $true
			return
		}

		# Validate if the file exists
		if (-Not (Test-Path $trueLaunchIniFilePath)) {
			LogThis "True Launch Bar settings file not found: $trueLaunchIniFilePath" -verboseMessage $true
			return
		}

		LogThis "Modifying True Launch Bar settings for $themeMode theme." -verboseMessage $true
		LogThis "Using $trueLaunchIniFilePath." -verboseMessage $true

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
				LogThis "Updated $key to $($settingsToApply[$key])" -verboseMessage $true

			} else {

				# If key is missing, append it (shouldn't happen, but just in case)
				$updatedContent = $updatedContent -replace "\[settings\]", "[settings]`r`n$key=$($settingsToApply[$key])"
				LogThis "Added missing key: $key=$($settingsToApply[$key])" -verboseMessage $true
			}
		}

		# Save the updated content back to the INI file
		Set-Content -Path $trueLaunchIniFilePath -Value $updatedContent -Encoding UTF8

		LogThis "True Launch Bar settings updated." -verboseMessage $true

		# Restart Explorer
		RestartExplorer
	}

	# Restart the 'Themes' Service
	function RestartThemeService {
		
		[bool]$IsAdmin = IsAdmin
		if ($restartThemeService -and $IsAdmin) {

			try {

				LogThis "Restarting the Themes service." -verboseMessage $true

				Restart-Service -Name "Themes" -Force -ErrorAction SilentlyContinue

				LogThis "Themes service restarted successfully." -verboseMessage $true

			} catch {

				LogThis "Failed to restart Themes service: $_"  -verboseMessage $true
			}
		}
	}

	# Restart Sysinternals Process Explorer
	function RestartProcessExplorer {

		# Check if procexp.exe or procexp64.exe are running
		$proc = Get-Process | Where-Object { $_.ProcessName -match "procexp(64)?" }

		if ($proc) {

			# Retrieve the executable path safely
			$exePath = ($proc | Select-Object -First 1).Path

			if (-not $exePath) {
				LogThis "Error: Could not retrieve Process Explorer's path." -verboseMessage $true
				return
			}

			LogThis "Restarting Process Explorer: $exePath" -verboseMessage $true

			# Stop Process Explorer
			Stop-Process -Id $proc.Id -Force

			Start-Sleep -Seconds 2  # Ensure it has fully closed

			# Restart minimized
			Start-Process -FilePath $exePath -ArgumentList "-t" -WindowStyle Minimized


		} else {
			LogThis "Process Explorer is not running. No restart needed." -verboseMessage $true
		}
	}

	# Restart Windows Explorer
	function RestartExplorer {

		LogThis "Waiting a few seconds before restarting Windows Explorer..." -verboseMessage $true

		# Delay so that it doesn't mess with Windows startup programs
		Start-Sleep -Seconds $waitExplorer

		LogThis "Restarting Windows Explorer." -verboseMessage $true

		# Stop all explorer instances
		Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue

		# Delay to ensure it's fully closed
		Start-Sleep -Seconds 3  
		
		$explorer = Get-Process | Where-Object { $_.ProcessName -eq "explorer" } -ErrorAction SilentlyContinue

		# Start if it hasn't already started (avoids new window)
		if (-Not ($explorer)) {Start-Process "explorer.exe" -ErrorAction SilentlyContinue}
		
		LogThis "Windows Explorer restarted." -verboseMessage $true
	}

	# Function to check if the script is running as admin
	function IsAdmin {

		$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
		$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
		return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
	}

	# Run script as Administrator
	function RunAsAdmin {

		# Skip elevation if running as SYSTEM user
		if ($env:SYSTEMROOT -and $env:USERNAME -eq "SYSTEM") {
			LogThis "Running as SYSTEM via Task Scheduler. Skipping elevation check." -verboseMessage $true
			return
		}

		# Relaunch script as admin if not already running as admin
		if (-Not (IsAdmin)) {

			Write-Host "This script requires administrative privileges. Requesting elevation..." -ForegroundColor Yellow
			Start-Process -FilePath "powershell.exe" `
						  -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
						  -Verb RunAs
			Exit 0
		}
	}

	# Run the .theme file
	function StartTheme {
		param (
			[string]$ThemePath
			)
		
		# Check if the theme file exists
		if (Test-Path $ThemePath) {

			LogThis "Activating the .theme file" -verboseMessage $true

			# Apply the theme; equivalent of double-clicking on a .theme file
			Start-Process $ThemePath

			# Wait a bit for the theme to apply and the Settings window to appear
			Start-Sleep -Seconds 4

			LogThis "Closing the Settings window" -verboseMessage $true

			# Close the Settings window by stopping the "ApplicationFrameHost" process
			$settingsProcess = Get-Process -Name "ApplicationFrameHost" -ErrorAction SilentlyContinue
			if ($settingsProcess) {
				Stop-Process -Id $settingsProcess.Id
			}

		} else {

			LogThis "Theme file not found: $ThemePath"
		}
	}
	
	<# Main function: Toggle the theme when running from Terminal
	using the command `./AutoTheme.ps1` #>
	function ToggleTheme {

		# Get current theme
		$CurrentTheme = (Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes" -Name CurrentTheme).CurrentTheme

		if ($CurrentTheme -match "dark")  {

			If (DoWeShuffle($lightPath)) {
				RandomFirstWall -wallpaperDirectory $wallLightPath
			}

			LogThis "Selected $lightPath" -verboseMessage $true

			# set Light theme
			StartTheme $lightPath

			# extra apps
			if ($restartProcexp) {RestartProcessExplorer}
			if ($customizeTrueLaunch) {UpdateTrueLaunch -themeMode "light" }

			# log it
			LogThis "$themeLight activated"
			ShowBurntToast -Text "Theme toggled. $themeLight activated." -AppLogo $appLogo

		}else {

			If (DoWeShuffle($darkPath)) {
				RandomFirstWall -wallpaperDirectory $wallDarkPath
			}

			LogThis "Selected $darkPath" -verboseMessage $true

			# set Dark theme
			StartTheme $darkPath

			# extra apps
			if ($restartProcexp) {RestartProcessExplorer}
			if ($customizeTrueLaunch) {UpdateTrueLaunch -themeMode "dark" }

			# log it
			LogThis "$themeDark activated"
			ShowBurntToast -Text "Theme toggled. $themeDark activated." -AppLogo $appLogo

		}
	}

	<# Main function: Create a 'temporary' scheduled task for next daylight events
	This task will be overwritten by the main task using this function. #>
	function CreateTemporaryTask {
		param (
			[DateTime]$NextTriggerTime,
			[String]$Name
		)
		
		# Schedule next run
		LogThis "Setting temporary Scheduled Task"
		$arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -NoProfile -File `"$PSCommandPath`""
		$fullCommand = "PowerShell.exe $arguments"
		LogThis "Full Command: $fullCommand"  -verboseMessage $true

		LogThis "Creating scheduled task action..." -verboseMessage $true
		$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $arguments

		$trigger = New-ScheduledTaskTrigger -Once -At $NextTriggerTime
		$userSid = $env:USERNAME
		$principal = New-ScheduledTaskPrincipal -UserId $userSid -LogonType Interactive -RunLevel Highest
		$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -Compatibility Win8

		$task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings
		LogThis "Scheduled task action created."  -verboseMessage $true

		# Unregister the old task if it exists
		if (TaskExists $Name) {

			Unregister-ScheduledTask -TaskName $Name -Confirm:$false
			LogThis "Unregistered existing task: $Name"  -verboseMessage $true
		}

		# Register the new task
		try {

			Register-ScheduledTask -TaskName $Name -InputObject $task | Out-Null
			LogThis "Registered new task: $Name"

		} catch {

			LogThis "Error registering task: $_"
		}	
	}

	<# Main function: Calculate daylight events or pick fixed hours 
	then select the Theme depending on daylight #>
	function ScheduleTheme {

		$Now = Get-Date

		if ($useFixedHours) {

			# stay offline
			$Sunrise = $lightThemeTime
			$Sunset = $darkThemeTime
			$TomorrowSunrise = $lightThemeTime

		} else {

			# go online
			$location = LocateThis

			# Extract latitude, longitude and Timezone for API call
			$lat = $location.Latitude
			$lng = $location.Longitude

			# Either API can be used, but the first one may have faulty control over DateTime formatting
			#$url = "https://api.sunrise-sunset.org/json?lat=$lat&lng=$lng"
			$url = "https://api.sunrisesunset.io/json?lat=$lat&lng=$lng"


			$Daylight = (Invoke-RestMethod $url).results
			LogThis "Fetched daylight data string = $Daylight" -verboseMessage $true
			
			# Parse and adjust the dates
			#$Sunrise = [DateTime]::ParseExact($Daylight.sunrise, "h:mm:ss tt", $null)
			#$Sunset = [DateTime]::ParseExact($Daylight.sunset, "h:mm:ss tt", $null)
			
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
			LogThis "Sunrise: $Sunrise, Sunset: $Sunset, TomorrowSunrise: $TomorrowSunrise, Now: $Now" -verboseMessage $true

		}

		# Get current theme
		$CurrentTheme = (Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes" -Name CurrentTheme).CurrentTheme

		# Determine the mode
		# Afer Sunrise and Before Sunset
		if ($Now -ge $Sunrise -and $Now -lt $Sunset) {

			# assign name for next temporary task
			$Name = "Sunset theme"

			$NextTriggerTime = $Sunset

			# if the Theme is already set, create temporary tasks and exit
			if ($CurrentTheme -match $themeLight)  {

				LogThis "Light mode is already set. No the switching needed."
				CreateTemporaryTask -NextTriggerTime $NextTriggerTime -Name $Name

				exit
			}			

			If (DoWeShuffle($lightPath)) {

				RandomFirstWall -wallpaperDirectory $wallLightPath	
			}		

			# set Light theme
			LogThis "Setting the theme  $lightPath" -verboseMessage $true
			StartTheme -ThemePath $lightPath

			# extra apps
			if ($restartProcexp) {RestartProcessExplorer}
			if ($customizeTrueLaunch) {UpdateTrueLaunch -themeMode "light" }

			# logging
			LogThis "$themeLight activated. Next trigger at: $NextTriggerTime"
			ShowBurntToast -Text "$themeLight activated. Next trigger at: $NextTriggerTime" -AppLogo $appLogo

		} else {

			if ($Now -ge $Sunset) {
				$NextTriggerTime = $TomorrowSunrise

			} else {

				$NextTriggerTime = $Sunrise
			}

			# assign name for next temporary task
			$Name = "Sunrise theme"

			# if the Theme is already set, create temporary tasks and exit
			if ($CurrentTheme -match $themeDark)  {

				LogThis "Dark mode is already set. No theme switching needed."
				CreateTemporaryTask -NextTriggerTime $NextTriggerTime -Name $Name

				exit
			}	

			If (DoWeShuffle($darkPath)) {

				RandomFirstWall -wallpaperDirectory $wallDarkPath	
			}		
			
			# set Dark theme
			LogThis "Setting the theme  $darkPath" -verboseMessage $true
			StartTheme -ThemePath $darkPath

			# extra apps
			if ($restartProcexp) {RestartProcessExplorer}
			if ($customizeTrueLaunch) {UpdateTrueLaunch -themeMode "dark" }

			# logging
			LogThis "$themeDark activated. Next trigger at: $NextTriggerTime"
			ShowBurntToast -Text "$themeDark activated. Next trigger at: $NextTriggerTime" -AppLogo $appLogo
		}	
	
		# Create the next temporary task
		CreateTemporaryTask -NextTriggerTime $NextTriggerTime -Name $Name
	}

# ============= RUNTIME  ==============

	try {

		# include config variables
		if (-Not (Test-Path $ConfigPath)) {
			Write-Error "Configuration file not found: $ConfigPath"
			Exit 1
		}
		. $ConfigPath
		
		# Trim old log sessions
		if ($trimLog -and -Not (IsRunningFromTerminal)) {TrimOldLog -logFilePath $logFile -maxSessions $maxLogEntries}
		
		# Start logging
		$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
		LogThis ""
		LogThis "$timestamp === Script started (Version: $scriptVersion)"

		# Optionally force admin mode
		if ($forceAsAdmin) {
			LogThis	"Running as Administrator." -verboseMessage $true
			RunAsAdmin
		}
						
		# Optionally restart Theme service, may solve issues with theme not being fully applied
		if ($restartThemeService){RestartThemeService}

		# Optionally check if the script was run recently
		if($checkLastRun){LastTime}

		# Update last run time
		UpdateTime			

		<# Here we call the functions to switch theme files,
		depending on whether running from command or from scheduled task. #>
		if (IsRunningFromTerminal) {

			LogThis "Script is running from Terminal." -verboseMessage $true
			LogThis "Toggling Theme, regardless of daylight."

			ToggleTheme

			LogThis "=== All done (Version: $scriptVersion)" -verboseMessage $true
			LogThis ""

		} else {

			LogThis "Script is running from Task Scheduler." -verboseMessage $true
			LogThis "Selecting and scheduling Theme based on daylight."

			ScheduleTheme

			LogThis "=== All done (Version: $scriptVersion)" -verboseMessage $true
			LogThis ""
		}

	} catch {

		LogThis "Error: $_"
	}
