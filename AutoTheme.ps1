<#
.SYNOPSIS
Changes the active Windows theme based on a predefined schedule.

.DESCRIPTION
This Powershell script automatically switches the Windows theme based on Sunrise and Sunset, or hours set by the user.
Rather than using registry/system settings, it selects a given .theme file. This allows for a much higher degree of customization.
The script is designed to run in the background as a scheduled task, ensuring that the system theme is updated without user intervention.
It only connects to the internet to verify Location and Sunrise and Sunset times depending on user location.
Alternatively, it can use hours provided by the user, thus staying offline.
The script is meant to be ran from Task Scheduler, and it will automatically create the next temporary task.
If otherwise the script is run from terminal, as './AutoTheme.ps1', it only switches between the themes.
#>

# Script version
$scriptVersion = "1.0.11"

# ============= Config file ==============

	$ConfigPath = "$PSScriptRoot\config.ps1"

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

			} else {

				Add-Content -Path $logFile -Value "$message"  # Log to file

			}
		}
	}

	# Handle BurntToast Notifications
	function BurnedToast {
		param(
			[string]$Text,
			[string]$AppLogo
		)

		# Install the BurntToast module if not already installed
		if (-not (Get-Module -Name BurntToast -ListAvailable)) {
			Install-Module -Name BurntToast -Scope CurrentUser
		}

		# for when the above is commented out or fails, we double-check.
		if (Get-Module -Name BurntToast -ListAvailable) {

			try {

				$logoFullPath = Join-Path -Path $PSScriptRoot -ChildPath $AppLogo
				New-BurntToastNotification -Text $Text -AppLogo $logoFullPath

				LogThis "Displayed BurntToast notification with text: $Text and logo: $logoFullPath"  -verboseMessage $true

			} catch {

				LogThis "Error displaying BurntToast notification: $_"  -verboseMessage $true
			}

		} else {

			LogThis "BurntToast module is not installed. Cannot display system notifications."  -verboseMessage $true
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
			[double]$FallbackLatitude = $UserLat,
			[double]$FallbackLongitude = $UserLng,
			[string]$FallbackTimezone = $UserTzid
		)

		LogThis "Getting location coordinates." -verboseMessage $true

		# If $UseUserLoc is set to true, return user-defined coordinates and timezone
		if ($UseUserLoc) {

			LogThis "Using user-defined coordinates and timezone." -verboseMessage $true

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
			$UserLat = $position.Coordinate.Point.Position.Latitude
			$longitude = $position.Coordinate.Point.Position.Longitude
			$UserTzid = [System.TimeZone]::CurrentTimeZone.StandardName

			LogThis "Retrieved device location and system timezone." -verboseMessage $true

			return @{
				Latitude = [double]$UserLat
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

		LogThis "Using user-defined coordinates and timezone." -verboseMessage $true

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
				LogThis "Found [Slideshow] section" -verboseMessage $true
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

	# Run the .theme file
	function StartTheme {
		param (
			[string]$ThemePath
			)
			
		#restart Theme service, solves issues with theme not being applied
		if ($themeServiceProblem) {	
			try {
				LogThis "Restarting the Themes service..." -verboseMessage $true
				Restart-Service -Name "Themes" -Force
				LogThis "Themes service restarted successfully." -verboseMessage $true
			} catch {
				LogThis "Failed to restart Themes service: $_"  -verboseMessage $true
			}
		}
		
		# Check if the theme file exists
		if (Test-Path $ThemePath) {

			# Apply the theme
			Start-Process $ThemePath

			# Wait a bit for the theme to apply and the Settings window to appear
			Start-Sleep -Seconds 4

			# Close the Settings window by stopping the "ApplicationFrameHost" process
			$settingsProcess = Get-Process -Name "ApplicationFrameHost" -ErrorAction SilentlyContinue
			if ($settingsProcess) {
				Stop-Process -Id $settingsProcess.Id
			}

		} else {

			LogThis "Theme file not found: $ThemePath"
		}
	}

	# Toggle the theme
	function ToggleTheme {

		# Get current theme
		$CurrentTheme = (Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes" -Name CurrentTheme).CurrentTheme

		if ($CurrentTheme -match "dark")  {

			If (DoWeShuffle($LightPath)) {
				RandomFirstWall -wallpaperDirectory $wallLightPath
			}

			LogThis "Selected $LightPath"  -verboseMessage $true
			StartTheme $LightPath
			LogThis "$themeLight activated"

		}else {

			If (DoWeShuffle($DarkPath)) {
				RandomFirstWall -wallpaperDirectory $wallDarkPath
			}

			LogThis "Selected $DarkPath"  -verboseMessage $true
			StartTheme $DarkPath
			LogThis "$themeDark activated"
		}
	}

	<# Prepend the substring '000_' to one randomly chosen
	wallpaper filename, so as to make it first pick. #>
	function RandomFirstWall {
		param (
			[string]$wallpaperDirectory
		)

		if (-Not ($RandomFirst)) {
			
			LogThis "The first wallpaper will not be randomized."  -verboseMessage $true			
			return
		}

		LogThis "Randomizing first wallpaper."  -verboseMessage $true

		# Retrieve file objects directly into $wallpapers
		LogThis "Looking in $wallpaperDirectory"  -verboseMessage $true
		$wallpapers = Get-ChildItem -Path $wallpaperDirectory -File

		# Check if there's already a file with "000_" prefix and rename it back
		$existingRenamedWallpaper = $wallpapers | Where-Object { $_.Name.StartsWith("000_") }
		if ($existingRenamedWallpaper) {
			$originalName = $existingRenamedWallpaper.Name -replace '^000_', ''
			$originalNameFull = Join-Path $wallpaperDirectory $originalName
			Rename-Item -Path $existingRenamedWallpaper.FullName -NewName $originalNameFull
			LogThis "Removed '000_' prefix from $($existingRenamedWallpaper.FullName)"  -verboseMessage $true
		}

		# Filter out wallpapers that still have "000_" in the name (unlikely after removal)
		$newWallpapers = $wallpapers | Where-Object { -not $_.Name.StartsWith("000_") }

		# Ensure there are wallpapers available to rename
		if ($newWallpapers.Count -eq 0) {
			LogThis "No wallpapers available for renaming."  -verboseMessage $true
			return
		}

		# Select a random wallpaper from the filtered list
		$RandomFirstWallpaper = $newWallpapers | Get-Random

		# Construct the new name and rename
		$newWallpaperName = "000_" + $RandomFirstWallpaper.Name
		$newWallpaperNameFull = Join-Path $wallpaperDirectory $newWallpaperName
		Rename-Item -Path $RandomFirstWallpaper.FullName -NewName $newWallpaperNameFull
		LogThis "Renamed $($RandomFirstWallpaper.FullName) to $newWallpaperNameFull"  -verboseMessage $true	
	}

	# Check if a Scheduled task exists
	function TaskExists {
		param (
			[string]$taskName
		)
		
		$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
		return $null -ne $task
	}

	<# Select the Theme depending on daylight or chosen hours,
	then schedules the next appropriate temporary Task #>
	function Main {

		if ($UseFixedHours) {

			# stay offline
			$Sunrise = $LightThemeTime
			$Sunset = $DarkThemeTime
			$TomorrowSunrise = $LightThemeTime

		} else {

			# go online
			$location = LocateThis

			# Extract latitude, longitude and Timezone for API call
			$lat = $location.Latitude
			$lng = $location.Longitude
			$tzid = $location.Timezone

			$url = "https://api.sunrise-sunset.org/json?lat=$lat&lng=$lng&tzid=$tzid"
			$Daylight = (Invoke-RestMethod $url).results
			LogThis "Fetched daylight data." -verboseMessage $true

			# Parse and adjust the dates
			$Sunrise = [DateTime]::ParseExact($Daylight.sunrise, "h:mm:ss tt", $null)
			$Sunset = [DateTime]::ParseExact($Daylight.sunset, "h:mm:ss tt", $null)

			$TomorrowDaylight = (Invoke-RestMethod "$url&date=tomorrow").results
			$TomorrowSunrise = [DateTime]::ParseExact($TomorrowDaylight.sunrise, "h:mm:ss tt", $null).AddDays(1)
			LogThis "Sunrise: $Sunrise, Sunset: $Sunset, TomorrowSunrise: $TomorrowSunrise, Now: $Now"  -verboseMessage $true

		}

		$Now = Get-Date

		# Get current theme
		$CurrentTheme = (Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes" -Name CurrentTheme).CurrentTheme

		# Determine the mode
		if ($Now -ge $Sunrise -and $Now -lt $Sunset) {

			if ($CurrentTheme -match "light")  {
				LogThis "The Mode is already set. No action needed."
				exit
			}			

			$NextTriggerTime = $Sunset

			If (DoWeShuffle($LightPath)) {
				RandomFirstWall -wallpaperDirectory $wallLightPath	
			}		

			LogThis "Setting the theme  $LightPath"  -verboseMessage $true
			StartTheme -ThemePath $LightPath

			LogThis "$themeLight activated. Next trigger at: $NextTriggerTime"
			BurnedToast -Text "$themeLight activated. Next trigger at: $NextTriggerTime" -AppLogo "autotheme.png"

			$Name = "Sunset theme"

		} else {

			if ($CurrentTheme -match "dark")  {
				LogThis "The Mode is already set. No action needed."
				exit
			}			

			if ($Now -ge $Sunset) {
				$NextTriggerTime = $TomorrowSunrise
			} else {
				$NextTriggerTime = $Sunrise
			}

			If (DoWeShuffle($DarkPath)) {
				RandomFirstWall -wallpaperDirectory $wallDarkPath	
			}		
		
			LogThis "Setting the theme  $LightPath" -verboseMessage $true
			StartTheme -ThemePath $DarkPath

			LogThis "$themeDark activated. Next trigger at: $NextTriggerTime"
			BurnedToast -Text "$themeDark activated. Next trigger at: $NextTriggerTime" -AppLogo "autotheme.png"

			$Name = "Sunrise theme"

		}	

		# Schedule next run
		LogThis "Setting temporary Scheduled Task"
		$arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -NoProfile -File `"$PSCommandPath`""
		$fullCommand = "PowerShell.exe $arguments"
		LogThis "Full Command: $fullCommand"  -verboseMessage $true

		LogThis "Creating scheduled task action..." -verboseMessage $true
		$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $arguments

		$trigger = New-ScheduledTaskTrigger -Once -At $NextTriggerTime
		#$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
		#$userSid = $currentUser.User.Value  # Get the SID of the current user
		$userSid = $env:USERNAME
		$principal = New-ScheduledTaskPrincipal -UserId $userSid -LogonType Interactive -RunLevel Highest
		$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable

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
			LogThis "Registered new task: $Name"  -verboseMessage $true
		} catch {
			LogThis "Error registering task: $_"
		}

	
	}

# ============= RUNTIME  ==============

	try {

		# include config variables
		if (-Not (Test-Path $ConfigPath)) {
			Write-Error "Configuration file not found: $ConfigPath"
			Exit 1
		}
		. $ConfigPath

		# Start logging
		$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
		LogThis "$timestamp === Script started (Version: $scriptVersion)"

		# Check if the script was run recently
		if($checkLastRun){LastTime}

		# Running from terminal or from Scheduled Task?
		if (IsRunningFromTerminal) {

			LogThis "Script is running from Terminal." -verboseMessage $true

			# Toggle the theme and exit
			LogThis "Toggling the Theme"
			ToggleTheme
			exit

		} else {

			LogThis "Script is running from Task Scheduler." -verboseMessage $true
			LogThis "Selecting Theme based on daylight"

			# Main function
			Main

		}

		# Update last run time
		UpdateTime

	} catch {

		LogThis "Error: $_"
	}

	LogThis "All done."
	exit