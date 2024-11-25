
# ========= User variables == CUSTOMIZE THIS! ===========

	# Name of theme files
	$themeLight = "Name-of-Light.theme"
	$themeDark = "Name-of-Dark.theme"

	<# Complete path to the .theme files. You can use a system path to default Windows themes
	(as proposed in the example below) or a custom path of your choice.
	However, consider that Windows will always copy your custom .theme files to LocalAppData. 
	In order to use .theme files located in the script folder, you can use this:
	$LightPath = Join-Path $PSScriptRoot $themeLight #>
	$LightPath = Join-Path (Join-Path $Env:LOCALAPPDATA "Microsoft\Windows\Themes") $themeLight
	$DarkPath =  Join-Path (Join-Path $Env:LOCALAPPDATA "Microsoft\Windows\Themes") $themeDark

	# Use fixed hours to switch Themes (keeps the script offline)
	$UseFixedHours = $false
	# Fixed hours for theme change (only needed if $UseFixedHours = $true)
	$LightThemeTime = "07:00 AM"
	$DarkThemeTime = "07:00 PM"

	<# Set to $true to always use a user-defined location.
	Alernatively, the script will attempt to retrieve location from the system
	or, failing that, from your ISP	
	(only needed if $UseFixedHours = $false) #>
	$UseUserLoc = $false

	<# User-defined coordinates  
	(only needed if $UseFixedHours = $false) #>
	$UserLat = "40.7128" 
	$UserLng = "-74.0060"

	<# User-defined timezone
	(only needed if $UseFixedHours = $false) #>
	$UserTzid = "America/New_York"  
	
	<# Randomize first wallpaper: Even if 'shuffle=1' is set in a `.theme` file
	Windows will always use the first wallpaper in alphabetic order as the first. #>
	$RandomFirst = $true

	<# Paths to the folders for light and dark wallpapers.
	(only needed if $randomFirst = true) #>
	$wallLightPath = "Path\to\Light\wallpapers"
	$wallDarkPath = "Path\to\Dark\wallpapers"

	
# ============= Developer variables ==============

	$log = false
	$verbose = $false
	$interval = "10" 
	$checkLastRun = $true
	$themeServiceProblem = $true
	
	$logFile = Join-Path $PSScriptRoot "AutoTheme.log"
	$lastRunFile = Join-Path $PSScriptRoot "ATLastRun.txt"