
# ========= User variables == CUSTOMIZE THIS! ===========

	# Name of theme files
	$themeLight = "myLight.theme"
	$themeDark = "myDark.theme"

	<# Complete path to the .theme files. You can use a system path to default Windows themes,
	a custom path of your choice, or even no path if the files are in the same folder as the script.
	However, consider that Windows will always copy your custom .theme files to LocalAppData. #>
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
	$UseUserLoc = $true

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

	$log = $true
	$verbose = $true
	$interval = "10" 
	$checkLastRun = $false
	
	$logFile = Join-Path $PSScriptRoot "AutoTheme.log"
	$lastRunFile = Join-Path $PSScriptRoot "ATLastRun.txt"