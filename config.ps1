
# ============= User variables =============

	# Name of theme files
	$themeLight = "Name-of-Light.theme"
	$themeDark = "Name-of-Dark.theme"

	<# Complete path to the .theme files. You can use a system path to default Windows themes
	(as proposed in the example below) or a custom path of your choice.
	However, consider that Windows will always copy your .theme files to LocalAppData. 
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
	(only needed if $UseUserLoc = $true nd $UseFixedHours = $false) #>
	$UserLat = "40.7128" 
	$UserLng = "-74.0060"
	
	<# Randomize first wallpaper: Even if 'shuffle=1' is set in a .theme file,
	Windows will always use the first wallpaper in alphabetic order as the first.
	Setting this to $true offers more variety as soon as the theme is applied. #>
	$RandomFirst = $true

	<# Paths to the folders for light and dark wallpapers.
	(only needed if $randomFirst = true) #>
	$wallLightPath = "Path\to\Light\wallpapers"
	$wallDarkPath = "Path\to\Dark\wallpapers"

# ============= Extra apps variables =============

	<# Sysinternals' Process Explorer doesn't automatically change theme when
	the system theme is changed. Use this variable if you want to restart it. #>
	$RestartProcexp = $false

	<# Change TrueLaunchBar colors (will cause Explorer to be restarted)
	Look into the 'Update-TrueLaunchBar-colors' function for more details #>
	$TrueLaunch = $false
	$TrueLaunchiniFilePath = Join-Path $Env:APPDATA "Tordex\True Launch Bar\settings\Setup.ini"
	
# ============= Developer variables ==============

	$log = $true
	$trimLog = $true
	$verbose = $false
	$interval = "10" 
	$checkLastRun = $true
	$themeServiceProblem = $true
	$maxLogEntries = "10""
	
	$appLogo = Join-Path $PSScriptRoot "autotheme.png"	
	$logFile = Join-Path $PSScriptRoot "AutoTheme.log"
	$lastRunFile = Join-Path $PSScriptRoot "ATLastRun.txt"