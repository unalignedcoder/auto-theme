# ============= User variables =============

	# Name of theme files
	$themeLight = "Name-of-Light.theme"
	$themeDark = "Name-of-Dark.theme"

	<# Complete path to the `.theme` files. You can use the default path to Windows themes
	(as proposed in the example below) or a custom path of your choice.
	Consider that Windows will always copy your `.theme` files to LocalAppData.
	You can use something like `$lightPath = Join-Path $PSScriptRoot $themeLight`
	if your `.theme` files are located within the script folder. #>
	$lightPath = Join-Path (Join-Path $Env:LOCALAPPDATA "Microsoft\Windows\Themes") $themeLight
	$darkPath =  Join-Path (Join-Path $Env:LOCALAPPDATA "Microsoft\Windows\Themes") $themeDark

	# Use fixed hours to switch Themes (keeps the script completely offline)
	$useFixedHours = $false

	<# Fixed hours for theme change (only needed if $useFixedHours = $true).
	You are free to use 12 or 24 hours formats here. #>
	$lightThemeTime = "07:00 AM"
	$darkThemeTime = "07:00 PM"

	<# Set to $true to always use a user-defined location.
	Alernatively, the script will attempt to retrieve location from the system
	or, failing that, from your ISP, which may not give accurate results. #>
	$useUserLoc = $false

	<# User-defined coordinates and timezone. You can obtain your coordinates from Google or similar services.
	You can find a list of timezone identifiers at this url:
	https://en.wikipedia.org/wiki/List_of_tz_database_time_zones
	(These values are only retrieved if $UseUserLoc = $true and $useFixedHours = $false.
	Yet, better set them, as the script will fall back to them, if all else fails.) #>
	$userLat = "40.7128"
	$userLng = "-74.0060"
	$UserTzid = "America/New_York"

	<# Randomize first wallpaper.
	Even if 'shuffle=1' is set in a `.theme` file, Windows will always use
	the first wallpaper in alphabetic order as the first shown.
	Setting this to $true offers more variety as soon as the theme is applied. 
	Be aware that, to this end, a randomly-picked wallpaper file
	will be temporarily renamed with a "000_" string prepended to it. #>
	$randomFirst = $true

	<# Paths to the folders for light and dark wallpapers.
	(only needed if $randomFirst = true) #>
	$wallLightPath = "C:\Path\to\Light\wallpapers"
	$wallDarkPath = "C:\Path\to\Dark\wallpapers"

	# Show the wallpaper name in notification
	$showWallName = $true

# ============= Extra apps variables =============

	<# Sysinternals Process Explorer doesn't automatically change theme when
	the system theme is changed. Use this variable if you want it to be restarted.
	If you run Process Explorer as Admin, the script should also run as Admin for this to work.
	You can use the $forceAsAdmin variable below for the purpose. #>
	$restartProcexp = $true

	<# Change TrueLaunchBar colors (will cause Explorer to be restarted)
	Look into the 'Update-TrueLaunchBar-colors' function for more details #>
	$customizeTrueLaunch = $true
	$trueLaunchIniFilePath = Join-Path $Env:APPDATA "Tordex\True Launch Bar\settings\setup.ini"

	<# MusicBee will not switch theme when the theme changes, unless restarted. #>
	$restartMusicBee = $true

# ============= Advanced variables ==============

	$log = $true
	$logFromTerminal = $false
	$trimLog = $true
	$verbose = $false
	$lastRunInterval = "5"
	$waitExplorer = "30"
	$checkLastRun = $true
	$maxLogEntries = "20"

	$restartThemeService = $false
	$forceAsAdmin = $false

	$appLogo = Join-Path $PSScriptRoot "AutoTheme.png"
	$logFile = Join-Path $PSScriptRoot "AutoTheme.log"
	$lastRunFile = Join-Path $PSScriptRoot "ATLastRun.txt"

