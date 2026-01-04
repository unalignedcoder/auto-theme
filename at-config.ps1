
# ============= Theme variables =============

	<# If `$true`, the script will use `.theme` files to select light and dark themes, as well as wallpaper slideshows and other personalizations. If `$false`, it will use its own native system to switch between dark and light modes, and implement a wallpaper slideshow accordingly. #>
	$useThemeFiles = $true

	<# Name of theme files
	* Only relevant when `$useThemeFiles = $true` #>
	$themeLight = "Name-of-Light.theme"
	$themeDark = "Name-of-Dark.theme"

	<# COMPLETE PATH to the `.theme` files. You can use the default path to Windows themes
	(as proposed in the example below) or a custom path of your choice.
	You can use `$lightPath = Join-Path $PSScriptRoot $themeLight`
	if your `.theme` files are located within the script folder.
	* Only relevant when `$useThemeFiles = $true` #>
	$lightPath = Join-Path (Join-Path $Env:LOCALAPPDATA "Microsoft\Windows\Themes") $themeLight
	$darkPath =  Join-Path (Join-Path $Env:LOCALAPPDATA "Microsoft\Windows\Themes") $themeDark

# ============= Wallpaper variables =============

	<# If `$true`, the script will only switch between Dark and Light modes, ignoring wallpapers.
	* Only relevant when `$useThemeFiles = $false` #>
	$noWallpaperChange = $false

	<# Randomize wallpaper - when NOT using `.theme` files.
	If `$true`, picks a random wallpaper from the wallpaper paths.
	* Only relevant when `$useThemeFiles = $false` #>
	$randomizeWallpaper = $true

	<# Interval at which to randomize wallpapers. This effectively replaces Windows' own slideshow feature.
	This will create a specific Scheduled task to handle wallpaper changes at the defined interval.
	Set to "0" to disable this feature.
	* Only relevant when `$useThemeFiles = $false` #>
	$slideShowInterval = "0" # minutes

    <# Display the wallpaper name when it changes.
    "rainmeter"    - Displays the name as on overlay, via Rainmeter (a 'skin' is provided as example.)
    "notification" - Shows a Windows BurntToast notification.
    "none"         - Changes the wallpaper with no UI feedback. 
	* Only relevant when `$useThemeFiles = $false` 
	Note: To achieve the same effect when using `.theme` files or the Windows standard slideshow, check my Wallpaper Name Notification script on Github. N#>
    $howToDisplayName = "none" 

    <# Path to Rainmeter executable
	* Only relevant when `$useThemeFiles = $false` #>
    $rainmeterPath = "C:\Program Files\Rainmeter\Rainmeter.exe"

	<# The Config Name of your Rainmeter skin. 
    This is the relative path from your 'Skins' folder to the folder containing the `.ini` file.
    Example: If the skin is in 'Skins\AutoTheme\Wallpaper\at.ini', use "AutoTheme\Wallpaper"
	You can find in the repository a functional example of a skin that shows the wallpaper name at the bottom of the screen. 
	* Only relevant when `$useThemeFiles = $false`	#>
	$rainmeterSkinName = "AutoTheme\Wallpaper"

	<# Randomize FIRST wallpaper - when using `.theme` files AND shuffling wallpapers.
	Even if `shuffle=1` is set in a `.theme` file, Windows will always use the first wallpaper in alphabetic order as the first shown.
	Setting this to `$true` offers more variety as soon as the theme is applied.
	Be aware that, to this end, a randomly-picked wallpaper file will be temporarily renamed with a `0_at` string prepended to it.
	* Only relevant when `$useThemeFiles = $true` #>
	$randomFirst = $true

	<# Paths to the folders for light and dark wallpapers.
	If `$useThemeFiles = $true`, indicate here the same FOLDER PATHS indicated in the `.theme` files.
	If `$useThemeFiles = $false`, indicate PATHS to IMAGES OR FOLDERS to be selected for Light and Dark modes. #>
	$wallLightPath = "C:\Path\to\Light\wallpapers"
	$wallDarkPath = "C:\Path\to\Dark\wallpapers"

	# Show the wallpaper name in notification
	$showWallName = $true

# ============= Time variables =============

	# Use fixed hours to switch Themes (keeps the script completely offline)
	$useFixedHours = $false

	<# Fixed hours for theme change (only needed if $useFixedHours = $true).
	You are free to use 12 or 24 hours formats here. #>
	$lightThemeTime = "07:00 AM"
	$darkThemeTime = "07:00 PM"

	<# Set to `$true` to always use a user-defined location.
	If `$false`, the script will attempt to retrieve location from the system
	or, failing that, from your ISP, which may not give accurate results. #>
	$useUserLoc = $false

	<# User-defined coordinates and timezone. You can obtain your coordinates from Google or similar services.
	You can find a list of timezone identifiers at this url:
	https://en.wikipedia.org/wiki/List_of_tz_database_time_zones 
	* Only relevant if `$useUserLoc = $true`#>
	$userLat = "40.7128"
	$userLng = "-74.0060"
	$UserTzid = "America/New_York"

	<# Add or remove minutes to Sunrise and Sunset times. This can help	to account for local peculiarities. Values can be negative, for example, -30 means the event triggers 30 minutes earlier. #>
	$sunriseOffset = 0
	$sunsetOffset = 0

# ============= Extra apps variables =============

	<# Sysinternals Process Explorer doesn't automatically change theme when the system theme is changed. Use this variable if you want it to be restarted.
	If you run Process Explorer as Admin, the script should also run as Admin for this to work.
	You can use the $forceAsAdmin variable below for the purpose. #>
	$restartProcexp = $false
	# If `$true`, Process Explorer will keep Admin rights (inheriting from this script).
	# If `$false`, it will be restarted as a standard user. All other apps are restarted as standard user.
	$restartProcexpElevated = $true

	<# Change TrueLaunchBar colors (will cause Explorer to be restarted.)
	Look into the 'Update-TrueLaunchBar-colors' function for more details #>
	$customizeTrueLaunch = $false
	$trueLaunchIniFilePath = Join-Path $Env:APPDATA "Tordex\True Launch Bar\settings\setup.ini"

	<# Changes color in the T-Clock font, so that it adapts to the current theme.
	When T-Clock redraws the clock background based on accent color, it may cause the taskbar to crash or flicker. #>
	$tClockPath = "C:\Path\to\T-Clock\Clock64.exe"
	$updateTClockColor = $false
	$restartTClockColor = $false # usually not needed

# ============= Developer variables ==============

	$log = $true
	$logFromTerminal = $false
	$trimLog = $true
	$verbose = $false
	$lastRunInterval = "5"
	$waitExplorer = "30"
	$checkLastRun = $true
	$maxLogEntries = "5"

	# Try turning accent color off for Taskbar if you have surge issues when switching theme or wallpaper
	$turnOffAccentColor = $false

	$restartThemeService = $false
	$forceAsAdmin = $false

	# required files
	$appLogo = Join-Path $PSScriptRoot "at.png"
	$logFile = Join-Path $PSScriptRoot "at.log"
	$lastRunFile = Join-Path $PSScriptRoot "atLastRun.txt"
	$workerPath = Join-Path $PSScriptRoot "at-wallpaper.ps1"

	<# Console Visibility Settings when run from Task Scheduler:
	* "ch"- Completely hidden (no flash) using `conhost --headless`
	* "ps"- Standard PowerShell window (may flash briefly)
	* "wt"- Opens in Windows Terminal (using existing window if open) #>
	$terminalVisibility = "ch"
	
    # Write Wallpaper name in Registry for other apps or scripts to read
    $writeRegistry = $false