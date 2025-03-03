# AUTO THEME
Powershell script which automatically switches Dark and Light themes and desktop backgrounds under Windows, depending on the time of day.

## Description
This Powershell script automatically switches the Windows theme depending on Sunrise and Sunset, or hours set by the user.
Rather than using registry/system settings, it works by activating given `.theme` files. 
This allows for a much higher degree of customization and compatibility.

## The forgotten benefits of using `.theme` files
Many scripts and apps try to automate dark and light theme functionality under Windows 10/11, but they do so by modifying directly system registry settings and in doing so incurr in many difficulties and potential compatibility problems for the user.

This script however directly starts `.theme` files as processes (as if the user double-clicked on them), therefore letting the system itself seamlessly operate the entire visual transition, be it just the application of dark mode, or with addition of visual styles, wallpapers and more.

In fact, in addition to switching light and dark themes, `.theme` files allow to set different wallpaper slideshows for each theme, while including other changes such as cursors, sounds, and more, all without forcing or tricking the system into unusual behavior.

All it takes are `.theme` files that the user has created, or that can be found ready-made in the system.

## Installation
1) Create custom **Light** and **Dark** themes as preferred. To do so, simply modify settings in the _Personalize_ window (including for example a wallpaper slideshow) and then save it as custom theme.

	![image](https://github.com/user-attachments/assets/4aeb4d7d-35c4-4712-8fc4-2d183a5fb3f1)

Alternatively, use any `.theme` file found by defualt in the system, or downloaded online. Notice that `.theme` files can always be directly modified in a simple text editor. 

2) Open the file `Config.ps1` and modify the following variables as preferred:

 	![image](https://github.com/user-attachments/assets/059c57d9-b502-4467-a3de-bbaeb7ed01b8)

	- `themeLight` and `themeDark` should be the names of your custom `.theme` files.
	- `LightPath` and `$DarPath`should be the paths to your custom `.theme` files. Usually Windows saves them in `C:\Users\%username%\AppData\Local\Microsoft\Windows\Themes\`.
	- `$UseFixedHours` should be set to `$true` if you want to use fixed hours for the theme change. If set to `$false`, the script will try to find the sunrise and sunset times for your location. When set to `$true`, the script will keep offline.
	- `$LightThemeTime` and `$DarkThemeTime` should be set to the hours when you want the theme to change. If `$UseFixedHours` is set to `$true`, the script will change the theme at these hours. If `$UseFixedHours` is set to `$false`, these variables will be ignored.
	- `$UseUserLoc` should be set to `$true` if the user provides the exact coordinates and timezone in the following variables. If set to `$false`, the script will try to find the location using System Location or the IP address.
	- `$RandomFirst` should be set to `$true` if you want the script to randomize the first wallpaper shown after the theme has been changed. Windows, in fact, even when `shuffle=1`is indicated in the `.theme` file, always starts with the first alphabetical image in the list. If this is set to `$true`, the user should provide the paths to the wallpaper slideshow folders in the following variables. These should be the same as indicated in the `[Slideshow]` section of the `.theme` file.

4) (optional) Run the script `Setup.ps1` to create the main scheduled task. The script will ask for system privileges if not run as admin, and then proceed to create the "Auto Theme" task. 

5) (alternative) You can of course create the task yourself using Task Scheduler, setting the triggers to anything you prefer. In this case, make sure that the Action is set up as follows:
	- Program/script: `Powershell.exe`
	- Add arguments: `-WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -NoProfile -File "C:\path\to\AutoTheme.ps1"`
	- Run with highest privileges.

	![image](https://github.com/user-attachments/assets/4ec93663-3001-46ce-ad56-b16a623de8b1)

	![image](https://github.com/user-attachments/assets/048e6e91-fe0e-4bf0-905c-3beb2aeb4385)

	![image](https://github.com/user-attachments/assets/f2dcbd94-beee-477d-8f5c-5868c3780dc0)
	
	<b>It is advisable to always add the "On Workstation Unlock" trigger to the task. When the workstation is locked, the task may be unable to apply the theme fully, leaving out for example Slideshow customization and resulting in a hybrid "Custom" theme.</b>

6) When triggered, the task will then run the script `AutoTheme.ps1`. The script itself will schedule the next temporary task ("Sunrise Theme" or "Sunset theme") to run at the next required theme change time, whether set by the user or identified through user location.

## Usage
This script is designed to run from Task Scheduler, and after the initial setup doesn't need interaction from the user. 

When run from terminal, however, using `.\AutoTheme.ps1`, the script will simply 'toggle' the theme (switching from one `.theme` file to the other) and then exit. This can be useful for testing purposes, but also for the odd times when we need to manually switch the theme regardless of task settings. 

![Animation](https://github.com/user-attachments/assets/4beeacf2-f026-4304-99ed-9b70f56f29fd)

For convenience. you can create a shortcut to the script and place it on your desktop or taskbar for quick access. In this case, the shortcut should be to `powershell.exe` followed by the path to the script `"C:\path\to\AutoTheme.ps1"`, indicating the same path in the `Start in` field:

![image](https://github.com/user-attachments/assets/954d4a76-3001-4bd8-9e7d-460c1db3888a)

## Extra apps
Workarounds have been added for a couple of apps who do not switch theme gracefully when the system theme changes: TrueLaunchBar and ProcessExplorer. More will be added in the future if there is demand. More details in the Config file.

