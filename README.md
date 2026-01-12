
# AUTO THEME
Powershell script which changes the active Windows theme and Desktop background based on daylight or a predefined schedule. Works in Windows 10/11. 

## Description
![Animation](https://github.com/user-attachments/assets/cd227523-78bf-42e4-a60e-ef21ce78c405)

**Auto Theme** is a powerful script designed to automatically and silently switch the Windows color mode, depending on daylight or hours set by the user. It can directly activate Windows <ins>Light</ins> and <ins>Dark</ins> modes and optionally handle its own <ins>dedicated wallpapers slideshows</ins> for each, so that you can have "dark" and "light" wallpapers showing at the right times.

Upon slideshow changes, it can <ins>display the name of the wallpaper</ins> in a notification or Rainmeter skin.

Alternatively, it can work by loading `.theme` files[^1]. This may allow for a higher degree of customization and compatibility on certain systems. 
If using `.theme` files and the standard Windows slideshow, displaying wallpaper names is still possible using my [companion script](https://github.com/unalignedcoder/wallpaper-name-notification).

**Auto Theme** includes a <ins>Desktop context menu</ins> to control the slideshow, toggle the theme or refresh the sunrise/sunset schedule. This is first created via the setup script, but can be disabled in the configuration file, and/or removed using the included `.reg` file in the repository.

<img width="697" height="160" alt="image" src="https://github.com/user-attachments/assets/b471b45f-307e-446b-908f-4c8d978acf83" />

This script is designed to run in the background as a scheduled task, ensuring that the system theme is updated without user intervention.

It operates <ins>almost entirely offline</ins>. It only connects to the internet to verify Location, in the case this is not retrivable from the system or is not provided by the user in the config file. 

Sunrise and sunset calculations are done natively (optionally the user can request to do this via free api services such as sunrisesunset.io, but it is mostly unnecessary.) 

Alternatively, the script can operate on fixed hours provided by the user.

When run from terminal, using the `./at.ps1 --Toggle`[^2] command (or just `./at.ps1` with no parameters), the script will 'toggle' the theme mode and then exit, <ins>ignoring any scheduled event</ins>.

![GIF 13 03 2025 1-30-58](https://github.com/user-attachments/assets/5ea7e34d-4e55-4cd4-a629-73f92ef2436c "The command can be run in terminal in verbose mode.")
<br /><sup>The command `.\at.ps1` can be run in terminal in verbose mode.</sup>

## Installation
1) Download the latest [release](https://github.com/unalignedcoder/auto-theme/releases) and extract it to your preferred folder.

2) Open the `at-config.ps1` file and modify variables as needed:

	<img width="847" height="1238" alt="All entries in the config file contain exhaustive explanations." src="https://github.com/user-attachments/assets/80a3b57f-047d-46fc-8be7-175300d562bb" />
	<br /><sup>All entries in the config file come with exhaustive explanations.</sup>
	
3) If using `.theme` files: 
	- Make sure `$useThemeFiles = $true` is set in the config file;
	- Modify settings in the _Personalize_ window (including colors or, for example, a wallpaper slideshow) and then save the theme;
	![image](https://github.com/user-attachments/assets/0999c082-16ec-456c-ba58-88783bc1abb3 "In the Personalize>Themes windows, right click on a theme and select 'Save for sharing'.")
	<br /><sup>In the Personalize>Themes windows, right click on a theme and select 'Save for sharing'.</sup>
	- Indicate the paths to the `.theme files` in the config file

4) Run the script `./at-setup.ps1`[^2] to create the main scheduled task and the desktop context menu. The script will ask for system privileges if not run as admin, and then proceed to create and launch the "Auto Theme" task. 

6) (alternative) You can of course create the task yourself using Task Scheduler, setting the triggers to anything you prefer. In this case, make sure that the Action is set up as follows:
	- Program/script: `conhost.exe`
	- Add arguments: `--headless Powershell.exe -WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -NoProfile -File "C:\path\to\at.ps1"`
	- Run with highest privileges.  
<b>It is advisable to always add the "On Workstation Unlock" trigger to the task. When the workstation is locked, the task may be unable to apply the theme fully, leaving out slideshow customizations and resulting in a hybrid "Custom" theme.</b> 

7) When triggered, the task will then run the script `at.ps1`. The script itself will schedule the next temporary task ("Sunrise Theme" or "Sunset theme") to run at the next required theme change time, whether set by the user or identified through user location. These task will be overwritten as a matter of course, to avoid clutter.

## Usage
This script is designed to run from Task Scheduler, and after the initial setup doesn't need interaction from the user.

The Scheduled tasks can run the script in a completely hidden manner, or visibly, per user choice.

For convenience, a shortcut to the script can be created and placed on the desktop or taskbar for quick access. In this case, the shortcut should be to `powershell.exe` followed by the path to the script `"C:\path\to\at.ps1"`, indicating the same path in the `Start in` field:
<p><img width="532" height="540" alt="A Windows shortcut can be created to directly toggle the theme." src="https://github.com/user-attachments/assets/b85cb2d7-91b1-44ef-90d4-b504b74c40df" />
<br /><sup>A Windows shortcut can be created to directly toggle the theme.</sup></p>

## Extra apps
Workarounds have been added for a number of apps which do not toggle theme gracefully when the system theme changes. More details in the Config file.

<p>&nbsp;</p>

<div align="center"><table border="1" cellspacing="0" cellpadding="20"><tr><td><p align="center">&nbsp;<br>Why, thank you for asking!</p><hr width="60%"><p align="center">👉 You can donate to all my projects <a href="https://www.buymeacoffee.com/unalignedcoder" target="_blank">here</a> 👈<br>&nbsp;</p></td></tr></table></div>

<p>&nbsp;</p>

## Changelog

### v1.0.45 (2026-01-13)
MAJOR UPDATE:
- Added parameters to control script functions
- Added the ability to move the slideshow to the next wallpaper
- Added a native OFFLINE sunrise/sunset calculation system
- Added a Desktop CONTEXT MENU to control the slideshow or theme
- Fixed a problem where wallpapers would not change if a task was created without the need to change theme
- Fixed geolocation (again)
- Added verification of the config file to the setup script
- Improved the Rainmeter sample skin
- Many minor fixes

### v1.0.43
- Fixed a problem that caused the wallpaper not to change if the PC was off a long time
- Improved the at-setup.ps1 script for the creation of the main scheduled task
- Many minor fixes

### v1.0.42
- Fixed loading/unloading of Rainmeter skins
- Added description to the Scheduled Tasks
- Attempt to correct suspect interference with Power settings
- Several minor fixes

### v1.0.40
MAJOR UPDATE:
- Added a native system to load Dark or Light modes and randomize wallpapers. 
- Renamed the script `at.ps1` for consistency with my other "short-named" projects.
- Changed the names of the generated tasks, make sure you delete the old ones in Task Scheduler.
- Added a "wrapper" script (`AutoTheme.ps1`) for compatibility with older tasks and existing shortcuts.
- Added a worker script (`at-wallpaper.ps1`) to handle the wallpaper slideshow natively via Task Scheduler.
- Fixed a problem with the script not recognizing it was running from Task Scheduler
- Improved geolocation
- Removed the MusicBee restart option, as not really helpful
- Many minor fixes

### v1.0.37
- Added logic to completely hide the console window when ran from Scheduler, or use Windows Terminal
- Fixed an issue with possible user identity spoofing when creating Scheduled Tasks, in main and setup scripts
- Scheduked Tasks are now created with battery settings to avoid the script being blocked on laptops
- Added logic to control elevation privileges of restarted apps.

### v1.0.33
- Corrected a problem with Taskbar colorization
- Improved variables in config file

### v1.0.32
- Added T-Clock to extra apps and Fixed theme switching issues with it.
- Changed function names to appease PS standards. Sigh.
- Several minor improvements and fixes.

### v1.0.27
- Improved the renaming of random wallpapers.
- Simplified/optimized several functions.
- Optionally show the wallpaper filename in notifications, when the theme changes. 
- Added MusicBee to the apps that can be optionally restarted upon theme change.
- Minor fixes.

### v1.0.19
- Improved the release-creation system
- Minor fixes

## Footnotes
[^1]:	`.theme` files are simple configuration instructions for the Windows pesonalization engine. You can find two examples included in this repository, with helpful comments.  
[^2]:	To run a PowerShell script on Windows, you need to set `Execution Policy` in PowerShell, using this command: `Set-ExecutionPolicy RemoteSigned` as Administrator.

