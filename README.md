
# AUTO THEME
Powershell script which changes the active Windows theme and Desktop background based on daylight or a predefined schedule. Works in Windows 10/11.  Not tested in Windows 7.

## Description
![Animation](https://github.com/user-attachments/assets/cd227523-78bf-42e4-a60e-ef21ce78c405)

**Auto Theme** is a powerful script designed to automatically and silently switch the Windows color mode, depending on daylight or hours set by the user. It can directly activate Windows <ins>Light</ins> and <ins>Dark</ins> modes and optionally handle its own <ins>dedicated wallpapers slideshows</ins> for each, so that you can have "dark" and "light" wallpapers showing at the right times. 
Upon slideshow changes, it can <ins>display the name of the wallpaper</ins> in a notification or Rainmeter skin.

Alternatively, it can work by loading `.theme` files[^1]. This may allow for a higher degree of customization and compatibility on certain systems. 
If using `.theme` files and the standard Windows slideshow, displaying wallpaper names is still possible using my [companion script](https://github.com/unalignedcoder/wallpaper-name-notification).

**Auto Theme** is designed to run in the background as a scheduled task, ensuring that the system theme is updated without user intervention.

It only connects to the internet to verify Location and retrieve Sunrise and Sunset times, via free api services such as sunrisesunset.io and ip-api.com. 

It can also stay completely offline by operating on fixed hours provided by the user.

When run from terminal, using the `.\at.ps1`[^2] command, the script will 'toggle' the theme mode and then exit, ignoring any scheduled event.

![GIF 13 03 2025 1-30-58](https://github.com/user-attachments/assets/5ea7e34d-4e55-4cd4-a629-73f92ef2436c "The command can be run in terminal in verbose mode.")
<br /><sup>The command `.\at.ps1` can be run in terminal in verbose mode.</sup>

## Installation
1) Download the latest [release](https://github.com/unalignedcoder/auto-theme/releases) and extract it to your preferred folder.

2) Open the `at-config.ps1` file and modify variables as preferred:

	<img width="847" height="1238" alt="All entries in the config file contain exhaustive explanations." src="https://github.com/user-attachments/assets/80a3b57f-047d-46fc-8be7-175300d562bb" />
	<br /><sup>All entries in the config file come with exhaustive explanations.</sup>
	
3) If using `.theme` files: 
	- Make sure `$useThemeFiles = $true` is set in the config file;
	- Modify settings in the _Personalize_ window (including colors or, for example, a wallpaper slideshow) and then save the theme;
	![image](https://github.com/user-attachments/assets/0999c082-16ec-456c-ba58-88783bc1abb3 "In the Personalize>Themes windows, right click on a theme and select 'Save for sharing'.")
	<br /><sup>In the Personalize>Themes windows, right click on a theme and select 'Save for sharing'.</sup>
	- Indicate the paths to the `.theme files` in the config file

4) (optional) Run the script `.\at-setup.ps1`[^2] to create the main scheduled task. The script will ask for system privileges if not run as admin, and then proceed to create the "Auto Theme" task. 

6) (alternative) You can of course create the task yourself using Task Scheduler, setting the triggers to anything you prefer. In this case, make sure that the Action is set up as follows:
	- Program/script: `Powershell.exe`
	- Add arguments: `-WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -NoProfile -File "C:\path\to\at.ps1"`
	- Run with highest privileges.
<b>It is advisable to always add the "On Workstation Unlock" trigger to the task. When the workstation is locked, the task may be unable to apply the theme fully, leaving out slideshow customizations and resulting in a hybrid "Custom" theme.</b> 

7) When triggered, the task will then run the script `at.ps1`. The script itself will schedule the next temporary task ("Sunrise Theme" or "Sunset theme") to run at the next required theme change time, whether set by the user or identified through user location. These task will be overwritten as a matter of course, to avoid clutter.

## Usage
This script is designed to run from Task Scheduler, and after the initial setup doesn't need interaction from the user.

The Scheduled tasks can run the script in a completely hidden manner, or in a visible way, per user choice.

For convenience, a shortcut to the script can be created and placed on the desktop or taskbar for quick access. In this case, the shortcut should be to `powershell.exe` followed by the path to the script `"C:\path\to\at.ps1"`, indicating the same path in the `Start in` field:
<p><img width="532" height="540" alt="A Windows shortcut can be created to directly toggle the theme." src="https://github.com/user-attachments/assets/b85cb2d7-91b1-44ef-90d4-b504b74c40df" />
<br /><sup>A Windows shortcut can be created to directly toggle the theme.</sup></p>

## Extra apps
Workarounds have been added for a number of apps which do not toggle theme gracefully when the system theme changes. More details in the Config file.

&nbsp;

<p align=center>Why, thank you for asking!<br />👉 You can donate to all my projects <a href="https://www.buymeacoffee.com/unalignedcoder" target="_blank" title="buymeacoffee.com">here</a>👈</p>

&nbsp;

[^1]:	`.theme` files are simple configuration instructions for the Windows pesonalization engine. You can find two examples included in this repository, with helpful comments.
[^2]:	To run a PowerShell script on Windows, you need to set `Execution Policy` in PowerShell, using this command: `Set-ExecutionPolicy RemoteSigned` as Administrator.


