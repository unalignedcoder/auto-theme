# AUTO THEME
Powershell script which changes the active Windows theme and Desktop background based on daylight or a predefined schedule. Works in Windows 10/11.  Not tested in Windows 7.

## Description
This script automatically switches the Windows active theme depending on Sunrise and Sunset daylight times, or hours set by the user.

It directly activates Windows Light and Dark modes at these times, and handles its own native dedicated wallpapers slideshows for each. 

Alternatively, it can work by activating dedicated `.theme` files. This may allow for a higher degree of customization and compatibility on certain systems.

**Auto Theme** is designed to run in the background as a scheduled task, ensuring that the system theme is updated without user intervention.

It only connects to the internet to verify Location and retrieve Sunrise and Sunset times, via free api services such as sunrisesunset.io and ip-api.com. 

It can also stay completely offline by operating on fixed hours provided by the user.

When run as the command `.\at.ps1` from terminal or desktop shortcut, the script toggles between themes, ignoring scheduled events.

## Installation
1) Download the latest [release](https://github.com/unalignedcoder/auto-theme/releases) and extract it to your preferred folder.

2) Open the `Config.ps1` file and modify variables as preferred:

 	![image](https://github.com/user-attachments/assets/166b21d9-7a56-4686-9376-641abc58727b "All entries in the config file contain exhaustive explanations.")
	<br /><sup>All entries in the config file come with exhaustive explanations.</sup>
	
3) If using `.theme` files[^1]: 
	- Make sure `$useThemeFiles = $true` is set in the config file;
	- Modify settings in the _Personalize_ window (including colors or, for example, a wallpaper slideshow) and then save the theme;
	![image](https://github.com/user-attachments/assets/0999c082-16ec-456c-ba58-88783bc1abb3 "In the Personalize>Themes windows, right click on a theme and select 'Save for sharing'.")
	<br /><sup>In the Personalize>Themes windows, right click on a theme and select 'Save for sharing'.</sup>
	- Indicate the paths to the `.theme files` in the config file

4) (optional) Run the script `.\Setup.ps1`[^2] to create the main scheduled task. The script will ask for system privileges if not run as admin, and then proceed to create the "Auto Theme" task. 

6) (alternative) You can of course create the task yourself using Task Scheduler, setting the triggers to anything you prefer. In this case, make sure that the Action is set up as follows:
	- Program/script: `Powershell.exe`
	- Add arguments: `-WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -NoProfile -File "C:\path\to\at.ps1"`
	- Run with highest privileges.
<b>It is advisable to always add the "On Workstation Unlock" trigger to the task. When the workstation is locked, the task may be unable to apply the theme fully, leaving out slideshow customizations and resulting in a hybrid "Custom" theme.</b> 

7) When triggered, the task will then run the script `at.ps1`. The script itself will schedule the next temporary task ("Sunrise Theme" or "Sunset theme") to run at the next required theme change time, whether set by the user or identified through user location. These task will be overwritten as a matter of course, to avoid clutter.

## Usage
This script is designed to run from Task Scheduler, and after the initial setup doesn't need interaction from the user.

The Scheduled tasks can run the script in a completely hidden manner, or in a visible way, per user choice.

When run from terminal, using `.\at.ps1`[^2], the script will 'toggle' the theme mode and then exit (optionally modifying wallpapers), ignoring any scheduled event. This can be useful for testing purposes, but also for the odd times when there is need to manually switch the theme regardless of task settings. 

![GIF 13 03 2025 1-30-58](https://github.com/user-attachments/assets/5ea7e34d-4e55-4cd4-a629-73f92ef2436c "The command can be run in terminal in verbose mode.")
<br /><sup>The command `.\at.ps1` can be run in terminal in verbose mode.</sup>

For convenience, a shortcut to the script can be created and placed on the desktop or taskbar for quick access. In this case, the shortcut should be to `powershell.exe` followed by the path to the script `"C:\path\to\AutoTheme.ps1"`, indicating the same path in the `Start in` field:

![image](https://github.com/user-attachments/assets/f8e2d534-7696-464d-9d83-e18a39ea9942 "A Windows shortcut can be created to directly toggle the theme.")
<br /><sup>A Windows shortcut can be created to directly toggle the theme.</sup>

## Extra apps
Workarounds have been added for a number of apps which do not toggle theme gracefully when the system theme changes. More details in the Config file.

&nbsp;

<p align=center>Why, thank you for asking!<br />👉 You can donate to all my projects <a href="https://www.buymeacoffee.com/unalignedcoder" target="_blank" title="buymeacoffee.com">here</a>👈</p>

&nbsp;

[^1]:	`.theme` files are simple configuration instructions for the Windows pesonalization engine. You can find two examples included in this repository, with helpful comments.
[^2]:	To run a PowerShell script on Windows, you need to set `Execution Policy` in PowerShell, using this command: `Set-ExecutionPolicy RemoteSigned` as Administrator.


