# AUTO THEME
Powershell script which automatically switches `.theme` files depending on the time of day.

## Description
This script automatically alternates between two  `.theme` files of your choice, either via scheduled task or when run directly. It finds sunrise or sunset times retriving location, or it stays offline using the hours indicated by the user.

## The ADVANTAGE of using `.theme` files
Many scripts and apps try to automate dark and light theme functionality under Windows 10/11, but they do so by modifying directly system registry settings. They then try to force a system refresh to show changes and in doing so incurr in many difficulties and potential compatibility problems for the user.

This script however directly starts `.theme` files as processes, therefore letting the system itself seamlessly operate the entire visual transition, be it dark mode, visual styles or wallpapers.

In fact, using `.theme` files allows to set different wallpaper slideshows for each theme, while including other changes such as cursors, sounds, and more, all without forcing or tricking the system into odd behavior.

All it takes are `.theme` files that the user has created, or that can be found ready-made in the system.

## Installation
1) Create custom **Light** and **Dark** themes to your preference. To do so, simply modify settings in the _Personalize_ window (including for example a wallpaper slideshow) and then save it as custom theme. You can then also modify the `.theme` file in your text editor as needed. Alternatively, use any `.theme` file found by defualt in the system, or downloaded online.

2) Open the file `Config.ps1` and modify the following variables to your preference:

	![image](https://github.com/user-attachments/assets/2d0b57f1-3f0a-4829-812e-d9cb6fa27031)

- `themeLight` and `themeDark` should be the names of your custom `.theme` files.
- `LightPath` and `$DarPath`should be the paths to your custom `.theme` files. Usually Windows saves them in `C:\Users\%username%\AppData\Local\Microsoft\Windows\Themes\`.
- `$UseFixedHours` should be set to `$true` if you want to use fixed hours for the theme change. If set to `$false`, the script will try to find the sunrise and sunset times for your location. When set to `$true`, the script will keep offline.
- `$LightThemeTime` and `$DarkThemeTime` should be set to the hours when you want the theme to change. If `$UseFixedHours` is set to `$true`, the script will change the theme at these hours. If `$UseFixedHours` is set to `$false`, these variables will be ignored.
- `$UseUserLoc` should be set to `$true` if the user provides the exact coordinates and timezone in the following variables. If set to `$false`, the script will try to find the location using System Location or the IP address.
- `$RandomFirst` should be set to `$true` if you want the script to randomize the first wallpaper shown after the theme has been changed. Windows, in fact, even when `shuffle=1`is indicated in the `.theme` file, always starts with the first alphabetical image in the list. If this is set to `$true`, the user should provide the paths to the wallpaper slideshow folders in the following variables. These should be the same as indicated in the `[Slideshow]` section of the `.theme` file.

3) (optional) run the script `Setup.ps1` to create the main scheduled task. The script will ask for the user's permission to create the task, and it will then import the included `AutoTheme.xml` file, which contains all necessary settings of the task. 
The task will then run the script `AutoTheme.ps1` at logon, upon workstation unlock or when the system starts up. The script itself will then schedule the next temporary task to run at the next required theme change time.

4) (alternative) You can of course create the task yourself using Task Scheduler, setting the triggers to anything you prefer. In this case, make sure that the Action is set up as follows:
	- Program/script: `Powershell.exe`
	- Add arguments: `-NonInteractive -ExecutionPolicy Bypass -NoProfile -File "C:\path\to\AutoTheme.ps1"`
	- Run with highest privileges.

![image](https://github.com/user-attachments/assets/4ec93663-3001-46ce-ad56-b16a623de8b1)

![image](https://github.com/user-attachments/assets/048e6e91-fe0e-4bf0-905c-3beb2aeb4385)

![image](https://github.com/user-attachments/assets/f2dcbd94-beee-477d-8f5c-5868c3780dc0)


## Usage
This script is designed to run from Task Scheduler, and after the initial setup doesn't need interaction from the user. When run from terminal, however, using `.\AutoTheme.ps1`, the script will simply 'toggle' the theme (switching from one to the other) and then exit. This can be useful for testing purposes, but also for the odd times when we need to manually switch the theme regardless of our hours settings. You can create a shortcut to the script and place it on your desktop or taskbar for quick access. In this case, the shortcut should be to `powershell.exe` followed by the path to the script `"C:\path\to\AutoTheme.ps1"`, indicating the same path in the `Start in` field:

![image](https://github.com/user-attachments/assets/954d4a76-3001-4bd8-9e7d-460c1db3888a)



