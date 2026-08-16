<# Worker script to handle wallpaper changes for the Auto Theme script, `at.ps1`.

This file has two jobs. Called with a `-Path`, it changes the wallpaper, which is what the
slideshow task and the 'Next wallpaper' menu entry do. Dot-sourced without one, it stops
after the definitions below and simply hands over the shared machinery: the Win32 API
declarations and the state storage functions. `at.ps1` and `at-setup.ps1` load it that way,
which is why that machinery lives here rather than in `at-config.ps1`, a file meant to hold
nothing but settings. #>
param (
    [string]$Path,
    # Set by at.ps1 when the theme itself changes. Absent on slideshow ticks and on 'Next wallpaper'.
    [switch]$ThemeSwitch
)

# Never remove my original comments from code.
$ConfigPath = Join-Path $PSScriptRoot "at-config.ps1"
if (Test-Path $ConfigPath) { . $ConfigPath }

# ============= Shared Win32 API & Power Management, for wallpaper changing ==============

$Win32Code = @'
using System;
using System.Runtime.InteropServices;

public class WinAPI {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessageTimeout(
        IntPtr hWnd, uint Msg, IntPtr wParam, string lParam,
        uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);
}
'@

# Load the API class if not already present
if (-not ([System.Management.Automation.PSTypeName]'WinAPI').Type) {
    Add-Type -TypeDefinition $Win32Code -ErrorAction SilentlyContinue | Out-Null
}

# Apply Background Power State immediately upon loading
# This prevents any script using this file from blocking monitor/system sleep.
try {
    [WinAPI]::SetThreadExecutionState([System.UInt32]2147483648) | Out-Null
} catch { }

# ============= State storage ==============

    <# The script's own bookkeeping lives in one of two places, chosen by `$writeRegistry`.
    Rainmeter can only read the Registry, so asking for it forces the Registry backend
    whatever the preference says. #>
    $useRegistryState = $writeRegistry -or ($howToDisplayName -eq "rainmeter")
    $stateRegKey = "HKCU:\SOFTWARE\Tuscoss\AutoTheme"
    $stateFile = Join-Path $PSScriptRoot "at-state.ini"

    # Read one stored value, or "" when we have never written it
    function Get-AtState {
        param (
            [string]$Name
        )

        try {

            if ($useRegistryState) {

                return (Get-ItemProperty -Path $stateRegKey -Name $Name -ErrorAction Stop).$Name
            }

            if (-not (Test-Path $stateFile)) { return "" }

            foreach ($line in (Get-Content -LiteralPath $stateFile -ErrorAction Stop)) {

                <# Split on the FIRST '=' only, so that a value keeps any of its own.
                Note we cannot use ConvertFrom-StringData here: it reads backslashes as
                escape sequences, which would quietly mangle every path we store. #>
                $split = $line.IndexOf('=')
                if ($split -lt 1) { continue }

                if ($line.Substring(0, $split).Trim() -eq $Name) { return $line.Substring($split + 1).Trim() }
            }

            return ""

        } catch {

            return ""
        }
    }

    # Store one value. Never throws: bookkeeping must not break a wallpaper change.
    function Set-AtState {
        param (
            [string]$Name,
            [string]$Value
        )

        try {

            if ($useRegistryState) {

                if (-not (Test-Path $stateRegKey)) { New-Item -Path $stateRegKey -Force | Out-Null }
                Set-ItemProperty -Path $stateRegKey -Name $Name -Value $Value -Type String -Force | Out-Null
                return
            }

            # Read what is already there, replace our own line, and write the lot back
            $lines = @()
            if (Test-Path $stateFile) { $lines = @(Get-Content -LiteralPath $stateFile -ErrorAction Stop) }

            $output = @()
            $written = $false

            foreach ($line in $lines) {

                $split = $line.IndexOf('=')

                if ($split -ge 1 -and $line.Substring(0, $split).Trim() -eq $Name) {

                    # Replace in place, so the file keeps a stable order
                    $output += "$Name=$Value"
                    $written = $true

                } else {

                    $output += $line
                }
            }

            if (-not $written) { $output += "$Name=$Value" }

            Set-Content -LiteralPath $stateFile -Value $output -Force

        } catch {

            Write-Information "State: could not store $Name : $_" -InformationAction Continue
        }
    }

    <# Make sure the store exists from the very first run, before anything has been written to it.
    For the Registry that is a hard requirement: a Rainmeter Registry measure pointed at a missing
    key just logs an error and shows nothing. The file gets the same treatment for consistency,
    so that whichever backend is chosen, it is there to look at once setup has run.
    `$skipStateProvision` lets the uninstaller read these files for their settings without us
    quietly putting back the very store it is about to take away. #>
    if (-not $skipStateProvision) {

        if ($useRegistryState -and -not (Test-Path $stateRegKey)) {

            try {

                New-Item -Path $stateRegKey -Force | Out-Null
                Set-ItemProperty -Path $stateRegKey -Name "wallpaperName" -Value "" -Type String -Force | Out-Null

            } catch { }

        } elseif (-not $useRegistryState -and -not (Test-Path $stateFile)) {

            try {

                Set-Content -LiteralPath $stateFile -Value "wallpaperName=" -Force

            } catch { }
        }
    }

# ============= Lock screen functions ==============

    <# Windows Spotlight silently overwrites any lock screen picture we set,
    so we switch it off when the user allows us to. #>
    function Disable-LockScreenSpotlight {

        $cdmPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"

        try {

            if (-not (Test-Path $cdmPath)) { return }

            # Rotating lock screen pictures, and the 'fun facts' overlay that comes with them
            Set-ItemProperty -Path $cdmPath -Name "RotatingLockScreenEnabled" -Value 0 -Type DWord -Force | Out-Null
            Set-ItemProperty -Path $cdmPath -Name "RotatingLockScreenOverlayEnabled" -Value 0 -Type DWord -Force | Out-Null

            # The Spotlight content feed itself
            Set-ItemProperty -Path $cdmPath -Name "SubscribedContent-338387Enabled" -Value 0 -Type DWord -Force | Out-Null

        } catch {

            Write-Information "Lock Screen: Could not disable Spotlight: $_" -InformationAction Continue
        }
    }

    <# Set the lock screen picture through the WinRT API.
    This is the same call the Settings app makes, so it needs no Admin rights,
    but it only exists in Windows PowerShell 5.1, not in PowerShell 7. #>
    function Set-LockScreenImage {
        param (
            [string]$ImagePath
        )

        # PowerShell 7 cannot load the WinRT projections, so we bow out quietly
        if ($PSVersionTable.PSEdition -eq "Core") {
            Write-Information "Lock Screen: Skipped, requires Windows PowerShell 5.1." -InformationAction Continue
            return $false
        }

        if (-not (Test-Path $ImagePath)) { return $false }

        try {

            # The API insists on a full path
            $ImagePath = (Resolve-Path -LiteralPath $ImagePath).Path

            Add-Type -AssemblyName System.Runtime.WindowsRuntime

            <# WinRT calls are asynchronous. We fish the two 'AsTask' overloads out of the
            extension class, so that we can wait for the calls synchronously. #>
            $asTaskOp = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
                $_.Name -eq "AsTask" -and $_.GetParameters().Count -eq 1 -and
                $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]

            $asTaskAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
                $_.Name -eq "AsTask" -and $_.GetParameters().Count -eq 1 -and
                $_.GetParameters()[0].ParameterType.FullName -eq "Windows.Foundation.IAsyncAction" })[0]

            # Merely referencing the types is what loads the WinRT projections
            [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime] | Out-Null
            [Windows.System.UserProfile.LockScreen,Windows.System.UserProfile,ContentType=WindowsRuntime] | Out-Null

            # Turn the file path into the StorageFile the API expects
            $fileOp = [Windows.Storage.StorageFile]::GetFileFromPathAsync($ImagePath)
            $fileTask = $asTaskOp.MakeGenericMethod([Windows.Storage.StorageFile]).Invoke($null, @($fileOp))
            $fileTask.Wait(10000) | Out-Null
            $storageFile = $fileTask.Result

            if (-not $storageFile) {
                Write-Information "Lock Screen: Could not open $ImagePath" -InformationAction Continue
                return $false
            }

            # Apply it
            $setAction = [Windows.System.UserProfile.LockScreen]::SetImageFileAsync($storageFile)
            $setTask = $asTaskAction.Invoke($null, @($setAction))
            $setTask.Wait(15000) | Out-Null

            Write-Information "Lock Screen: Applied $([System.IO.Path]::GetFileName($ImagePath))" -InformationAction Continue
            return $true

        } catch {

            Write-Information "Lock Screen: Failed to set picture: $_" -InformationAction Continue
            return $false
        }
    }

# ============= Wallpaper selection ==============

    <# Which of the two wallpaper folders we are working on, so that Light and Dark
    each keep their own place in the sequence. #>
    function Get-WallpaperSlotName {
        param (
            [string]$FolderPath
        )

        $target = $FolderPath.TrimEnd('\', '/')

        <# Matching on the configured paths, rather than reading the current mode, keeps
        this independent of whether the theme registry flip happens before or after us. #>
        if ($target -eq $wallLightPath.TrimEnd('\', '/')) { return "lastLightWallpaperUsed" }
        if ($target -eq $wallDarkPath.TrimEnd('\', '/')) { return "lastDarkWallpaperUsed" }

        <# An unrecognised folder, passed by hand: fall back to whichever mode Windows
        reports as active, so we still keep a sensible place in the sequence. #>
        try {

            $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            $isLight = (Get-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -ErrorAction Stop).AppsUseLightTheme -eq 1

        } catch {

            $isLight = $true
        }

        if ($isLight) { return "lastLightWallpaperUsed" } else { return "lastDarkWallpaperUsed" }
    }

# ============= Execution ==============

<# Everything above is shared machinery. Without a `-Path` there is no wallpaper to change,
which means we were dot-sourced purely for the definitions, so we stop here and let the
caller get on with it. #>
if (-not $Path) { return }

[string]$cleanName = ""
[string]$targetFile = ""

# Restore terminal visibility for the worker process
Write-Information "Wallpaper Worker: Processing path $Path" -InformationAction Continue

if (Test-Path $Path) {

    # Which mode's place in the sequence we are keeping
    $slot = Get-WallpaperSlotName -FolderPath $Path

    # 1. Selection Logic (Folders or Files)
    if (Test-Path $Path -PathType Container) {
        $files = Get-ChildItem -Path $Path -File | Where-Object { $_.Extension -match "jpg|jpeg|png|bmp" }
        if ($files) {

            if ($randomizeWallpaper) {

                $target = $files | Get-Random

            } else {

                <# Sequential mode: step to the picture after the last one shown in this
                mode, so that Light and Dark each carry on where they left off. #>
                $sorted = @($files | Sort-Object Name)

                $last = Get-AtState -Name $slot
                $index = [Array]::FindIndex([object[]]$sorted, [Predicate[object]]{ param($f) $f.FullName -eq $last })

                # A picture we have no record of, or one since deleted, yields -1, which lands us on the first
                $target = $sorted[(($index + 1) % $sorted.Count)]
            }

            $targetFile = $target.FullName
        }
    } elseif ((Get-Item $Path).Extension -match "jpg|jpeg|png|bmp") {
        $targetFile = $Path
    }

    # 2. Apply Wallpaper
    if ($targetFile) {
        # WinAPI is declared at the top of this file
        $null = [WinAPI]::SystemParametersInfo(0x0014, 0, $targetFile, 0x01 -bor 0x02)
        
        $cleanName = [System.IO.Path]::GetFileNameWithoutExtension($targetFile)
        $cleanName = $cleanName -replace '^_0_at_', ''

        Write-Information "Wallpaper Worker: Applied $cleanName" -InformationAction Continue

        <# Record what we just applied. This has to happen before the Rainmeter bangs
        below, so that the refresh reads the new name rather than the old one. #>
        Set-AtState -Name $slot -Value $targetFile
        Set-AtState -Name "wallpaperPath" -Value $targetFile
        Set-AtState -Name "wallpaperName" -Value $cleanName

        <# 2b. Apply the same picture to the Lock Screen.
        Depending on the config, this either follows every wallpaper change,
        or only the ones caused by an actual theme switch. #>
        if ($changeLockScreen -and (-not $lockScreenOnThemeSwitchOnly -or $ThemeSwitch)) {

            if ($disableSpotlight) { Disable-LockScreenSpotlight }

            # Assign to $null: this script returns $cleanName, so nothing else may reach the pipeline
            $null = Set-LockScreenImage -ImagePath $targetFile
        }
    }

    # 3. UI Feedback Logic
    switch ($howToDisplayName) {
        "rainmeter" {
            if (Test-Path $rainmeterPath) {
                Write-Information "Wallpaper Worker: Activating Rainmeter skin: $rainmeterSkinFile" -InformationAction Continue
                
                # Bracketed syntax ensures Rainmeter processes the instructions reliably
                $bangs = "[!ActivateConfig `"$rainmeterSkinName`" `"$rainmeterSkinFile`"][!Refresh `"$rainmeterSkinName`"]"
                Start-Process $rainmeterPath -ArgumentList $bangs -WindowStyle Hidden | Out-Null
            }
        }
        "notification" {
            if ($cleanName -and (Get-Module -Name BurntToast -ListAvailable)) {
                New-BurntToastNotification -Text "New Wallpaper:", $cleanName | Out-Null
            }
            if (Test-Path $rainmeterPath) {
                Start-Process $rainmeterPath -ArgumentList "!DeactivateConfig `"$rainmeterSkinName`"" -WindowStyle Hidden | Out-Null
            }
        }
        "none" {
            if (Test-Path $rainmeterPath) {
                Start-Process $rainmeterPath -ArgumentList "!DeactivateConfig `"$rainmeterSkinName`"" -WindowStyle Hidden | Out-Null
            }
        }
    }
} else {
    Write-Information "Wallpaper Worker: Path not found!" -InformationAction Continue
}

return $cleanName