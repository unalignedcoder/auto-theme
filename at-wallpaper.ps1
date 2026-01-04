# Worker script to handle wallpaper changes for the Auto Theme script, `at.ps1`

param (
    [string]$Path
)

# Never remove my original comments from code.
$ConfigPath = Join-Path $PSScriptRoot "at-config.ps1"
if (Test-Path $ConfigPath) { . $ConfigPath }

# ============= Win32 API Definition ==============
$Win32Code = @"
using System;
using System.Runtime.InteropServices;
public class WinAPI {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@
if (-not ([System.Management.Automation.PSTypeName]'WinAPI').Type) {
    Add-Type -TypeDefinition $Win32Code | Out-Null
}

# ============= Execution ==============
[string]$cleanName = ""
[string]$targetFile = ""

Write-Information "Wallpaper Worker: Processing path $Path" -InformationAction Continue

if (Test-Path $Path) {
    
    # 1. Handle Folders
    if (Test-Path $Path -PathType Container) {
        $files = Get-ChildItem -Path $Path -File | Where-Object { $_.Extension -match "jpg|jpeg|png|bmp" }
        if ($files) {
            $target = if ($randomizeWallpaper) { $files | Get-Random } else { $files | Sort-Object Name | Select-Object -First 1 }
            $targetFile = $target.FullName
        }
    } 
    # 2. Handle Fixed Files
    elseif ((Get-Item $Path).Extension -match "jpg|jpeg|png|bmp") {
        $targetFile = $Path
    }

    # 3. Apply if a valid file was found
    if ($targetFile) {
        $null = [WinAPI]::SystemParametersInfo(0x0014, 0, $targetFile, 0x01 -bor 0x02)
        
        $cleanName = [System.IO.Path]::GetFileNameWithoutExtension($targetFile)
        $cleanName = $cleanName -replace '^_0_at_', ''
        Write-Information "Wallpaper Worker: Applied $cleanName" -InformationAction Continue
    }

    # ============= UI Feedback Logic ==============
    
    # Update Registry for Rainmeter/Meta purposes
    if ($writeRegistry -or $howToDisplayName -eq "rainmeter") {
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallpaperName" -Value $cleanName -Force | Out-Null
    }

    switch ($howToDisplayName) {
        "rainmeter" {
            if (Test-Path $rainmeterPath) {
                Write-Information "Wallpaper Worker: Activating/Refreshing $rainmeterSkinName" -InformationAction Continue
                Start-Process $rainmeterPath -ArgumentList "!ActivateConfig `"$rainmeterSkinName`"", "!Refresh `"$rainmeterSkinName`"" -WindowStyle Hidden | Out-Null
            }
        }
        "notification" {
            if ($cleanName -and (Get-Module -Name BurntToast -ListAvailable)) {
                New-BurntToastNotification -Text "New Wallpaper:", $cleanName | Out-Null
            }
            # Cleanup: Kill the overlay if user switched to notifications
            if (Test-Path $rainmeterPath) {
                Start-Process $rainmeterPath -ArgumentList "!DeactivateConfig `"$rainmeterSkinName`"" -WindowStyle Hidden | Out-Null
            }
        }
        "none" {
            # Cleanup: Kill the overlay if user chose none
            if (Test-Path $rainmeterPath) {
                Start-Process $rainmeterPath -ArgumentList "!DeactivateConfig `"$rainmeterSkinName`"" -WindowStyle Hidden | Out-Null
            }
        }
    }

} else {
    Write-Information "Wallpaper Worker: Path not found!" -InformationAction Continue
}

return $cleanName