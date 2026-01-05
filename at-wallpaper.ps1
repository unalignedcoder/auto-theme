# Worker script to handle wallpaper changes for the Auto Theme script, `at.ps1`
param (
    [string]$Path
)

# Never remove my original comments from code.
$ConfigPath = Join-Path $PSScriptRoot "at-config.ps1"
if (Test-Path $ConfigPath) { . $ConfigPath }

# ============= Execution ==============

[string]$cleanName = ""
[string]$targetFile = ""

# Restore terminal visibility for the worker process
Write-Information "Wallpaper Worker: Processing path $Path" -InformationAction Continue

if (Test-Path $Path) {
    # 1. Selection Logic (Folders or Files)
    if (Test-Path $Path -PathType Container) {
        $files = Get-ChildItem -Path $Path -File | Where-Object { $_.Extension -match "jpg|jpeg|png|bmp" }
        if ($files) {
            $target = if ($randomizeWallpaper) { $files | Get-Random } else { $files | Sort-Object Name | Select-Object -First 1 }
            $targetFile = $target.FullName
        }
    } elseif ((Get-Item $Path).Extension -match "jpg|jpeg|png|bmp") {
        $targetFile = $Path
    }

    # 2. Apply Wallpaper
    if ($targetFile) {
        # WinAPI is now provided by at-config.ps1
        $null = [WinAPI]::SystemParametersInfo(0x0014, 0, $targetFile, 0x01 -bor 0x02)
        
        $cleanName = [System.IO.Path]::GetFileNameWithoutExtension($targetFile)
        $cleanName = $cleanName -replace '^_0_at_', ''
        
        Write-Information "Wallpaper Worker: Applied $cleanName" -InformationAction Continue
    }

    # 3. Registry Update
    if ($writeRegistry -or $howToDisplayName -eq "rainmeter") {
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallpaperName" -Value $cleanName -Force | Out-Null
    }

    # 4. UI Feedback Logic
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