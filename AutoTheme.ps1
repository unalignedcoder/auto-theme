<#
    LEGACY WRAPPER
    This file exists for backward compatibility with existing Scheduled Tasks and shortcuts.
    It redirects all calls to the primary script: at.ps1
#>
& "$PSScriptRoot\at.ps1" @args