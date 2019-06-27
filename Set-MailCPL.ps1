<#
.SYNOPSIS
    Fix "application not found" error running the Control Panel Mail settings app

.DESCRIPTION
    This script resolves the "application not found" error running the Control Panel Mail settings app on a Windows 10 build 1903 PC with Office 365 installed locally.

.EXAMPLE
    .\Set-MailCPL.ps1

#>

[cmdletbinding()]
Param()

# Only create HKCR PSDrive if it doesn't already exist
if (-not (Get-PSDrive HKCR -ErrorAction SilentlyContinue)) {
    New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR
}


# Create arrays with the registry keys needed for the Control Panel Mail settings app
# Note: These keys are valid on Windows 10 1903, other OSes/OS builds may use different keys
$AppKeyPathArray = "HKLM:\SOFTWARE\WOW6432Node\Classes\CLSID\{A0D4CD32-5D5D-4f72-BAAA-767A7AD6BAC5}\shell\open\command", "HKCR:\WOW6432Node\CLSID\{A0D4CD32-5D5D-4f72-BAAA-767A7AD6BAC5}\shell\open\command", "HKCR:\CLSID\{A0D4CD32-5D5D-4f72-BAAA-767A7AD6BAC5}\shell\open\command"
$IconKeyPathArray = "HKLM:\SOFTWARE\WOW6432Node\Classes\CLSID\{A0D4CD32-5D5D-4f72-BAAA-767A7AD6BAC5}\DefaultIcon", "HKCR:\WOW6432Node\CLSID\{A0D4CD32-5D5D-4f72-BAAA-767A7AD6BAC5}\DefaultIcon", "HKCR:\CLSID\{A0D4CD32-5D5D-4f72-BAAA-767A7AD6BAC5}\DefaultIcon"

# Set correct path values to use with app and icon keys
# Note: These paths are valid on Windows 10 1903 with Office 365 32-bit Desktop installed, other OSes/OS builds/Office versions may use different paths
$AppString = '"C:\Program Files (x86)\Microsoft Office\root\Client\AppVLP.exe" rundll32.exe shell32.dll,Control_RunDLL "C:\Program Files (x86)\Microsoft Office\root\Office16\MLCFG32.CPL"'
$IconString = "C:\Program Files (x86)\Microsoft Office\root\Office16\MLCFG32.CPL,0"

# Loop through app related registry keys and set correct values
ForEach ($AppArrayElement in $AppKeyPathArray) {
    try {
        Set-ItemProperty -LiteralPath $AppArrayElement -name '(default)' -Value $AppString -ErrorAction Stop
        Write-Host "Successfully set key $AppArrayElement"
    }
    catch {
        Write-Host "Error setting key $AppArrayElement"
    }
}

# Loop through icon related registry keys and set correct values
ForEach ($IconArrayElement in $IconKeyPathArray) {
    try {
        Set-ItemProperty -LiteralPath $IconArrayElement -name '(default)' -Value $IconString -ErrorAction Stop
        Write-Host "Successfully set key $IconArrayElement"
    }
    catch {
        Write-Host "Error setting key $IconArrayElement"
    }
}