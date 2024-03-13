<#
.SYNOPSIS
    Script to install or uninstall latest version of Zoom via PowerShell. This script can be packaged as win32 app and then deployed via intune.

.DESCRIPTION
    This script allows you to install or uninstall Zoom on a Windows system.
    
.NOTES
    Author: Lovepreet Singh
    Date: March 13, 2024
    Version: 1.0
    
    This script is provided as-is and without warranty. Use it at your own risk.

.Usage
    Copy and paste the below commands in Command prompt (run as admin), Or in Intune

    For Install:> Powershell.exe -NoProfile -ExecutionPolicy ByPass -File .\Install_Zoom.ps1 --install
    For Uninstall:> Powershell.exe -NoProfile -ExecutionPolicy ByPass -File .\Install_Zoom.ps1 --uninstall
    
#>

param (
    [switch]$Install,
    [switch]$Uninstall
)

#Give a app name and specify the permalink
$AppName = "Zoom"
$DownloadURL = "https://zoom.us/client/latest/ZoomInstallerFull.msi?archType=x64"

# Specify the full path to the MSI file. In my case i am storing this in the Temp folder.
$MSIFilePath = "$env:TEMP\$AppName.msi"

if ($Install) {
    $InstallCommand = "msiexec /i '$MSIFilePath' /qn"

     
    # Suppress progress reporting                           #By suppressing the progress bar, the download speed increases 10x. This is a global variable used by powershell itself.
    $ProgressPreference = 'SilentlyContinue'     

    # Download the MSI file
    Invoke-WebRequest -Uri $DownloadURL -OutFile $MSIFilePath

    # Install Zoom silently and enable Zoom auto updates
    Start-Process -FilePath "msiexec" -ArgumentList "/i `"$MSIFilePath`" /qn /lex zoommsi.log ZoomAutoUpdate=1" -Wait

    # Remove the downloaded MSI file
    Remove-Item -Path $MSIFilePath -Force
}

#Below is the Uninstall parameter of this script. At first i wanted to copy the Install parameter and just replace the /i with /x to uninstall. But this is a bad idea. Why?
# Maybe the user will uninstall after 6 months, by then Zoom might have a new version with new MSI, with a differnt msi product code, so the uninstall can fail, instead we will below mentioned method

elseif ($Uninstall) {

    $Query = "SELECT * FROM Win32_Product WHERE Name LIKE '%$AppName%'"

    # Query for products that match the criteria
    $Product = Get-WmiObject -Query $Query | Select-Object -ExpandProperty IdentifyingNumber

    

    # Un-Install Zoom silently
    Start-Process -FilePath "msiexec" -ArgumentList "/x $Product /qn" -Wait

}
