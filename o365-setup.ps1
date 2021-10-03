<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description - Install the relevant Microsoft Online PowerShell modules

Source - https://github.com/directorcia/Office365/blob/master/o365-setup.ps1

Prerequisites = 1
1. Run PowerShell environment as an administrator

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Start Script`n"

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    write-host -foregroundcolor $processmessagecolor "NuGet provider"
    Install-PackageProvider -Name NuGet -Force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install Azure AD module"
    Install-Module -Name AzureAD -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install Azure Information Protection module"
##    Install-Module -Name AADRM -force                       ## Support for the AADRM module ends July 15, 2020
    $aadrmcheck = get-module -listavailable -name aadrm
    if ($aadrmcheck) {
        write-host -foregroundcolor $processmessagecolor "Older module Azure AD Rights management module (AADRM) is installed"
        write-host -foregroundcolor $processmessagecolor "Uninstalling AADRM module as support ended July 15, 2020 "
        uninstall -module aadrm -force -confirm:$false
        write-host -foregroundcolor $processmessagecolor "Now Azure Information Protection module will be installed"
    }
    Install-module -name aipservice -Force -confirm:$false

    write-host -foregroundcolor $processmessagecolor "Install Teams Module"
    Install-Module -Name MicrosoftTeams -Force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install SharePoint Online module"
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install Microsoft Online module"
    Install-Module -Name MSOnline -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install Exchange Online module"
    Install-Module PowershellGet -Force -confirm:$false
    Install-Module -Name ExchangeOnlineManagement -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install Azure module"
    ## Old Azure module
    ## Install-Module -name AzureRM -Force
    ## New Az module
    Install-Module -name Az -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install SharePoint PnP module"
    Install-Module PnP.PowerShell -Force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install Microsoft Graph Module"
    Install-Module -Name Microsoft.Graph -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install Windows Autopilot Module"
    ## will also update dependent AzureAD and Microsoft.Graph.Intune modules
    Install-Module -Name WindowsAutoPilotIntune -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install Centralised Add-in Deployment"
    Install-module -name O365CentralizedAddInDeployment -confirm:$false
}
else {
    write-host -foregroundcolor $errormessagecolor "*** ERROR *** - Please re-run PowerShell environment as Administrator`n"
}


write-host -foregroundcolor $systemmessagecolor "Script completed`n"