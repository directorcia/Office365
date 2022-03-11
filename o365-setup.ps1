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
$warningmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host
if ((get-host).name -match "ISE") {Write-host "`n`n`n`n`n`n`n`n`n"}     # move text down for ISE where dialog will hide
write-host -foregroundcolor $systemmessagecolor "Start Script`n"

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    write-host -ForegroundColor $processmessagecolor "Checking PowerShell Execution Policy"
    $result = Get-ExecutionPolicy -Scope CurrentUser
    if ($result -ne "Bypass") {
        write-host -ForegroundColor $warningmessagecolor " - [Warning] - Currentuser not set to bypass to allow scripts to run"
        write-host -ForegroundColor $processmessagecolor " - Setting Currentuser to bypass to allow scripts to run"        
        set-executionpolicy -executionpolicy bypass -scope currentuser -force
    }
    else {
        write-host -ForegroundColor $processmessagecolor " - Currentuser is set to bypass to allow scripts to run"
    }
    write-host -foregroundcolor $processmessagecolor "`n(1 of 16) Install NuGet provider"
    Install-PackageProvider -Name NuGet -Force -confirm:$false | Out-Null
    write-host -foregroundcolor $processmessagecolor "(2 of 16) Install Azure AD module"
    Install-Module -Name AzureAD -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "(3 of 16) Install Azure Information Protection module"
##    Install-Module -Name AADRM -force                       ## Support for the AADRM module ends July 15, 2020
    $aadrmcheck = get-module -listavailable -name aadrm
    if ($aadrmcheck) {
        write-host -foregroundcolor $warningmessagecolor " - [Warning] Older module Azure AD Rights management module (AADRM) is still installed"
        write-host -foregroundcolor $processmessagecolor " - Uninstalling AADRM module as support ended July 15, 2020 "
        uninstall-module aadrm -force -confirm:$false
        write-host -foregroundcolor $processmessagecolor " - New Azure Information Protection module will now be installed"
    }
    Install-module -name aipservice -Force -confirm:$false

    write-host -foregroundcolor $processmessagecolor "(4 of 16) Install Teams Module"
    Install-Module -Name MicrosoftTeams -Force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "(5 of 16) Install SharePoint Online module"
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "(6 of 16) Install Microsoft Online module"
    Install-Module -Name MSOnline -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "(7 of 16) Install PowerShell Get module"
    Install-Module -Name PowerShellGet -force -confirm:$false -allowclobber
    write-host -foregroundcolor $processmessagecolor "(8 of 16) Install Exchange Online module"
    Install-Module -Name ExchangeOnlineManagement -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "(9 of 16) Install Azure module"
    ## Old Azure module
    ## Install-Module -name AzureRM -Force
    ## New Az module
    Install-Module -name Az -force -allowclobber -confirm:$false
    write-host -foregroundcolor $processmessagecolor "(10 of 16) Install SharePoint PnP module"
    $pnpcheck = get-module -listavailable -name SharePointPnPPowerShellOnline
    if ($pnpcheck) {
        write-host -foregroundcolor $warningmessagecolor " - [Warning] Older SharePoint PnP module is still installed"
        write-host -foregroundcolor $processmessagecolor " - Uninstalling older SharePoint PnP module"
        uninstall-module SharePointPnPPowerShellOnline -allversions -force -confirm:$false
        write-host -foregroundcolor $processmessagecolor " - New SharePoint PnP module will now be installed"
    }
    Install-Module PnP.PowerShell -Force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "(11 of 16) Install Microsoft Graph Module"
    Install-Module -Name Microsoft.Graph -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "(12 of 16) Install Windows Autopilot Module"
    ## will also update dependent AzureAD and Microsoft.Graph.Intune modules
    Install-Module -Name WindowsAutoPilotIntune -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "(13 of 16) Install Centralised Add-in Deployment"
    Install-module -name O365CentralizedAddInDeployment -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "(14 of 16) Install PowerApps Module"
    install-module -name Microsoft.PowerApps.PowerShell -force -confirm:$false -AllowClobber
    write-host -foregroundcolor $processmessagecolor "(15 of 16) Install PowerApps Administration Module"
    install-module -name Microsoft.PowerApps.Administration.PowerShell -force -confirm:$false -allowclobber
    write-host -foregroundcolor $processmessagecolor "(16 of 16) Install Microsoft 365 Commerce Module"
    install-module -name MSCommerce -force -confirm:$false
}
else {
    write-host -foregroundcolor $errormessagecolor "*** ERROR *** - Please re-run PowerShell environment as Administrator`n"
}

write-host -foregroundcolor $systemmessagecolor "Script completed`n"