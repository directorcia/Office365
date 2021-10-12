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

write-host -foregroundcolor $systemmessagecolor "Start Script`n"

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    <# Test for Execution policy #>
    write-host -ForegroundColor $processmessagecolor -nonewline "Checking PowerShell Execution Policy "
    $result = Get-ExecutionPolicy -Scope CurrentUser
    if ($result -ne "Bypass") {
        write-host -ForegroundColor $warningmessagecolor "[Warning] - Execution policy for Currentuser not set to bypass to allow scripts to run"
        write-host -ForegroundColor $processmessagecolor "Setting Powershell execution policy for Currentuser to bypass to allow scripts to run"        
        set-executionpolicy -executionpolicy bypass -scope currentuser -force
    }
    else {
        write-host -ForegroundColor $processmessagecolor "Execution policy for Currentuser is set to bypass to allow scripts to run"
    }
    write-host -foregroundcolor $processmessagecolor "NuGet provider"
    Install-PackageProvider -Name NuGet -Force -confirm:$false | Out-Null
    write-host -foregroundcolor $processmessagecolor "Install Azure AD module"
    Install-Module -Name AzureAD -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install Azure Information Protection module"
##    Install-Module -Name AADRM -force                       ## Support for the AADRM module ends July 15, 2020
    $aadrmcheck = get-module -listavailable -name aadrm
    if ($aadrmcheck) {
        write-host -foregroundcolor $warningmessagecolor "[Warning] Older module Azure AD Rights management module (AADRM) is still installed"
        write-host -foregroundcolor $processmessagecolor "Uninstalling AADRM module as support ended July 15, 2020 "
        uninstall-module aadrm -force -confirm:$false
        write-host -foregroundcolor $processmessagecolor "New Azure Information Protection module will now be installed"
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
    $pnpcheck = get-module -listavailable -name SharePointPnPPowerShellOnline
    if ($pnpcheck) {
        write-host -foregroundcolor $warningmessagecolor "[Warning] Older SharePoint PnP module is still installed"
        write-host -foregroundcolor $processmessagecolor "Uninstalling older SharePoint PnP module"
        uninstall-module SharePointPnPPowerShellOnline -allversions -force -confirm:$false
        write-host -foregroundcolor $processmessagecolor "New SharePoint PnP module will now be installed"
    }
    Install-Module PnP.PowerShell -Force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install Microsoft Graph Module"
    Install-Module -Name Microsoft.Graph -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install Windows Autopilot Module"
    ## will also update dependent AzureAD and Microsoft.Graph.Intune modules
    Install-Module -Name WindowsAutoPilotIntune -force -confirm:$false
    write-host -foregroundcolor $processmessagecolor "Install Centralised Add-in Deployment"
    Install-module -name O365CentralizedAddInDeployment -force -confirm:$false
}
else {
    write-host -foregroundcolor $errormessagecolor "*** ERROR *** - Please re-run PowerShell environment as Administrator`n"
}


write-host -foregroundcolor $systemmessagecolor "Script completed`n"