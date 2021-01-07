<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Update all the relevant Microsoft Online PowerShell modules

Source - https://github.com/directorcia/Office365/blob/master/o365-update.ps1

## Prerequisites = 1
## 1. Run PowerShell environment as an administrator

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

Function test-install($modulename) {
    if (Get-Module -ListAvailable -Name $modulename) {          ## If module exists then update
        update-module -name $modulename -force
    } 
    else {                                                      ## If module doesn't exist then prompt to update
        do {
            write-host -foregroundcolor $warningmessagecolor -nonewline "    [Warning]"$modulename" module not found. "
            $result = Read-host -prompt "Install this module (Y/N)?"
        } until (-not [string]::isnullorempty($result))
        if ($result -eq 'Y' -or $result -eq 'y') {
            write-host -foregroundcolor $processmessagecolor "Installing module",$modulename
            install-Module -Name $modulename -Force
        }
    }
}

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Start Script`n"

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    write-host -foregroundcolor $processmessagecolor "Update Azure AD module"
    test-install -modulename AzureAD
    write-host -foregroundcolor $processmessagecolor "Update Azure Information Protection module"
    test-install -modulename AIPService
    write-host -foregroundcolor $processmessagecolor "Update Teams Module"
    test-install -modulename MicrosoftTeams
    write-host -foregroundcolor $processmessagecolor "Update SharePoint Online module"
    test-install -modulename Microsoft.Online.SharePoint.PowerShell
    write-host -foregroundcolor $processmessagecolor "Update Microsoft Online module"
    test-install -modulename MSOnline
    write-host -foregroundcolor $processmessagecolor "Update Exchange Online V2 module"
    test-install -modulename PowershellGet
    test-install -modulename ExchangeOnlineManagement
    write-host -foregroundcolor $processmessagecolor "Update Azure module"
    test-install -modulename Az 
    write-host -foregroundcolor $processmessagecolor "Update SharePoint PnP module"
    test-install -modulename SharePointPnPPowerShellOnline
    write-host -foregroundcolor $processmessagecolor "Update Microsoft Graph module"
    test-install -modulename Microsoft.Graph 
    write-host -foregroundcolor $processmessagecolor "Update Intune Module"
    test-install -modulename Microsoft.Graph.Intune 
    write-host -foregroundcolor $processmessagecolor "Update Windows Autopilot Module"
    test-install -modulename WindowsAutoPilotIntune
}
Else {
    write-host -foregroundcolor $errormessagecolor "*** ERROR *** - Please re-run PowerShell environment as Administrator`n"
}
write-host -foregroundcolor $systemmessagecolor "Script completed`n"