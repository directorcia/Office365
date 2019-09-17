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

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Start Script`n"

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    write-host -foregroundcolor $processmessagecolor "Update Azure AD module"
    Update-Module -Name AzureAD -force
    write-host -foregroundcolor $processmessagecolor "Update Azure AD Right Management module"
    Update-Module -Name AADRM -force
    write-host -foregroundcolor $processmessagecolor "Update Teams Module"
    Update-Module -Name MicrosoftTeams -Force
    write-host -foregroundcolor $processmessagecolor "Update SharePoint Online module"
    Update-Module -Name Microsoft.Online.SharePoint.PowerShell -force
    write-host -foregroundcolor $processmessagecolor "Update Microsoft Online module"
    Update-Module -Name MSOnline -force
    write-host -foregroundcolor $processmessagecolor "Update Azure module"
    ## Old Azure module
    ## Update-Module -name AzureRM -Force
    ## New Az module
    Update-Module -name Az -force
    write-host -foregroundcolor $processmessagecolor "Update SharePoint PnP module"
    update-Module SharePointPnPPowerShellOnline -Force
}
Else {
    write-host -foregroundcolor $errormessagecolor "*** ERROR *** - Please re-run PowerShell environment as Administrator`n"
}
write-host -foregroundcolor $systemmessagecolor "Script completed`n"