<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/o365-spo-getsharing.ps1

Description - Report on the sharing state of SharePoint Online sites

Prerequisites = 1
1. Ensure SharePoint online PowerShell module installed or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"
write-host

write-host -ForegroundColor white "Sharing Capability Settings"
Write-Host -ForegroundColor White "----------------------------------------"
write-host -ForegroundColor white "* Disabled - external user sharing (share by email) and guest link sharing both disabled"
write-host -ForegroundColor white "* ExternalUserSharingOnly - external user sharing (share by email) enabled, but guest link sharing disabled"
write-host -ForegroundColor white "* ExistingExternalUserSharingOnly - (DEFAULT) Allow sharing only with external users that already exist in organisation’s directory"
write-host -ForegroundColor white "* ExternalUserAndGuestSharing - external user sharing (share by email) and guest link sharing both enabled"
Write-Host

## ensure that SharePoint Online module has been installed and loaded

Write-host -ForegroundColor $processmessagecolor "Getting all Sharepoint sites in tenant"

get-sposite | Select-object url,sharingcapability

write-host -foregroundcolor $systemmessagecolor "Script completed`n"
