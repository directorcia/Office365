## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Source -https://github.com/directorcia/Office365/blob/master/o365-spo-getsharing.ps1

## Description
## Script designed to report onthe sharing state of SharePoint Online sites

## Prerequisites = 1
## 1. Ensure SharePoint online PowerShell module installed or updated

## Variables
$systemmessagecolor = "cyan"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor Cyan "Script started"
write-host

write-host -ForegroundColor white "Sharing Capability Settings"
Write-Host -ForegroundColor White "----------------------------------------"
write-host -ForegroundColor white "* Disabled – external user sharing (share by email) and guest link sharing both disabled"
write-host -ForegroundColor white "* ExternalUserSharingOnly – external user sharing (share by email) enabled, but guest link sharing disabled"
write-host -ForegroundColor white "* ExistingExternalUserSharingOnly - (DEFAULT) Allow sharing only with external users that already exist in organisation’s directory"
write-host -ForegroundColor white "* ExternalUserAndGuestSharing - external user sharing (share by email) and guest link sharing both enabled"
Write-Host

## ensure that SharePoint Online module has been installed and loaded

Write-host -ForegroundColor Cyan "Getting all Sharepoint sites in tenant"

get-sposite | Select-object url,sharingcapability

write-host -foregroundcolor Cyan "Script complete"
