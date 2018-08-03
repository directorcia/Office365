## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Source -

## Description
## Script designed to report onthe sharing state of SharePoint Online sites

## Prerequisites = 1
## 1. Ensure SharePoint online PowerShell module installed or updated

## Variables

Clear-Host

write-host -foregroundcolor Cyan "Script started"
write-host

write-host -ForegroundColor white "Sharing Capability Settings"
Write-Host -ForegroundColor White "----------------------------------------"
write-host -ForegroundColor white "* Disabled – external user sharing (share by email) and guest link sharing are both disabled"
write-host -ForegroundColor white "* ExternalUserSharingOnly – external user sharing (share by email) is enabled, but guest link sharing is disabled"
write-host -ForegroundColor white "* ExistingExternalUserSharingOnly - (DEFAULT) Allow sharing only with the external users that already exist in your organization’s directory"
write-host -ForegroundColor white "* ExternalUserAndGuestSharing - external user sharing (share by email) and guest link sharing are both enabled"
Write-Host

## ensure that SharePoint Online module has been installed and loaded

Write-host -ForegroundColor Cyan "Getting all Sharepoint sites in tenant"

get-sposite | Select-object url,sharingcapability

write-host -foregroundcolor Cyan "Script complete"
