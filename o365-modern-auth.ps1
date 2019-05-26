<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Enables modern authentication for a tenant

## Source - https://github.com/directorcia/Office365/blob/master/o365-modern-auth.ps1

Prerequisites = 3
1. Ensure connected to Skype for Business online
2. Ensure connected to Exchange Online
3. Ensure connected to Sharepoint Online

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"
write-host
$org=get-organizationconfig
write-host -ForegroundColor $processmessagecolor "Exchange setting is currently",$org.OAuth2ClientProfileEnabled
## Run this command to enable modern authentication for Exchange Online
Set-OrganizationConfig -OAuth2ClientProfileEnabled $true
write-host -foregroundcolor $processmessagecolor "Exchange command completed"
$org=get-organizationconfig
write-host -ForegroundColor $processmessagecolor "Exchange setting updated to",$org.OAuth2ClientProfileEnabled
write-host
$org=get-csoauthconfiguration
write-host -ForegroundColor $processmessagecolor "Skype setting is currently",$org.clientadalauthoverride
## Run this command to enable modern authentication for Skype for Business Online
Set-CsOAuthConfiguration -ClientAdalAuthOverride Allowed
write-host -foregroundcolor $processmessagecolor "Skype command completed"
$org=get-csoauthconfiguration
write-host -ForegroundColor $processmessagecolor "Skype setting updated to",$org.clientadalauthoverride
write-host
## To set Sharepoint apps that don’t use modern authentication to block
$org=Get-SPOTenant
write-host -ForegroundColor $processmessagecolor "SharePoint setting is currently",$org.legacyauthprotocolsenabled
Set-spotenant -legacyauthprotocolsenabled $false
$org=Get-SPOTenant
write-host -ForegroundColor $processmessagecolor "SharePoint setting is updated to",$org.legacyauthprotocolsenabled
write-host -foregroundcolor $systemmessagecolor "Script completed`n"