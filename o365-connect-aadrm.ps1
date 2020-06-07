<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the Azure AD Rights Management module with MFA enabled

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-aadrm.ps1

Prerequisites = 1
1. Ensure AADRM module is installeed or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## Connect to Azure AD Rights Management Service
connect-aadrmservice
write-host -foregroundcolor $processmessagecolor "Now connected to the Azure AD Rights Management Service`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"