<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into Azure AD with MFA enabled

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-aad.ps1

Prerequisites = 2
1. Ensure azuread module installed or updated
2. Ensure msonline MFA module installed or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

import-module azuread
write-host -foregroundcolor $processmessagecolor "AzureAD module loaded"

## Connect to Azure AD service
## You will be manually prompted to enter credentials and MFA
Connect-AzureAD
write-host -foregroundcolor $processmessagecolor "Now connected to Azure AD Service with MFA`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"