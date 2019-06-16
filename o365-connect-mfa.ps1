<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the Office 365 admin portal using MFA

## Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa.ps1

Prerequisites = 1
1. Ensure msonline module installed or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

import-module msonline
write-host -foregroundcolor $processmessagecolor "MSOnline module loaded"

## Connect to Office 365 admin service
connect-msolservice
write-host -foregroundcolor $processmessagecolor "Now connected to Office 365 Admin service`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"