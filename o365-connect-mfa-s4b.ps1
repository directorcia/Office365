<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the Skype for Business Online portal with multi factor enabled

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-s4b.ps1

Prerequisites = 2
1. Ensure Skype for Business online PowerShell module installed or updated
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

import-module skypeonlineconnector
write-host -foregroundcolor $processmessagecolor "Skype for Business module loaded"

## Connect to Skype for Business Online Service
## You will be manually prompted to enter your credentials and MFA
$sfboSession=new-csonlinesession
import-pssession $sfboSession
write-host -foregroundcolor $processmessagecolor "Now connected to Skype for Business Online services`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"