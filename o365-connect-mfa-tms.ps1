<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into Microsoft Teams with MFA enabled

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-tms.ps1

Prerequisites = 1
1. Ensure Micosoft Teams Module is install or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

import-module MicrosoftTeams
write-host -foregroundcolor $processmessagecolor "Microsoft Teams module loaded"

## Connect to Microsoft Teams service
## You will be manually prompted to enter credentials and MFA
Connect-MicrosoftTeams
write-host -foregroundcolor $processmessagecolor "Now connected to Microsoft Teams Service`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"