<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into Microsoft Teams

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-tms.ps1

Prerequisites = 1
1. Ensure Micosoft Teams Module is install or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$savedcreds=$false                      ## false = manually enter creds, True = from file
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials if required

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

import-module MicrosoftTeams
write-host -foregroundcolor $processmessagecolor "Microsoft Teams module loaded"

## Get tenant login credentials
if ($savedcreds) {
    ## Get creds from local file
    $cred =import-clixml -path $credpath
}
else {
    ## Get creds manually
    $cred=get-credential
}

## Connect to Microsoft Teams service
Connect-MicrosoftTeams -credential $cred
write-host -foregroundcolor $processmessagecolor "Now connected to Microsoft Teams Service`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"