<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the Skype for Business Online portal

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-s4b.ps1

Prerequisites = 2
1. Ensure msonline module installed or updated
2. Ensure Skype for Business online PowerShell module installed or updated

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

import-module msonline
write-host -foregroundcolor $processmessagecolor "MSOnline module loaded"

import-module skypeonlineconnector
write-host -foregroundcolor $processmessagecolor "Skype for Business module loaded"

## Get tenant login credentials
if ($savedcreds) {
    ## Get creds from local file
    $cred =import-clixml -path $credpath
}
else {
    ## Get creds manually
    $cred=get-credential
}

## Connect to Office 365 admin service
connect-msolservice -credential $cred
write-host -foregroundcolor $processmessagecolor "Now connected to Office 365 Admin service"

## Connect to Skype for Business Online Service
$sfboSession=new-csonlinesession -credential $cred
import-pssession $sfboSession
write-host -foregroundcolor $processmessagecolor "Now connected to Skype for Business Online services`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"