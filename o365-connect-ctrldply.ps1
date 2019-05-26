<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the Office 365 Centralized Deployment for add ins

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-ctrldply.ps1

Prerequisites = 1
1. Ensure powershell cmdlest for Centralized deployment installed or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$savedcreds = $false                      ## false = manually enter creds, True = from file
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials if required

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

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
Connect-OrganizationAddInService -credential $cred
write-host -foregroundcolor $processmessagecolor "Now connected to Office 365 Centralized Deployment`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"