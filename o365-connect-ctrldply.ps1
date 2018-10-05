## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into the Office 365 Centralized Deployment for add ins

## Source - https://github.com/directorcia/Office365/blob/master/o365-connect-ctrldply.ps1

## Prerequisites = 1
## 1. Ensure powershell cmdlest for Centralized deployment installed or updated

## Variables
$systemmessagecolor = "cyan"
$savedcreds=$false                      ## false = manually enter creds, True = from file
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials if required

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started"

## https://www.microsoft.com/en-us/download/details.aspx?id=55267
## Version 1.2.0.0 Date = 26 April 2018

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
write-host -foregroundcolor $systemmessagecolor "Now connected to Office 365 Centralized Deployment"