## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into the Azure AD portal

## Source - https://github.com/directorcia/Office365/blob/master/O365-connect-aad.ps1

## Prerequisites = 1
## 1. Ensure azuread module installed or updated

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$savedcreds=$false                      ## false = manually enter creds, True = from file
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials if required

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## ensure that install-module azuread has been run
## ensure that update-module azuread has been run to get latest module
## https://www.powershellgallery.com/packages/AzureAD/
## Current version = 2.0.1.16, 21 June 2018
import-module azuread
write-host -foregroundcolor $processmessagecolor "AzureAD module loaded"

## Get tenant login credentials
if ($savedcreds) {
    ## Get creds from local file
    $cred =import-clixml -path $credpath
}
else {
    ## Get creds manually
    $cred=get-credential
}

## Connect to AzuerAD service
Connect-azuread -credential $cred
write-host -foregroundcolor $processmessagecolor "Now connected to Azure AD Service`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"