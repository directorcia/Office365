## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.


## Description
## Script designed to log into the Azure AD portal

## Prerequisites = 1
## 1. Ensure azuread module installed or updated

## Variables
$savedcreds=$false                      ## false = manually enter creds, True = from file
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials if required

Clear-Host

write-host -foregroundcolor green "Script started"

## set-executionpolicy remotesigned
## May be required once to allow ability to runs scripts in PowerShell

## ensure that install-module azuread has been run
## ensure that update-module azuread has been run to get latest module
## https://www.powershellgallery.com/packages/AzureAD/
## Current version = 2.0.1.16, 21 June 2018
import-module azuread
write-host -foregroundcolor green "AzureAD module loaded"

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
write-host -foregroundcolor green "Now connected to Azure AD Service"