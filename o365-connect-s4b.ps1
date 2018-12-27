## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into the Skype for Business Online portal

## Source - https://github.com/directorcia/Office365/blob/master/o365-connect-s4b.ps1

## Prerequisites = 2
## 1. Ensure msonline module installed or updated
## 2. Ensure Skype for Business online PowerShell module installed or updated

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$savedcreds=$false                      ## false = manually enter creds, True = from file
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials if required

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started"

## ensure that install-module msonline has been run
## ensure that update-module msonline has been run to get latest module

## https://www.powershellgallery.com/packages/MSOnline/
## Current version = 1.1.183.8, 18 May 2018

clear-host

import-module msonline
write-host -foregroundcolor $processmessagecolor "MSOnline module loaded"

## Download and install https://www.microsoft.com/en-au/download/details.aspx?id=39366 (Skype for Business Online Module)
## Current version = 7.0.1994.0, 26 February 2018
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