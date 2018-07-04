## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into the Skype for Business Online portal

## Prerequisites = 2
## 1. Ensure msonline module installed or updated
## 2. Ensure Skype for Business online PowerShell module installed or updated

## Variables
$savedcreds=$false                      ## false = manually enter creds, True = from file
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials if required

Clear-Host

write-host -foregroundcolor green "Script started"

## set-executionpolicy remotesigned
## May be required once to allow ability to runs scripts in PowerShell

## ensure that install-module msonline has been run
## ensure that update-module msonline has been run to get latest module
import-module msonline
write-host -foregroundcolor green "MSOnline module loaded"

## Download and install https://www.microsoft.com/en-au/download/details.aspx?id=39366 (Skype for Business Online Module)
## Current version = 7.0.1994.0, 26 February 2018
import-module skypeonlineconnector
write-host -foregroundcolor green "Skype for Business module loaded"

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
write-host -foregroundcolor green "Now connected to Office 365 Admin service"

## Connect to Skype for Business Online Service
$sfboSession=new-csonlinesession -credential $cred
import-pssession $sfboSession
write-host -foregroundcolor green "Now connected to Skype for Business Online services"