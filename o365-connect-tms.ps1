## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.


## Description
## Script designed to log into Microsoft Teams

## Prerequisites = 1
## 1. Ensure Micosoft Teams Module is install or updated

## Variables
$savedcreds=$false                      ## false = manually enter creds, True = from file
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials if required

Clear-Host

write-host -foregroundcolor green "Script started"

## set-executionpolicy remotesigned
## May be required once to allow ability to runs scripts in PowerShell

## ensure that install-module -name microsoftteams has been run
## ensure that update-module -name microsoftteams has been run to get latest module
## https://www.powershellgallery.com/packages/MicrosoftTeams/
## Current version = 0.9.3, 25 April 2018
import-module MicrosoftTeams
write-host -foregroundcolor green "Microsoft Teams module loaded"

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
write-host -foregroundcolor green "Now connected to Microsoft Teams Service"