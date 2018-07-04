## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.


## Description
## Script designed to log into Microsoft Teams with MFA enabled

## Prerequisites = 1
## 1. Ensure Micosoft Teams Module is install or updated

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

## Connect to Microsoft Teams service
## You will be manually prompted to enter credentials and MFA
Connect-MicrosoftTeams
write-host -foregroundcolor green "Now connected to Microsoft Teams Service"