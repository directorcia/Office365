## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into Microsoft Teams with MFA enabled

## Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-tms.ps1

## Prerequisites = 1
## 1. Ensure Micosoft Teams Module is install or updated

## Variables
$systemmessagecolor = "cyan"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started"

## ensure that install-module -name microsoftteams has been run
## ensure that update-module -name microsoftteams has been run to get latest module
## https://www.powershellgallery.com/packages/MicrosoftTeams/
## Current version = 0.9.3, 25 April 2018
import-module MicrosoftTeams
write-host -foregroundcolor $systemmessagecolor "Microsoft Teams module loaded"

## Connect to Microsoft Teams service
## You will be manually prompted to enter credentials and MFA
Connect-MicrosoftTeams
write-host -foregroundcolor $systemmessagecolor "Now connected to Microsoft Teams Service"