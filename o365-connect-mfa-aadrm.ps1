## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into the Azure AD Rights Management module with MFA enabled

## Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-aadrm.ps1

## Prerequisites = 0

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## ensure that install-module aadrm has been run
## ensure that update-module aadrm has been run to get latest module
## https://www.powershellgallery.com/packages/AADRM/
## Current version = 2.13.1.0, 3 May 2018

## Connect to Azure AD Rights Management Service
connect-aadrmservice
write-host -foregroundcolor $processmessagecolor "Now connected to the Azure AD Rights Management Service`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"