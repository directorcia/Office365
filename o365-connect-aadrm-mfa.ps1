## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into the Azure AD Rights Management module with MFA

## Source - 

## Prerequisites = 0

## Variables
$systemmessagecolor = "cyan"

Clear-Host

write-host -foregroundcolor green "Script started"

## set-executionpolicy remotesigned
## May be required once to allow ability to runs scripts in PowerShell

## ensure that install-module aadrm has been run
## ensure that update-module aadrm has been run to get latest module
## https://www.powershellgallery.com/packages/AADRM/
## Current version = 2.13.1.0, 3 May 2018


## Connect to Azure AD Rights Management Service
connect-aadrmservice
write-host -foregroundcolor green "Now connected to the Azure AD Rights Management Service"