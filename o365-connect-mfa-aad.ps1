## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into Azure AD with MFA enabled

## Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-aad.ps1

## Prerequisites = 2
## 1. Ensure azuread module installed or updated
## 2. Ensure msonline MFA module installed or updated

## Variables
$systemmessagecolor = "cyan"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started"

## ensure that install-module azuread has been run
## ensure that update-module azuread has been run to get latest module

## https://www.powershellgallery.com/packages/AzureAD/
## Current version = 2.0.1.16, 21 June 2018
import-module azuread
write-host -foregroundcolor $systemmessagecolor "AzureAD module loaded"

## ensure that Exchange Online MFA modules has been run
## Download and install MFA cmdlets from - https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps

## Connect to Azure AD service
## You will be manually prompted to enter credentials and MFA
Connect-AzureAD
write-host -foregroundcolor $systemmessagecolor "Now connected to Azure AD Service with MFA"