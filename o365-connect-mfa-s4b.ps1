## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into the Skype for Business Online portal with multi factor enabled

## Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-s4b.ps1

## Prerequisites = 1
## 1. Ensure Skype for Business online PowerShell module installed or updated
## 2. Ensure msonline MFA module installed or updated

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## Download and install https://www.microsoft.com/en-au/download/details.aspx?id=39366 (Skype for Business Online Module)
## Current version = 7.0.1994.0, 26 February 2018
import-module skypeonlineconnector
write-host -foregroundcolor $processmessagecolor "Skype for Business module loaded"

## ensure that Exchange Online MFA modules has been run
## Download and install MFA cmdlets from - https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps

## Connect to Skype for Business Online Service
## You will be manually prompted to enter your credentials and MFA
$sfboSession=new-csonlinesession
import-pssession $sfboSession
write-host -foregroundcolor $processmessagecolor "Now connected to Skype for Business Online services`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"