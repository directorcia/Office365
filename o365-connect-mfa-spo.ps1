## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into the SharePoint Online portal when multi factor security is enabled

## Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-spo.ps1

## Prerequisites = 2
## 1. Ensure msonline MFA module installed or updated
## 2. Ensure SharePoint online PowerShell module installed or updated

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
## Change <tenantname> to be your own tenant
$tenanturl= "https://<tenantname>-admin.sharepoint.com" ## SharePoint Admin URL for tenant

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## Download and install https://www.microsoft.com/en-au/download/details.aspx?id=35588 (SharePoint Online Module)
## Current version = 16.0.7813.1200, 27 June 2018
import-module microsoft.online.sharepoint.powershell -disablenamechecking
write-host -foregroundcolor $processmessagecolor "SharePoint Online module loaded"

## ensure that Exchange Online MFA modules has been run
## Download and install MFA cmdlets from - https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps

## Connect to SharePoint Online Service
## You will be manually prompted to login using MFA
connect-sposervice -url $tenanturl
write-host -foregroundcolor $processmessagecolor "Now connected to SharePoint Online services MFA`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"