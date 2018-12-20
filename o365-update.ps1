## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to update all the relevant Microsoft Online PowerShell modules

## Source - https://github.com/directorcia/Office365/blob/master/o365-update.ps1

## Prerequisites = 1
## 1. Run PowerShell environment as an administrator

## Variables
$systemmessagecolor = "cyan"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Start Script"

write-host -foregroundcolor $systemmessagecolor "Update Azure AD module"
Update-Module -Name AzureAD -force
write-host -foregroundcolor $systemmessagecolor "Update Azure AD Right Management module"
Update-Module -Name AADRM -force
write-host -foregroundcolor $systemmessagecolor "Update Teams Module"
Update-Module -Name MicrosoftTeams -Force
write-host -foregroundcolor $systemmessagecolor "Update SharePoint Online module"
Update-Module -Name Microsoft.Online.SharePoint.PowerShell -force
write-host -foregroundcolor $systemmessagecolor "Update Microsoft Online module"
Update-Module -Name MSOnline -force
write-host -foregroundcolor $systemmessagecolor "Update Azure module"
Update-Module -name AzureRM -Force
## New Az module
## Update-Module -name Az -force

write-host -foregroundcolor $systemmessagecolor "Finish Script"