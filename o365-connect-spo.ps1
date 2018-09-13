## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into the Office 365 admin portal and the SharePoint Online portal

## Source - 

## Prerequisites = 2
## 1. Ensure msonline module installed or updated
## 2. Ensure SharePoint online PowerShell module installed or updated

## Variables
$systemmessagecolor = "cyan"
$savedcreds=$false                      ## false = manually enter creds, True = from file
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials if required
## Change <tenantname> to be your own tenant
$tenanturl= "https://<tenantname>-admin.sharepoint.com" ## SharePoint Admin URL for tenant

Clear-Host

write-host -foregroundcolor green "Script started"

## set-executionpolicy remotesigned
## May be required once to allow ability to runs scripts in PowerShell

## ensure that install-module msonline has been run
## ensure that update-module msonline has been run to get latest module

## https://www.powershellgallery.com/packages/MSOnline/
## Current version = 1.1.183.8, 18 May 2018

import-module msonline
write-host -foregroundcolor green "MSOnline module loaded"

## Download and install https://www.microsoft.com/en-au/download/details.aspx?id=35588 (SharePoint Online Module)
## Current version = 16.0.7813.1200, 27 June 2018
import-module microsoft.online.sharepoint.powershell -disablenamechecking
write-host -foregroundcolor green "SharePoint Online module loaded"

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

#Connect to SharePoint Online Service
connect-sposervice -url $tenanturl -credential $cred
write-host -foregroundcolor green "Now connected to SharePoint Online services"