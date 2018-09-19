## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Source - https://github.com/directorcia/Office365/blob/master/o365-spo-admins.ps1

## Description
## Script designed to log into the show the all the SharePoint Online Site COllection Administrators across all site collections

## Prerequisites = 1
## 1. Ensure SharePoint online PowerShell module installed or updated

## Variables
$systemmessagecolor = "cyan"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started"

## Ensure that SharePoint Online modeule has been installed and loaded

Write-host -ForegroundColor $systemmessagecolor "Getting all Sharepoint sites in tenant"
$SiteCollections  = Get-SPOSite -Limit All

foreach ($site in $SiteCollections) ## Loop through all Site Collections in tenant
{
    Write-host -ForegroundColor Green "Checking site:",$site.url

    $siteusers = get-spouser -site $site.Url    ## get all users for that SharePoint site
    foreach ($siteuser in $siteusers){          ## loop through all the users in the site
        If ($siteuser.issiteadmin -eq $true) {  ## if a users is a Site Collection administrator
            Write-host "Site Admin =", $siteuser.displayname,"["$siteuser.loginname"]"
        }
     }
     write-host
}
write-host -foregroundcolor $systemmessagecolor "Script complete"
