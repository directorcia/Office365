## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Source -

## Description
## Script designed to log into the show the all the SharePoint Online Site COllection Administrators across all site collections

## Prerequisites = 1
## 1. Ensure SharePoint online PowerShell module installed or updated

## Variables

Clear-Host

write-host -foregroundcolor Cyan "Script started"

## ensure that SharePoint Online modeule has been installed and loaded

Write-host -ForegroundColor Cyan "Getting all Sharepoint sites in tenant"
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
write-host -foregroundcolor Cyan "Script complete"
