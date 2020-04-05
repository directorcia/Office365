<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the SharePoint Online portal when multi factor security is enabled

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-spo.ps1

Prerequisites = 2
1. Ensure msonline MFA module installed or updated
2. Ensure SharePoint online PowerShell module installed or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
## Change <tenantname> to be your own tenant
$tenant_input = $true                                       ## change to false if you don't wish to have admin domain auto detected
$tenanturl = "https://<tenantname>-admin.sharepoint.com"    ## SharePoint Admin URL for tenant if auto detect not enabled

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## Connect to Office 365 admin service
write-host -foregroundcolor $processmessagecolor "Connecting to Office 365 Admin service"
connect-msolservice
write-host -foregroundcolor $processmessagecolor "Now connected to Office 365 Admin service"

## Auto detect SharePoint Online admin domain
If ($tenant_input -eq $true) {                          ## if the auto detect option is set
    write-host -foregroundcolor $processmessagecolor "Determining SharePoint Administration URL"
    $domains = get-msoldomain                           ## get a list of all domains in tenant
    foreach ($domain in $domains) {                     ## loop through all these domains
        if ($domain.name.contains('onmicrosoft')) {     ## find the onmicrosoft.com domain
            $onname = $domain.name.split(".")           ## split the onmicrosoft.com domain when found at the period. Will produce an array that contains each string as an element
            $tenantname = $onname[0]                    ## the first string in this array is the name of the tenant
        }                                               ## end of find the on.microsoft.com domain
    }                                                   ## end of the domain checking look
$tenanturl = "https://" + $tenantname + "-admin.sharepoint.com"
}                                                       ## end of the auto check option

Write-host -ForegroundColor $processmessagecolor "SharePoint admin URL =", $tenanturl

import-module microsoft.online.sharepoint.powershell -disablenamechecking
write-host -foregroundcolor $processmessagecolor "SharePoint Online module loaded"

## Connect to SharePoint Online Service
## You will be manually prompted to login using MFA
connect-sposervice -url $tenanturl
write-host -foregroundcolor $processmessagecolor "Now connected to SharePoint Online services MFA`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"