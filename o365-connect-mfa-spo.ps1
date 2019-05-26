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
$tenant_input = $true                   ## change to false if you don't wish to be prompted for the tenant name
$tenanturl= "https://<tenantname>-admin.sharepoint.com" ## SharePoint Admin URL for tenant

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

If ($tenant_input -eq $true){
    # Prompt user for tenant name
    $tenantname = Read-Host -prompt "Input tenant name (NOT full tenant URL)"
    $tenanturl = "https://"+$tenantname+"-admin.sharepoint.com"
    Write-host -ForegroundColor $processmessagecolor "SharePoint admin URL =",$tenanturl
}

import-module microsoft.online.sharepoint.powershell -disablenamechecking
write-host -foregroundcolor $processmessagecolor "SharePoint Online module loaded"

## Connect to SharePoint Online Service
## You will be manually prompted to login using MFA
connect-sposervice -url $tenanturl
write-host -foregroundcolor $processmessagecolor "Now connected to SharePoint Online services MFA`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"