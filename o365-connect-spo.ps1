<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the Office 365 admin portal and the SharePoint Online portal

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-spo.ps1

Prerequisites = 2
1. Ensure msonline module installed or updated
2. Ensure SharePoint online PowerShell module installed or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$savedcreds=$false                      ## false = manually enter creds, True = from file
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials if required
## Change <tenantname> to be your own tenant
$tenant_input = $true                   ## change to false if you don't wish to be prompted for the tenant name
$tenanturl= "https://<tenantname>-admin.sharepoint.com" ## SharePoint Admin URL for tenant if prompting not enabled

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

If ($tenant_input -eq $true){
    # Prompt user for tenant name
    $tenantname = Read-Host -prompt "Input tenant name (NOT full tenant URL)"
    $tenanturl = "https://"+$tenantname+"-admin.sharepoint.com"
    Write-host -ForegroundColor $processmessagecolor "SharePoint admin URL =",$tenanturl
}

import-module msonline
write-host -foregroundcolor $processmessagecolor "MSOnline module loaded"

import-module microsoft.online.sharepoint.powershell -disablenamechecking
write-host -foregroundcolor $processmessagecolor "SharePoint Online module loaded"

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
write-host -foregroundcolor $processmessagecolor "Now connected to Office 365 Admin service"

#Connect to SharePoint Online Service
connect-sposervice -url $tenanturl -credential $cred
write-host -foregroundcolor $processmessagecolor "Now connected to SharePoint Online services`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"