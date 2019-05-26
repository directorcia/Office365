<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Deploy popular Outlook addins centrally

Source - https://github.com/directorcia/Office365/blob/master/o365-addin-deploy.ps1

Reference - https://docs.microsoft.com/en-us/office365/enterprise/use-the-centralized-deployment-powershell-cmdlets-to-manage-add-ins

Prerequisites = 1
1. Ensure connected to the Office 365 Central Deployment Service

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## Deploy addins from Office store
## You will receive an error if the addin is already installed in tenant
## Change the locale to suit your region
New-OrganizationAddIn -AssetId 'WA104381180' -Locale 'en-AU' -ContentMarket 'en-AU' ## Report Message
New-OrganizationAddIn -AssetId 'WA104005406' -Locale 'en-AU' -ContentMarket 'en-AU' ## Message Header Analyzer
New-OrganizationAddIn -AssetId 'WA104379803' -Locale 'en-AU' -ContentMarket 'en-AU' ## FindTime

## Enable in tenant
Set-OrganizationAddIn -ProductId 6046742c-3aee-485e-a4ac-92ab7199db2e -Enabled $true ## Report Message
Set-OrganizationAddIn -ProductId 62916641-fc48-44ae-a2a3-163811f1c945 -Enabled $true ## Message Header Analyser
Set-OrganizationAddIn -ProductId 9758a0e2-7861-440f-b467-1823144e5b65 -Enabled $true ## FindTime

## Assign addins to all users
Set-OrganizationAddInAssignments -ProductId 6046742c-3aee-485e-a4ac-92ab7199db2e -AssignToEveryone $true ## Report Message
Set-OrganizationAddInAssignments -ProductId 62916641-fc48-44ae-a2a3-163811f1c945 -AssignToEveryone $true ## Message Header Analyzer
Set-OrganizationAddInAssignments -ProductId 9758a0e2-7861-440f-b467-1823144e5b65 -AssignToEveryone $true ## FindTime

write-host -foregroundcolor $systemmessagecolor "Script Completed`n"