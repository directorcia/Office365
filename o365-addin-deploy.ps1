<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Deploy popular Outlook addins centrally

Source - https://github.com/directorcia/Office365/blob/master/o365-addin-deploy.ps1
Reference - https://learn.microsoft.com/en-us/microsoft-365/enterprise/use-the-centralized-deployment-powershell-cmdlets-to-manage-add-ins?view=o365-worldwide

Prerequisites = 1
1. Ensure connected to the Office 365 Central Deployment Service

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$errormessagecolor = "red"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"
write-host -foregroundcolor $processmessagecolor "Connect to Central Add In Service"
try {
    Connect-OrganizationAddInService -ErrorAction Stop
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n[001] - Unable to connect to Central Admin service"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message,"`n"
    exit 1
}

## Deploy addins from Office store
## You will receive an error if the addin is already installed in tenant
## Change the locale to suit your region
write-host -foregroundcolor $processmessagecolor "Deploy Report Message"
try {
    New-OrganizationAddIn -AssetId 'WA104381180' -Locale 'en-AU' -ContentMarket 'en-AU' -ErrorAction Stop | Out-Null ## Report Message
}
catch
{
    Write-Host -ForegroundColor $errormessagecolor "`n[002] - Failed to add asset = WA104381180 (Typically it is already installed)"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message,"`n"
}
write-host -foregroundcolor $processmessagecolor "Deploy Message Header Analyzer"
try {
    New-OrganizationAddIn -AssetId 'WA104005406' -Locale 'en-AU' -ContentMarket 'en-AU' -ErrorAction Stop | Out-Null ## Message Header Analyzer
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n[003] - Failed to add asset = WA104005406 (Typically it is already installed)"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message,"`n"
}

## Enable in tenant
write-host -foregroundcolor $processmessagecolor "Enable Report Message in tenant"
try {
    Set-OrganizationAddIn -ProductId 6046742c-3aee-485e-a4ac-92ab7199db2e -Enabled $true | Out-Null ## Report Message
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n[005] - Failed to enable asset (Typically it is already installed)"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message,"`n"
}
write-host -foregroundcolor $processmessagecolor "Enable Message Header Analyser in tenant"
try {
    Set-OrganizationAddIn -ProductId 62916641-fc48-44ae-a2a3-163811f1c945 -Enabled $true | Out-Null ## Message Header Analyser
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n[006] - Failed to enable asset (Typically it is already installed)"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message,"`n"
}

## Assign addins to all users
write-host -foregroundcolor $processmessagecolor "Assign Report Message to all users"
try {
    Set-OrganizationAddInAssignments -ProductId 6046742c-3aee-485e-a4ac-92ab7199db2e -AssignToEveryone $true  | Out-Null ## Report Message
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n[008] - Failed to assign asset (Typically it is already installed)"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message,"`n"
}
write-host -foregroundcolor $processmessagecolor "Assign Message Header Analyser to all users"
try {
    Set-OrganizationAddInAssignments -ProductId 62916641-fc48-44ae-a2a3-163811f1c945 -AssignToEveryone $true  | Out-Null ## Message Header Analyzer
}
Catch {
    Write-Host -ForegroundColor $errormessagecolor "`n[009] - Failed to assign asset (Typically it is already installed)"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message,"`n"
}

write-host -foregroundcolor $systemmessagecolor "Script Completed"