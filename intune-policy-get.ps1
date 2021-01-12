<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Return the name of all policies configured in EndPoint Manager (Intune and Endpoint)
Source - https://github.com/directorcia/office365/blob/master/intune-policy-get.ps1

Prerequisites = 1
1. Ensure connected to Intune - Use https://github.com/directorcia/Office365/blob/master/Intune-connect.ps1

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

Clear-Host
write-host -foregroundcolor $systemmessagecolor "Script started"
Try {
    Import-Module Microsoft.Graph.Intune | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n[001] - Failed to import Intune module - ", $_.Exception.Message
    exit 1
}
try {
    Connect-MSGraph | Out-Null
}
catch {
       Write-Host -ForegroundColor $errormessagecolor "`n[002] - Failed to connect to Intune - ", $_.Exception.Message
       exit 2 
}

<#          Intune policies             #>
write-host -foregroundcolor $processmessagecolor "`nIntune Compliance policies"
$pols = Get-IntuneDeviceCompliancePolicy
Foreach ($pol in $pols){
    write-host "  - "$pol.displayname
}

write-host -foregroundcolor $processmessagecolor "`nIntune Configuration policies"
$pols = Get-IntuneDeviceConfigurationPolicy
Foreach ($pol in $pols){
    write-host "  - "$pol.displayname
}

write-host -foregroundcolor $processmessagecolor "`nIntune App protection policies"
$pols = Get-IntuneappprotectionPolicy
Foreach ($pol in $pols){
    write-host "  - "$pol.displayname
}

write-host -foregroundcolor $processmessagecolor "`nIntune App configuration policies (targeted)"
$pols = Get-IntuneappconfigurationPolicytargeted
Foreach ($pol in $pols){
    write-host "  - "$pol.displayname
}

<#      EndPoint Policies       #>
$uri = "https://graph.microsoft.com/beta/deviceManagement/intents"
$Configs = (Invoke-MSGraphRequest -Url $uri -HttpMethod GET).Value 

write-host -foregroundcolor $processmessagecolor "`nEndPoint policies"
foreach ($config in $configs) {
    write-host "  - "$config.displayname
}

Write-Host -ForegroundColor $systemmessagecolor "`nScript Finished"