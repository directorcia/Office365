<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Connect to the Mirosoft Graph
Source - https://github.com/directorcia/office365/blob/master/msgraph-policy-get.ps1

Prerequisites = 1
1. Ensure Microsoft.Graph module is loaded

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

Clear-Host
write-host -foregroundcolor $systemmessagecolor "Script started"
Try {
    Import-Module Microsoft.Graph | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n[001] - Failed to import Graph module - ", $_.Exception.Message
    exit 1
}
write-host -foregroundcolor $processmessagecolor "Microsoft Graph module loaded"
try {
    Connect-MSGraph | Out-Null
}
catch {
       Write-Host -ForegroundColor $errormessagecolor "`n[002] - Failed to connect to Graph - ", $_.Exception.Message
       exit 2 
}
write-host -foregroundcolor $processmessagecolor "Connected to Microsoft Graph`n"
Write-Host -ForegroundColor $systemmessagecolor "`nScript Finished"