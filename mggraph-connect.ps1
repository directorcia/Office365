param(                        
    [switch]$debug = $false, ## if -debug parameter don't prompt for input
    [switch]$noupdate = $false, ## if -noupdate used then module will not be checked for more recent version
    [switch]$noprompt = $false   ## if -noprompt parameter used don't prompt user for input
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Connect to the Mirosoft Graph
Source - 

Prerequisites = 1
1. Ensure Microsoft.Graph module is loaded

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

Clear-Host
if ($debug) {
    write-host "Script activity logged at ..\mggraph-connect.txt"
    start-transcript "..\mggraph-connect.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

write-host -foregroundcolor $systemmessagecolor "Microsoft Graph connect script started"
Try {
    Import-Module Microsoft.Graph | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n[001] - Failed to import Graph module - ", $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null
    }
    exit 1
}
write-host -foregroundcolor $processmessagecolor "Microsoft Graph module loaded"
try {
    Connect-MgGraph | Out-Null
#    Select-MgProfile v1.0                      # Use this to force a version of the Graph version (v1.0 or beta)
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n[002] - Failed to connect to Graph - ", $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null
    }
    exit 2 
}
write-host -foregroundcolor $processmessagecolor "Connected to Microsoft Graph`n"
Write-Host -ForegroundColor $systemmessagecolor "`nMicrosoft Graph connect script finished"
if ($debug) {
    Stop-Transcript | Out-Null
}