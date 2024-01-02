param(                        
    [switch]$debug = $false ## if -debug parameter don't prompt for input
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Disconnect from the Mirosoft Graph
Source - https://github.com/directorcia/Office365/blob/master/mggraph-disconnect.ps1

Prerequisites = 1
1. Ensure Microsoft.Graph module is loaded

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host
if ($debug) {
    write-host "Script activity logged at ..\mggraph-disconnect.txt"
    start-transcript "..\mggraph-disconnect.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

write-host -foregroundcolor $systemmessagecolor "Microsoft Graph disconnect script started"

write-host -foregroundcolor $processmessagecolor "`nDisconnect any existing Graph sessions"
try {
    Disconnect-MgGraph -erroraction stop | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n", $_.Exception.Message,"`n"
    if ($debug) {
        Stop-Transcript | Out-Null
    }
    exit 1
}

Write-Host -ForegroundColor $systemmessagecolor "`nMicrosoft Graph disconnect script finished`n"
if ($debug) {
    Stop-Transcript | Out-Null
}