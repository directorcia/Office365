<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into Microsoft Intune

Source - https://github.com/directorcia/Office365/blob/master/Intune-connect.ps1

Prerequisites = 1
1. Ensure Intune Graph module (Microsoft.Graph.Intune) is installed and updated 

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

if (Get-Module -ListAvailable -Name windowsautopilotintune) {
    Write-Host -foregroundcolor $processmessagecolor "Window Autopilot module found"
} 
else {
    Write-Host -foregroundcolor $warningmessagecolor "[Warning] - Windows Autopilot module not found."
    write-host -foregroundcolor $warningmessagecolor "Install this module if you want to use Windows Autopilot commands"
}

Try {
    Import-Module Microsoft.Graph.Intune | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[001] - Failed to import Intune module - ", $_.Exception.Message
    exit 1
}
write-host -foregroundcolor $processmessagecolor "Intune module loaded"
try {
    Connect-MgGraph | Out-Null
}
catch {
       Write-Host -ForegroundColor $errormessagecolor "[002] - Failed to connect to Intune - ", $_.Exception.Message
       exit 2 
}

write-host -foregroundcolor $processmessagecolor "Now connected to Intune service`n"
write-host -foregroundcolor $systemmessagecolor "Script Finished`n"