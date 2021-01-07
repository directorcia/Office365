<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into Microsoft Teams with MFA enabled

Source - https://github.com/directorcia/Office365/blob/master/win10-connect-autopilot.ps1

Prerequisites = 1
1. Ensure Windows Autopilot Graph module (Microsoft.Graph.Intune) is installed and updated 

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

Try {
    Import-Module Microsoft.Graph.Intune | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[001] - Failed to import Windows Autopilot module - ", $_.Exception.Message
    exit 1
}
write-host -foregroundcolor $processmessagecolor "Windows Autopilot module loaded"
try {
    Connect-MSGraph | Out-Null
}
catch {
       Write-Host -ForegroundColor $errormessagecolor "[002] - Failed to connect to Windows Autopilot - ", $_.Exception.Message
       exit 2 
}

write-host -foregroundcolor $processmessagecolor "Now connected to Windows Intune service`n"
write-host -foregroundcolor $systemmessagecolor "Script Finished`n"