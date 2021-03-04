<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the Office 365 Security and Compliance Center when multi factor security is enabled
Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-sac.ps1
Reference - https://docs.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps

Prerequisites = 1
1. Ensure that Exchange Online PowerShell module V2 installed or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

if (get-module -listavailable -name ExchangeOnlineManagement) {
    ## Has the Exchange Online PowerShell V2 module been loaded?
    write-host -ForegroundColor $processmessagecolor "Exchange Online PowerShell V2 found"
}
else {
    write-host -ForegroundColor yellow -backgroundcolor $errormessagecolor "[001] - Exchange Online PowerShell V2 module not installed. Please install and re-run script`n"
    Stop-Transcript                 ## Terminate transcription
    exit 1                          ## Terminate script
}

import-module ExchangeOnlineManagement
write-host -foregroundcolor $systemmessagecolor "Exchange Online V2 module loaded"

Connect-IPPSSession
write-host -foregroundcolor $processmessagecolor "Connected to Secruity and Compliance Center MFA`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"