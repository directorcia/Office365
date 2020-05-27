<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log to Exchange Online using the V2 modules
Reference - https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/exchange-online-powershell-v2/exchange-online-powershell-v2?view=exchange-ps

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-exov2.ps1

Prerequisites = 1
1. Ensure Exchange Online V2 module is loaded

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

Try {
    Import-Module ExchangeOnlineManagement | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[001] - Failed to import Exchange module - ", $_.Exception.Message
    exit 1
}
write-host -foregroundcolor $processmessagecolor "Exchange Online V2 module loaded"
try {
    Connect-ExchangeOnline -ShowProgress $false | Out-Null
}
catch {
       Write-Host -ForegroundColor $errormessagecolor "[002] - Failed to connect to Exchange Online - ", $_.Exception.Message
       exit 2 
}

write-host -foregroundcolor $processmessagecolor "Connected to Exchange Online`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"