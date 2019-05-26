<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the Office 365 Security and Compliance Center when multi factor security is enabled

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-mfa-sac.ps1

Prerequisites = 1
1. Ensure msonline MFA module installed or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

$CompConnectionUri = "https://ps.compliance.protection.outlook.com/powershell-liveid/"
Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName | ?{ $_ -notmatch "_none_" } | select -First 1)
write-host -foregroundcolor $systemmessagecolor "Exchange Online MFA module loaded"

$Comp = New-EXOPSSession -ConnectionUri $CompConnectionUri -Credential $EXOcreds
$CompImportresults = Import-PSSession $Comp -AllowClobber
write-host -foregroundcolor $processmessagecolor "Connected to Secruity and Compliance Center MFA`n"
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"