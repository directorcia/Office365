## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into the Office 365 Security and Compliance Center when multi factor security is enabled

## Prerequisites = 1
## 1. Ensure msonline MFA module installed or updated

Clear-Host

write-host -foregroundcolor green "Script started"

## set-executionpolicy remotesigned
## May be required once to allow ability to runs scripts in PowerShell

## ensure that Exchange Online MFA modules has been run
## Download and install MFA cmdlets from - https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps

$CompConnectionUri = "https://ps.compliance.protection.outlook.com/powershell-liveid/"
Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName | ?{ $_ -notmatch "_none_" } | select -First 1)
write-host -foregroundcolor green "Exchange Online MFA module loaded"

$Comp = New-EXOPSSession -ConnectionUri $CompConnectionUri -Credential $EXOcreds
$CompImportresults = Import-PSSession $Comp -AllowClobber
write-host -foregroundcolor green "Connected to Secruity and Compliance Center MFA"