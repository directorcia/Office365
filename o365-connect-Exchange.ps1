clear-host

## set-executionpolicy remotesigned
## May be required once to allow ability to runs scripts in PowerShell

import-module msonline

## $cred=get-credential
connect-msolservice -credential $cred

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxyMethod=RPS -Credential $Cred -Authentication Basic -AllowRedirection
import-PSSession $Session