## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Variables
$savedcreds=$false                      ## false = manually enter creds, True = auto
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials

clear-host

## set-executionpolicy remotesigned
## May be required once to allow ability to runs scripts in PowerShell

## ensure that install-module msonline has been run
## ensure that update-module msonline has been run to get latest module
import-module msonline

if ($savedcreds) {
    ## Get creds from local file
    $cred =import-clixml -path $credpath
}
else {
    ## Get creds manually
    $cred=get-credential 
}

connect-msolservice -credential $cred

## Start Exchange Online session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxyMethod=RPS -Credential $Cred -Authentication Basic -AllowRedirection
import-PSSession $Session