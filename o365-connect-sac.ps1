## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to log into the Office 365 Security and Compliance Center

## Prerequisites = 1
## 1. Ensure msonline module installed or updated

## Variables
$savedcreds=$false                      ## false = manually enter creds, True = from file
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials if required

Clear-Host

write-host -foregroundcolor green "Script started"

## set-executionpolicy remotesigned
## May be required once to allow ability to runs scripts in PowerShell

## ensure that install-module msonline has been run
## ensure that update-module msonline has been run to get latest module
import-module msonline
write-host -foregroundcolor green "MSOnline module loaded"

## Get tenant login credentials
if ($savedcreds) {
    ## Get creds from local file
    $cred =import-clixml -path $credpath
}
else {
    ## Get creds manually
    $cred=get-credential 
}

## Connect to Office 365 admin service
connect-msolservice -credential $cred
write-host -foregroundcolor green "Now connected to Office 365 Admin service"

## Connect to the Office 365 Security and Compliance Center
Write-Output "Getting the Security & Compliance Center cmdlets" 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
    -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ `
    -Credential $cred -Authentication Basic -AllowRedirection
 
Import-PSSession $Session
write-host -foregroundcolor green "Now connected to Office 365 Security and Compliance Center"