## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed save login credentials to a local XML file for later re-use

## Source - 

## Prerequisites = 0

## Variables
$systemmessagecolor = "cyan"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor green "Script started"

## Variables
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials

## Save manually inputed creds to local file
Get-Credential | Export-CliXml  -Path $credpath
