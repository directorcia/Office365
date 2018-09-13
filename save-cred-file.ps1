## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed save login credentials to a local XML file for later re-use

## Source - 

## Prerequisites = 0

## Variables
$systemmessagecolor = "cyan"

Clear-Host

write-host -foregroundcolor green "Script started"

## Variables
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials

## Save manually inputed creds to local file
Get-Credential | Export-CliXml  -Path $credpath
