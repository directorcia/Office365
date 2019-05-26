<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Save login credentials to a local XML file for later re-use

Source - https://github.com/directorcia/Office365/blob/master/save-cred-file.ps1

Prerequisites = 0

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## Save manually inputed creds to local file
Get-Credential | Export-CliXml  -Path $credpath

write-host -foregroundcolor $systemmessagecolor "Script completed`n"
