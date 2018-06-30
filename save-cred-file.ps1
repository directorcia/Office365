## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Prerequisites = 0

## Variables
$credpath = "c:\downloads\tenant.xml"   ## local file with credentials

## Save manually inputed creds to local file
Get-Credential | Export-CliXml  -Path $credpath
