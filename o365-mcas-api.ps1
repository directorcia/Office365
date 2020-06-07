<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Get data from Microsoft Cloud App Security

Source - https://github.com/directorcia/Office365/blob/master/o365-mcas-api.ps1

Prerequisites = 1
1. Create an API Token in your tenant via - https://blog.ciaops.com/2019/10/08/connecting-to-cloud-app-security-api/

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

$mcasuri = "<your MCAS URI here>"                    ## This MUST be changed before the script will run correctly. Note URI format is full web address including https://
## e.g. https://tenantname.us.portal.cloudappsecurity.com
$mcastoken = "<your MCAS Token here>"                  ## This MUST be changed before the script will run correctly
## e.g. HtsYSNUasismYwewgatgsHTMPQFSB7GSw429nsgsiwemw81gfs6nw01n62nsgfvaGHKw7M72hshsbvs6702iJj782902HJsaoanweuiAnonopnma072802nASlolana0nmapJNsn8hajBNAu2kb
$endpoint = "alerts"

$response = Invoke-RestMethod `
    -uri "$mcasuri/api/v1/$endpoint/" `
    -headers @{authorization = "Token $mcastoken" } `
    -method POST `
    -body ($body | convertto-json -depth 2) `
    -verbose

$response.data

write-host -foregroundcolor $systemmessagecolor "`nScript Completed`n"