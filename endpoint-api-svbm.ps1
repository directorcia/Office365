<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Connect to the Defender for Endpoint API and return software vulnerabilities by machine

Source - https://github.com/directorcia/Office365/blob/master/endpoint-api-svbm.ps1
Documentation - https://blog.ciaops.com/2021/06/15/using-the-defender-for-endpoint-api-and-powershell/

Prerequisites = 1
1. Azure AD app setup per - https://blog.ciaops.com/2019/04/17/using-interactive-powershell-to-access-the-microsoft-graph/
2. Enter the Client Id, Tenant Id and Client Secret into the variable lines ebfore running

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

# Application (client) ID, tenant ID and secret
$clientid = "<Update>"
$tenantid = "<Update>"
$clientsecret = "<Update>"

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

# Construct URI
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Construct Body
$body = @{
    client_id     = $clientId
    scope         = "https://api.securitycenter.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

write-host -foregroundcolor $processmessagecolor "Get OAuth 2.0 Token"
# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# Graph API call in PowerShell using obtained OAuth token (see other gists for more details)

# Specify the URI to call and method
$uri = "https://api.securitycenter.microsoft.com/api/machines/SoftwareVulnerabilitiesByMachine"
$method = "GET"

write-host -foregroundcolor $processmessagecolor "Run Graph API Query"
# Run Graph API query 
$query = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token" } -ErrorAction Stop -UseBasicParsing

## Screen output

write-host -foregroundcolor $processmessagecolor "Parse results"
$ConvertedOutput = $query.content | ConvertFrom-Json

$ConvertedOutput.value | select-object Devicename, cveid,lastseentimestamp,softwarename,softwarevendor, softwareversion,vulnerabilityseveritylevel | Sort-Object devicename,lastseentimestamp | format-table

write-host -foregroundcolor $systemmessagecolor "`nScript Completed`n"
