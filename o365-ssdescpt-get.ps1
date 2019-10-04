<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Get all items from tenant secure score

Source - https://github.com/directorcia/Office365/blob/master/o365-ssdescpt-get.ps1

Prerequisites = 1
1. Azure AD app setup per - https://blog.ciaops.com/2019/04/17/using-interactive-powershell-to-access-the-microsoft-graph/
2. Change clientid, tenantid and clientsecret variables below

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

# Application (client) ID, tenant ID and secret
$clientId = "<your clientID here>"              ## This MUST be changed before the script will run correctly
$tenantId = "<your tenantID here>"              ## This MUST be changed before the script will run correctly
$clientSecret = '<your client secret here>'     ## This MUST be changed before the script will run correctly

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

start-transcript "..\o365-ssdescpt-get $(get-date -f yyyyMMddHHmmss).txt"       ## write output file to parent directory

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## Script from - https://www.lee-ford.co.uk/getting-started-with-microsoft-graph-with-powershell/

# Azure AD OAuth Application Token for Graph API
# Get OAuth token for a AAD Application (returned as $token)

# Construct URI
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Construct Body
$body = @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
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
$uri = "https://graph.microsoft.com/beta/security/securescores"
$method = "GET"

write-host -foregroundcolor $processmessagecolor "Run Graph API Query"
# Run Graph API query 
$query = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token" } -ErrorAction Stop 

write-host -foregroundcolor $processmessagecolor "Parse results"
$ConvertedOutput = $query | Select-Object -ExpandProperty content | ConvertFrom-Json

write-host -foregroundcolor $processmessagecolor "Display results`n"
foreach ($control in $convertedoutput) {
    $names = $control.value.controlscores.description
    $item = 0
    foreach ($name in $names) {
        $item++
        write-host -foregroundcolor green -BackgroundColor Black "`n*** Item", $item, "***"
        write-host $name        
    }
}

write-host -foregroundcolor $systemmessagecolor "`nScript Completed`n"

stop-transcript