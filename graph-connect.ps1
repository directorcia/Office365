## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to connect to the Microsoft Graph

## Source - https://github.com/directorcia/Office365/blob/master/graph-connect.ps1

## Prerequisites = 1
## 1. Azure AD app setup per - https://blog.ciaops.com/2019/04/17/using-interactive-powershell-to-access-the-microsoft-graph/

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## ensure that install-module msonline has been run
## ensure that update-module msonline has been run to get latest module

clear-host

## Script from - https://www.lee-ford.co.uk/getting-started-with-microsoft-graph-with-powershell/

# Azure AD OAuth Application Token for Graph API
# Get OAuth token for a AAD Application (returned as $token)

# Application (client) ID, tenant ID and secret
$clientId = "e621f6e4-6367-45d7-8baa-a93abec4554f"
$tenantId = "5243d63d-7632-4d07-a77e-de0fea1b77a4"
$clientSecret = '4*^);_vK](3+.U}});]Ei?^>[a!u=$#.?^M=]))$/?Xd})8///t|'

# Construct URI
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Construct Body
$body = @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# Graph API call in PowerShell using obtained OAuth token (see other gists for more details)

# Specify the URI to call and method
$uri = "https://graph.microsoft.com/v1.0/security/alerts"
$method = "GET"

# Run Graph API query 
$query = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop

$query

write-host -foregroundcolor $systemmessagecolor "Script Completed`n"