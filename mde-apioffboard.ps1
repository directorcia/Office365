param(                         ## if no parameter used then login without MFA and use interactive mode
    [switch]$debug = $false    ## if -debug parameter capture log information
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Connect to the Microsoft Graph and offboard selected devices from Microsoft Defender for Endpoint
Source - https://github.com/directorcia/office365/blob/master/mde-apioffboard.ps1

Prerequisites = 1
1. Azure AD app setup per - https://blog.ciaops.com/2019/04/17/using-interactive-powershell-to-access-the-microsoft-graph/
2. Azure application permissions set:

/WindowsDefenderATP/
Application permissions = Machine.Offboard, Machine.Read.All, Machine.ReadWrite.all

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor ="red"
$warningmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host
if ($debug) {
    write-host -foregroundcolor $processmessagecolor "Script activity logged at ..\mde-apioffboard.txt"
    start-transcript "..\mde-apioffboard.txt"
}

write-host -foregroundcolor $systemmessagecolor "Defender for Endpoint Offboarding via API script started`n"
# Application (client) ID, tenant ID and secret
$applicationid = ""     # enter Azure AD app application id value here
$appsecret = ""         # enter Azure AD app application secret here
$tenantid = ""          # enter Azure AD tenant id here

## Script from - https://www.lee-ford.co.uk/getting-started-with-microsoft-graph-with-powershell/

# Azure AD OAuth Application Token for Graph API
# Get OAuth token for a AAD Application (returned as $token)

# Construct URI
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Construct Body
$body = @{
    client_id     = $applicationid
    scope         = "https://api.securitycenter.microsoft.com/.default"
    client_secret = $appSecret
    grant_type    = "client_credentials"
}
write-host -foregroundcolor $processmessagecolor "Get OAuth 2.0 Token"
# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# Graph API call in PowerShell using obtained OAuth token (see other gists for more details)

# Specify the URI to call and method
$uri = "https://api.securitycenter.microsoft.com/api/machines"
$method = "GET"

write-host -foregroundcolor $processmessagecolor "Get Microsoft Defender for Endpoint devices"

# Run Graph API query
try { 
    $query = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop -UseBasicParsing
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[001]", $_.Exception.Message
}

write-host -foregroundcolor $processmessagecolor "Parse results"
$ConvertedOutput = $query.content | ConvertFrom-Json

$results = $convertedoutput.value | select-object computerdnsname,osplatform,onboardingstatus,id,aaddeviceid | sort-object computerdnsname | Out-GridView -PassThru -title "Select devices to offboard (Multiple selections permitted)"           

foreach ($result in $results) {
    write-host -foregroundcolor gray -backgroundcolor blue "`nDevice =",$result.computerdnsname,"[",$result.id,"]"
    # Specify the URI to call and method
    $uri = "https://api.securitycenter.microsoft.com/api/machines/$($result.id)/offboard"
    $method = "POST"
    $body = @{Comment = "Offboard machine by automation"} | ConvertTo-Json       
    write-host -foregroundcolor $processmessagecolor "  POST API request to offboard device"
    # Run Graph API query
    $catchflag = $false
    try { 
        $query = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -body $body -ErrorAction Stop -UseBasicParsing
    }
    catch {
        if ($_ -match "ActiveREquestAlreadyExists") {
            write-host -foregroundcolor $warningmessagecolor "  *** [WARNING] Offboarding for this device is already in progress"
        } else {
            Write-Host -ForegroundColor $errormessagecolor "`n[002]", $_.Exception.Message
        }
        $catchflag = $true
    }
    if (-not $catchflag) {
        write-host -foregroundcolor $processmessagecolor "  Device offboarding process successfully initiated"
    }
}

write-host -foregroundcolor $systemmessagecolor "`nDefender for Endpoint Offboarding via API script completed`n"
if ($debug) {
    stop-transcript | Out-Null
}