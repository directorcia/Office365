param(                           ## if no parameter used then login without MFA and use interactive mode
    [switch]$debug = $false, ## if -debug parameter capture log information
    [switch]$prompt = $false     ## if -prompt parameter used prompt user for input
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Connect to the Microsoft Graph and get vulnerabilities from Microsoft Defender for Endpoint
Documentation - 
Source - https://github.com/directorcia/Office365/blob/master/mde-vul-get.ps1

Prerequisites = 1
1. Azure AD app setup per - https://blog.ciaops.com/2019/04/17/using-interactive-powershell-to-access-the-microsoft-graph/
2. Azure application permissions for WindowsDefenderATP API set:

api/vulnerabilities/machinesVulnerabilities - https://learn.microsoft.com/en-us/microsoft-365/security/defender-endpoint/get-all-vulnerabilities-by-machines?view=o365-worldwide
Application permissions = Vulnerability.Read.All

/api/machines - https://learn.microsoft.com/en-us/microsoft-365/security/defender-endpoint/get-machines?view=o365-worldwide
Application permissions = Machine.Read.All, Machine.ReadWrite.All

Reference - https://blog.ciaops.com/2021/06/15/using-the-defender-for-endpoint-api-and-powershell/

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"
$filename = "..\vulnerabilites.csv"

# Azure AD Application for tenant. Update to suit your environment
$applicationid = "<Update with your details>"
$objectid = "<Update with your details>"
$appsecret = "<Update with your details>"
$tenantid = "<Update with your details>"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host
if ($debug) {
    write-host -foregroundcolor $processmessagecolor "Script activity logged at ..\mde-vul-get.txt"
    start-transcript "..\mde-vul-get.txt"
}

write-host -foregroundcolor $systemmessagecolor "Defender for Endpoint Vulnerabilitities`n"

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

# Specify the URI to call and method to retrieve all machines
$uri = "https://api.securitycenter.microsoft.com/api/machines"
$method = "GET"
$body = @{}

if ($prompt) { pause }

# Run Graph API query
try { 
    $query = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token" } -body $body -ErrorAction Stop -UseBasicParsing
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[002]", $_.Exception.Message
}  

$ConvertedOutput = $query.content | ConvertFrom-Json
$Machinestraw = $ConvertedOutput.value

$MachineSummary = @()                           ## Initialise results array
foreach ($machine in $Machinestraw) {
    $machineSummary += [pscustomobject]@{       
        ID                     = $machine.id
        mergedIntoMachineId    = $machine.mergedIntoMachineId 
        isPotentialDuplication = $machine.isPotentialDuplication
        isExcluded             = $machine.isExcluded
        exclusionReason        = $machine.exclusionReason
        computerDnsName        = $machine.computerDnsName
        firstSeen              = $machine.firstSeen
        firstseenlocal         = (get-date $machine.firstSeen)      # Convert UTC result to local time
        lastSeen               = $machine.lastSeen
        lastseenlocal          = (get-date $machine.lastSeen)       # Convert UTC result to local time
        osPlatform             = $machine.osPlatform
        osVersion              = $machine.osVersion
        osProcessor            = $machine.osProcessor
        version                = $machine.version
        lastIpAddress          = $machine.lastIpAddress
        lastExternalIpAddress  = $machine.lastExternalIpAddress
        agentVersion           = $machine.agentVersion
        osBuild                = $machine.osBuild
        healthStatus           = $machine.healthStatus
        deviceValue            = $machine.deviceValue
        rbacGroupId            = $machine.rbacGroupId
        rbacGroupName          = $machine.rbacGroupName
        riskScore              = $machine.riskScore
        exposureLevel          = $machine.exposureLevel
        isAadJoined            = $machine.isAadJoined
        aadDeviceId            = $machine.aadDeviceId
        machineTags            = $machine.machineTags
        defenderAvStatus       = $machine.defenderAvStatus
        onboardingStatus       = $machine.onboardingStatus
        osArchitecture         = $machine.osArchitecture
        managedBy              = $machine.managedBy
        managedByStatus        = $machine.managedByStatus
        ipAddresses            = $machine.ipAddresses
        vmMetadata             = $machine.vmMetadata      
    }
}
write-host -foregroundcolor $processmessagecolor "`nList devices"
$MachineSummary | Select-Object computerDnsName, OSPlatform, lastseenlocal, lastIpAddress, onboardingStatus | sort-object computerDnsName | Format-Table

if ($prompt) { pause }

# Specify the URI to call and method
$uri = "https://api.securitycenter.microsoft.com/api/vulnerabilities/machinesVulnerabilities"
$method = "GET"
$body = @{}
try { 
    $query = Invoke-WebRequest -Method $method -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token" } -body $body -ErrorAction Stop -UseBasicParsing
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[002]", $_.Exception.Message
}  
$ConvertedOutput = $query.content | ConvertFrom-Json
$vulnerraw = $ConvertedOutput.value

$vulnerSummary = @()                           ## Initialise results array
foreach ($vulner in $vulnerraw) {
    $VulnerSummary += [pscustomobject]@{        ## Build SKU summary array item
        ID             = $vulner.id
        Machine        = ($MachineSummary -match $vulner.machineid).computerdnsname
        cveID          = $vulner.cveId 
        machineId      = $vulner.machineId
        fixingKbId     = $vulner.fixingKbId
        productName    = $vulner.productName
        productVendor  = $vulner.productVendor
        productVersion = $vulner.productVersion
        severity       = $vulner.severity
    }
}
write-host -foregroundcolor $processmessagecolor "List vulnerabilities"
$VulnerSummary | Select-Object machine, cveId, productName | sort-object machine | Format-Table
write-host -foregroundcolor $processmessagecolor "List vulnerabilities to CSV file", $filename
$vulnerSummary | export-csv $filename -NoTypeInformation      ## Export to text file in parent directory

write-host -foregroundcolor $systemmessagecolor "`nDefender for Endpoint Threat Vulnerabilities"
if ($debug) {
    stop-transcript | Out-Null
}