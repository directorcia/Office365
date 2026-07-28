<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Creates or updates an Entra application for reset-style operations by calling the Microsoft Graph REST API directly.
Documentation - https://github.com/directorcia/Office365/wiki/Entra-ID-application-for-reset-operations
Source - https://github.com/directorcia/Office365/blob/master/eid-resetapp-set.ps1

Prerequisites:
1. Azure CLI installed and signed in with sufficient permissions: https://aka.ms/azure-cli
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$AppName = 'Reset',

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [switch]$SkipConsentUrl
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Helper used for consistent console output during the script run.
function Write-ProgressMessage {
    param([string]$Message)
    Write-Host -ForegroundColor Green $Message
}

# Prompts the user before making tenant-changing changes unless -Force is supplied.
function Confirm-Action {
    param([string]$Description)

    if ($Force) {
        return $true
    }

    Write-Host "`n$Description"
    $response = Read-Host 'Continue? [y/N]'

    return $response -match '^(y|yes)$'
}

# Acquires an access token for Microsoft Graph using the Azure CLI.
function Get-GraphAccessToken {
    if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
        throw 'Azure CLI (az) is required. Install it and sign in with az login.'
    }

    $token = & az account get-access-token --resource-type ms-graph --query accessToken -o tsv 2>$null
    if (-not $token) {
        throw 'Unable to acquire a Microsoft Graph access token. Run az login and ensure your account has the required permissions.'
    }

    return $token
}

# Returns the current Azure tenant ID from the signed-in Azure CLI context.
function Get-TenantId {
    $tenantId = & az account show --query tenantId -o tsv 2>$null
    if (-not $tenantId) {
        throw 'Unable to determine the current tenant from Azure CLI.'
    }

    return $tenantId
}

# Sends a request to the Microsoft Graph REST API using the current access token.
function Invoke-GraphRequest {
    param(
        [Parameter(Mandatory)]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter()]
        [object]$Body
    )

    $headers = @{
        Authorization = "Bearer $script:GraphToken"
        'Content-Type' = 'application/json'
    }

    $params = @{
        Method = $Method
        Uri = "https://graph.microsoft.com/v1.0$Path"
        Headers = $headers
        ErrorAction = 'Stop'
    }

    if ($null -ne $Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
    }

    try {
        return Invoke-RestMethod @params
    }
    catch {
        $responseBody = $_.ErrorDetails.Message
        if (-not $responseBody) {
            $responseBody = $_.Exception.Message
        }
        throw "Graph request failed for $Method $Path: $responseBody"
    }
}

Clear-Host
Write-ProgressMessage 'Connecting to Microsoft Graph...'

# Load the Azure CLI token and tenant context once so every REST call can reuse them.
$script:GraphToken = Get-GraphAccessToken
$tenantId = Get-TenantId

# Look for an existing app with the requested display name before deciding whether to create or update.
$escapedName = $AppName -replace "'", "''"
$appsResponse = Invoke-GraphRequest -Method GET -Path "/applications?`$filter=displayName eq '$escapedName'&`$top=1"
$existingApp = @($appsResponse.value | Where-Object { $_ }) | Select-Object -First 1

if ($existingApp) {
    $actionDescription = "This will update the existing application '$AppName' in tenant $tenantId."
}
else {
    $actionDescription = "This will create a new application '$AppName' in tenant $tenantId."
}

# Confirm the action before any tenant-changing operations begin.
if (-not (Confirm-Action -Description $actionDescription)) {
    Write-Host 'Operation cancelled.'
    return
}

# Create a new app if none exists; otherwise patch the existing one in place.
if ($existingApp) {
    Write-ProgressMessage "Updating existing application '$($existingApp.displayName)'..."
    $null = Invoke-GraphRequest -Method PATCH -Path "/applications/$($existingApp.id)" -Body @{
        displayName = $AppName
        web = @{ redirectUris = @('https://portal.office.com/') }
    }
    $app = Invoke-GraphRequest -Method GET -Path "/applications/$($existingApp.id)"
    $createdNew = $false
}
else {
    Write-ProgressMessage "Creating a new application named '$AppName'..."
    $app = Invoke-GraphRequest -Method POST -Path '/applications' -Body @{
        displayName = $AppName
        web = @{ redirectUris = @('https://portal.office.com/') }
    }
    $createdNew = $true
}

# Add a new client secret so the app can authenticate.
Write-ProgressMessage 'Adding an application password...'
$passwordResponse = Invoke-GraphRequest -Method POST -Path "/applications/$($app.id)/addPassword" -Body @{
    passwordCredential = @{
        displayName = "Created by $AppName script"
    }
}

# Ensure the app has a service principal, which is required for some consent and assignment scenarios.
Write-ProgressMessage 'Ensuring a service principal exists...'
$servicePrincipalsResponse = Invoke-GraphRequest -Method GET -Path "/servicePrincipals?`$filter=appId eq '$($app.appId)'&`$top=1"
$servicePrincipal = @($servicePrincipalsResponse.value | Where-Object { $_ }) | Select-Object -First 1

if (-not $servicePrincipal) {
    $servicePrincipal = Invoke-GraphRequest -Method POST -Path '/servicePrincipals' -Body @{
        appId = $app.appId
    }
}

# Capture the identifiers needed for the final summary and permission update.
$applicationId = $app.appId
$objectId = $app.id

# Define the Microsoft Graph application permissions being granted to the app.
$resourceAppId = '00000003-0000-0000-c000-000000000000'  # Microsoft Graph App ID
$permissions = @(
    @{ id = '3011c876-62b7-4ada-afa2-506cbbecc68c'; type = 'Role' }, # User.EnableDisableAccount.All
    @{ id = '19dbc75e-c2e2-444c-a770-ec69d8559fc7'; type = 'Role' }, # Directory.ReadWrite.All
    @{ id = '741f803b-c850-494e-b5df-cde7c675a1ca'; type = 'Role' }  # User.ReadWrite.All
)

$requiredResourceAccess = @(
    @{
        resourceAppId = $resourceAppId
        resourceAccess = $permissions
    }
)

# Apply the required Graph permissions to the app object.
Write-ProgressMessage 'Updating the application with the required permissions...'
$null = Invoke-GraphRequest -Method PATCH -Path "/applications/$objectId" -Body @{
    requiredResourceAccess = $requiredResourceAccess
}

# Build the admin-consent URL for the app so the tenant admin can approve the permissions.
$consentUrl = "https://login.microsoftonline.com/${tenantId}/oauth2/authorize?client_id=$applicationId&response_type=code&redirect_uri=https%3A%2F%2Fportal.office.com%2F&response_mode=query&state=12345&prompt=admin_consent"
$status = if ($createdNew) { 'Created' } else { 'Updated' }

# Print a concise summary of the results for the operator.
Write-Host "`nApplication summary:"
Write-Host "  Status: $status"
Write-Host "  Tenant ID: $tenantId"
Write-Host "  Object ID: $objectId"
Write-Host "  Application (client) ID: $applicationId"
Write-Host "  App Secret: $($passwordResponse.secretText)"
Write-Host "  Service Principal ID: $($servicePrincipal.id)"
Write-Host "  Consent URL: $consentUrl"

# Copy the consent URL so the admin can open it quickly.
if (-not $SkipConsentUrl) {
    Set-Clipboard -Value $consentUrl
    Write-Host "`nConsent URL copied to the clipboard."
    Write-Host 'Open it in a browser and grant admin consent for the application.'
}
