<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Creates or updates an Entra application for reset-style operations by calling the Microsoft Graph REST API directly.
This version authenticates to Microsoft Graph using a device-code sign-in flow and does not require any Azure CLI commands to run.
Documentation - https://github.com/directorcia/Office365/wiki/Entra-ID-application-provisioning-and-user-re%E2%80%90enablement
Source - https://github.com/directorcia/Office365/blob/master/eid-resetapp-set-direct.ps1

Prerequisites:
1. PowerShell 7 or later
2. A browser available on the machine running the script
3. Sufficient permissions in the target tenant to create or update applications and service principals
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$AppName,

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [switch]$SkipConsentUrl,

    [Parameter()]
    [string]$TenantId = 'organizations',

    [Parameter()]
    [string]$UserPrincipalName,

    [Parameter()]
    [string]$UserObjectId,

    [Parameter()]
    [switch]$ReEnableUser,

    [Parameter()]
    [switch]$AppOnly,

    [Parameter()]
    [string]$ClientId,

    [Parameter()]
    [string]$ClientSecret,

    [Parameter()]
    [string]$CertificateThumbprint,

    [Parameter()]
    [string]$CertificatePath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Parameters:
# - AppName: The display name of the Entra application to create or update.
# - Force: Skips the confirmation prompt for non-interactive or scripted use.
# - SkipConsentUrl: Prevents copying the admin-consent URL to the clipboard.
# - TenantId: The tenant identifier, or 'organizations' to let the user choose their home tenant.

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

# Acquires a Microsoft Graph session by using the Microsoft Graph PowerShell authentication flow.
# This avoids the Azure CLI dependency while still supporting either interactive sign-in or app-only authentication.
function Get-GraphAccessToken {
    param(
        [string]$Tenant,
        [switch]$AppOnlyMode,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$CertificateThumbprint,
        [string]$CertificatePath
    )

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Microsoft.Graph.Authentication is required. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $scopes = @(
        'Application.ReadWrite.All',
        'Directory.ReadWrite.All',
        'Application.Read.All',
        'Directory.Read.All',
        'AppRoleAssignment.ReadWrite.All',
        'offline_access',
        'openid',
        'profile'
    )

    $effectiveTenant = if ([string]::IsNullOrWhiteSpace($Tenant) -or $Tenant -eq 'organizations') { $null } else { $Tenant }

    if ($AppOnlyMode) {
        if (-not $ClientId) {
            throw 'Provide -ClientId when using -AppOnly.'
        }

        if ($CertificateThumbprint -or $CertificatePath) {
            if ($CertificatePath) {
                $certificate = Get-PfxCertificate -FilePath $CertificatePath
            }
            elseif ($CertificateThumbprint) {
                $certificate = Get-ChildItem -Path Cert:\CurrentUser\My\$CertificateThumbprint -ErrorAction Stop
            }

            if (-not $certificate) {
                throw 'A valid certificate could not be loaded for app-only authentication.'
            }

            Write-ProgressMessage 'Connecting to Microsoft Graph using app-only certificate authentication...'
            if ($effectiveTenant) {
                Connect-MgGraph -ClientId $ClientId -Certificate $certificate -TenantId $effectiveTenant -NoWelcome -ErrorAction Stop | Out-Null
            }
            else {
                Connect-MgGraph -ClientId $ClientId -Certificate $certificate -NoWelcome -ErrorAction Stop | Out-Null
            }
            return $null
        }

        if (-not $ClientSecret) {
            throw 'Provide -ClientSecret when using -AppOnly without a certificate.'
        }

        $secureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
        $credential = [pscredential]::new($ClientId, $secureClientSecret)

        Write-ProgressMessage 'Connecting to Microsoft Graph using app-only secret authentication...'
        if ($effectiveTenant) {
            Connect-MgGraph -ClientSecretCredential $credential -TenantId $effectiveTenant -NoWelcome -ErrorAction Stop | Out-Null
        }
        else {
            Connect-MgGraph -ClientSecretCredential $credential -NoWelcome -ErrorAction Stop | Out-Null
        }
        return $null
    }

    Write-ProgressMessage 'Starting interactive Microsoft Graph sign-in...'

    if ($effectiveTenant) {
        try {
            Connect-MgGraph -Scopes $scopes -UseDeviceCode -NoWelcome -TenantId $effectiveTenant -ErrorAction Stop
        }
        catch {
            Write-Warning "Tenant '$effectiveTenant' could not be used for sign-in. Retrying without a specific tenant hint."
            Connect-MgGraph -Scopes $scopes -UseDeviceCode -NoWelcome -ErrorAction Stop
        }
    }
    else {
        Connect-MgGraph -Scopes $scopes -UseDeviceCode -NoWelcome -ErrorAction Stop
    }

    return $null
}

# Returns the tenant ID from the current Graph connection context.
function Get-TenantIdFromContext {
    $context = Get-MgContext -ErrorAction Stop
    if ($context -and $context.TenantId) {
        return $context.TenantId
    }

    throw 'The Microsoft Graph connection did not expose a tenant identifier.'
}

# Sends a request to the Microsoft Graph REST API using the authenticated Graph session.
# This wrapper centralizes error handling so the rest of the script can focus on provisioning.
function Invoke-GraphApiRequest {
    param(
        [Parameter(Mandatory)]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter()]
        [object]$Body
    )

    $uri = "https://graph.microsoft.com/v1.0$Path"
    $bodyJson = $null
    if ($null -ne $Body) {
        $bodyJson = ($Body | ConvertTo-Json -Depth 10 -Compress)
    }

    try {
        if ($null -ne $bodyJson) {
            return Invoke-MgGraphRequest -Method $Method -Uri $uri -Body $bodyJson -ContentType 'application/json' -ErrorAction Stop
        }

        return Invoke-MgGraphRequest -Method $Method -Uri $uri -ErrorAction Stop
    }
    catch {
        $responseBody = $_.ErrorDetails.Message
        if (-not $responseBody) {
            $responseBody = $_.Exception.Message
        }
        throw ("Graph request failed for {0} {1}: {2}" -f $Method, $Path, $responseBody)
    }
}

# Re-enables a specific user by UPN or object ID using the same authenticated Graph session.
function ReEnable-UserByIdentity {
    param(
        [Parameter()]
        [string]$UserPrincipalName,

        [Parameter()]
        [string]$UserObjectId
    )

    if (-not $UserPrincipalName -and -not $UserObjectId) {
        throw 'Provide either -UserPrincipalName or -UserObjectId.'
    }

    $identifier = if ($UserPrincipalName) { $UserPrincipalName } else { $UserObjectId }
    $encodedIdentifier = [uri]::EscapeDataString($identifier)

    Write-ProgressMessage "Looking up user '$identifier'..."
    $user = Invoke-GraphApiRequest -Method GET -Path "/users/${encodedIdentifier}?`$select=id,userPrincipalName,accountEnabled"

    $accountEnabledValue = $null
    if ($user -is [hashtable]) {
        if ($user.ContainsKey('accountEnabled')) {
            $accountEnabledValue = [bool]$user['accountEnabled']
        }
    }
    else {
        $accountEnabledProperty = $user.PSObject.Properties['accountEnabled']
        if ($null -ne $accountEnabledProperty) {
            $accountEnabledValue = [bool]$accountEnabledProperty.Value
        }
    }

    if ($null -ne $accountEnabledValue -and $accountEnabledValue) {
        Write-Host "User '$identifier' is already enabled."
        return $user
    }

    if ($null -eq $accountEnabledValue) {
        Write-Host "User '$identifier' did not expose an accountEnabled value in the lookup response; proceeding with the re-enable operation."
    }

    $actionDescription = "This will re-enable the user '$identifier'."
    if (-not (Confirm-Action -Description $actionDescription)) {
        Write-Host 'User re-enable cancelled.'
        return $null
    }

    Write-ProgressMessage "Re-enabling user '$identifier'..."
    $updatedUser = Invoke-GraphApiRequest -Method PATCH -Path "/users/${encodedIdentifier}" -Body @{ accountEnabled = $true }
    Write-Host "User '$identifier' was re-enabled successfully."
    return $updatedUser
}

if (-not $AppName -or [string]::IsNullOrWhiteSpace($AppName)) {
    $AppName = 'Reset-Operations-App'
    Write-Host "No app name supplied. Using default name: $AppName"
}

Clear-Host
Write-ProgressMessage 'Connecting to Microsoft Graph...'

# Connect to Microsoft Graph before making any Graph calls, then resolve the tenant from the active context.
$null = Get-GraphAccessToken -Tenant $TenantId -AppOnlyMode:$AppOnly -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -CertificatePath $CertificatePath
$tenantId = Get-TenantIdFromContext

if ($AppOnly -and $ReEnableUser) {
    Write-ProgressMessage 'App-only recovery mode detected. Skipping provisioning and using the configured app identity to re-enable the target user.'
    $status = 'RecoveryOnly'
    $applicationId = $ClientId
    $objectId = $null
    $consentUrl = $null
    $passwordResponse = $null
    $servicePrincipal = $null
}
else {
    # Look for an existing application using the requested display name.
    # This makes the script idempotent and avoids creating duplicate apps when it is run again.
    $escapedName = $AppName -replace "'", "''"
    $appsResponse = Invoke-GraphApiRequest -Method GET -Path "/applications?`$filter=displayName eq '$escapedName'&`$top=1"
    $existingApp = @($appsResponse.value | Where-Object { $_ }) | Select-Object -First 1

    if ($existingApp) {
        $actionDescription = "This will update the existing application '$AppName' in tenant $tenantId."
    }
    else {
        $actionDescription = "This will create a new application '$AppName' in tenant $tenantId."
    }

    # Confirm the action before any tenant-changing operations begin.
    # This protects administrators from accidentally creating or modifying an application in the wrong tenant.
    if (-not (Confirm-Action -Description $actionDescription)) {
        Write-Host 'Operation cancelled.'
        return
    }

    # Create a new application when none exists; otherwise patch the existing object in place.
    # The script preserves the same application name and updates the redirect URI to a common Office portal value.
    if ($existingApp) {
        Write-ProgressMessage "Updating existing application '$($existingApp.displayName)'..."
        $null = Invoke-GraphApiRequest -Method PATCH -Path "/applications/$($existingApp.id)" -Body @{
            displayName = $AppName
            web = @{ redirectUris = @('https://portal.office.com/') }
        }
        $app = Invoke-GraphApiRequest -Method GET -Path "/applications/$($existingApp.id)"
        $createdNew = $false
    }
    else {
        Write-ProgressMessage "Creating a new application named '$AppName'..."
        $app = Invoke-GraphApiRequest -Method POST -Path '/applications' -Body @{
            displayName = $AppName
            web = @{ redirectUris = @('https://portal.office.com/') }
        }
        $createdNew = $true
    }

    # Add a new client secret so the app can authenticate as a confidential client later on.
    Write-ProgressMessage 'Adding an application password...'
    $passwordResponse = Invoke-GraphApiRequest -Method POST -Path "/applications/$($app.id)/addPassword" -Body @{
        passwordCredential = @{
            displayName = "Created by $AppName script"
        }
    }

    # Ensure the application also has a service principal object.
    # A service principal is required for consent flows and for some role or permission assignment scenarios.
    Write-ProgressMessage 'Ensuring a service principal exists...'
    $servicePrincipalsResponse = Invoke-GraphApiRequest -Method GET -Path "/servicePrincipals?`$filter=appId eq '$($app.appId)'&`$top=1"
    $servicePrincipal = @($servicePrincipalsResponse.value | Where-Object { $_ }) | Select-Object -First 1

    if (-not $servicePrincipal) {
        $servicePrincipal = Invoke-GraphApiRequest -Method POST -Path '/servicePrincipals' -Body @{
            appId = $app.appId
        }
    }

    # Capture the identifiers needed for the final summary and permission update.
    $applicationId = $app.appId
    $objectId = $app.id

    # Define the Microsoft Graph application permissions that will be assigned to the app.
    # The chosen permissions are broad administrative permissions and should be reviewed carefully before use in production.
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

    # Apply the required Graph permissions to the application object so it can perform the intended operations.
    Write-ProgressMessage 'Updating the application with the required permissions...'
    $null = Invoke-GraphApiRequest -Method PATCH -Path "/applications/$objectId" -Body @{
        requiredResourceAccess = $requiredResourceAccess
    }

    # Build the admin-consent URL so a tenant administrator can approve the requested permissions.
    $consentUrl = "https://login.microsoftonline.com/${tenantId}/oauth2/authorize?client_id=${applicationId}&response_type=code&redirect_uri=https%3A%2F%2Fportal.office.com%2F&response_mode=query&state=12345&prompt=admin_consent"
    $status = if ($createdNew) { 'Created' } else { 'Updated' }
}

if ($AppOnly -and $ReEnableUser) {
    Write-ProgressMessage 'Running in app-only recovery mode.'
    Write-Host 'This mode uses the application identity directly and does not require a delegated user login.'
}
# Print a concise summary of the results for the operator.
Write-Host "`nApplication summary:"
Write-Host "  Status: $status"
Write-Host "  Tenant ID: $tenantId"
Write-Host "  Object ID: $objectId"
Write-Host "  Application (client) ID: $applicationId"
if ($passwordResponse) {
    Write-Host "  App Secret: $($passwordResponse.secretText)"
}
if ($servicePrincipal) {
    Write-Host "  Service Principal ID: $($servicePrincipal.id)"
}
if ($consentUrl) {
    Write-Host "  Consent URL: $consentUrl"
}

# Optionally re-enable a user after the app provisioning work has completed.
if ($ReEnableUser) {
    if (-not $UserPrincipalName -and -not $UserObjectId) {
        throw 'Specify either -UserPrincipalName or -UserObjectId when using -ReEnableUser.'
    }

    $reEnabledUser = ReEnable-UserByIdentity -UserPrincipalName $UserPrincipalName -UserObjectId $UserObjectId
    if ($reEnabledUser) {
        $userIdentifier = if ($reEnabledUser.userPrincipalName) {
            $reEnabledUser.userPrincipalName
        }
        else {
            $reEnabledUser.id
        }

        Write-Host "  Re-enabled User: $userIdentifier"
    }
}

# Copy the consent URL so the admin can open it quickly.
if (-not $SkipConsentUrl) {
    Set-Clipboard -Value $consentUrl
    Write-Host "`nConsent URL copied to the clipboard."
    Write-Host 'Open it in a browser and grant admin consent for the application.'
}
