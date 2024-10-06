<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Script designed to get Microsoft Teams configuration information for a tenant
Documentation - 
Source - 

Prerequisites = 1
1. Microsoft Graph PowerShell module installed - https://www.powershellgallery.com/packages/Microsoft.Graph/

#>

## Variables
$processmessagecolor = "green"
$appname = "Reset"

clear-host
Write-Host -foregroundcolor $processmessagecolor "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes "Application.ReadWrite.All", "Directory.ReadWrite.All", "Application.Read.All","Directory.Read.All","AppRoleAssignment.ReadWrite.All" -NoWelcome

Write-Host -foregroundcolor $processmessagecolor "Creating a new application..."
$app = New-MgApplication -DisplayName $appname -ErrorAction Stop -Web @{ RedirectUris = @('https://portal.office.com/') }
Write-Host -foregroundcolor $processmessagecolor "Adding an application password..."
$apppwd = Add-MgApplicationPassword -ApplicationId $app.Id -ErrorAction Stop ## Get Azure AD app password

# Step 2: Create a service principal for the application
$servicePrincipal = New-MgServicePrincipal -AppId $app.AppId

# Output the new application's details
$applicationid = $app.AppId
$tenantid = (get-mgcontext).tenantid
$objectid = $app.Id

Write-Host "`nApplication Created:"
Write-Host "              Tenant ID: $($tenantid)"
Write-Host "              Object ID: $($objectid)"
Write-Host "Application (client) ID: $($applicationid)"
Write-Host "             App Secret: $($apppwd.SecretText)"
Write-Host "   Service Principal ID: $($servicePrincipal.Id)`n"

Write-Host -foregroundcolor $processmessagecolor "Setting permissions for the application..."
# Define variables
$resourceAppId = "00000003-0000-0000-c000-000000000000"  # Microsoft Graph App ID
# https://learn.microsoft.com/en-us/graph/permissions-reference

$permissions = @(
    # https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.users/update-mguser?view=graph-powershell-1.0#description
    @{
        id = "3011c876-62b7-4ada-afa2-506cbbecc68c"  # GUID for User.EnableDisableAccount.All 3011c876-62b7-4ada-afa2-506cbbecc68c
        type = "Role"                                # Use "Role" for application permissions
    },
    @{
        id = "19dbc75e-c2e2-444c-a770-ec69d8559fc7"  # GUID for Directory.ReadWrite.All 19dbc75e-c2e2-444c-a770-ec69d8559fc7
        type = "Role"                                # Use "Role" for application permissions
    },
    @{
        id = "741f803b-c850-494e-b5df-cde7c675a1ca"  # GUID for User.ReadWrite.All 741f803b-c850-494e-b5df-cde7c675a1ca
        type = "Role"                                # Use "Role" for application permissions
    }
)

# Prepare requiredResourceAccess property
$requiredResourceAccess = @(
    @{
        resourceAppId = $resourceAppId
        resourceAccess = $permissions
    }
)

Write-Host -foregroundcolor $processmessagecolor "Updating the application with new permissions..."
# Update the application with new permissions
Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $requiredResourceAccess -ErrorAction Stop

Write-Host -foregroundcolor $processmessagecolor "Granting admin consent for all permissions..."
# Grant admin consent for all permissions in browser
$applicationid = $app.AppId
$tenantid = (get-mgcontext).tenantid
$consenturl = "https://login.microsoftonline.com/$tenantid/oauth2/authorize?client_id=$applicationid&response_type=code&redirect_uri=https%3A%2F%2Fportal.office.com%2F&response_mode=query&state=12345&prompt=admin_consent"
set-clipboard -Value $consenturl
write-host "`nConsent address copied to clipboard"
write-host "  - Please open the consent URL in a browser and grant admin consent for the application.`n"
write-host "`nConsent URL: $consenturl`n"
