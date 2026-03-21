param(
    # Enable transcript logging to a local text file for troubleshooting.
    [switch]$debug = $false, ## if -debug parameter don't prompt for input
    # Export final device results to CSV in addition to console output.
    [switch]$csv = $false, ## export to CSV
    # Ask the operator to confirm the signed-in Graph account before continuing.
    [switch]$prompt = $false, ## if -prompt parameter used user prompted for input
    # Force a fresh interactive sign-in so account picker appears each run.
    [switch]$forceLogin = $true,
    # Optional verified domain to validate tenant context without using tenant ID (example: contoso.com).
    [string]$expectedTenantDomain = "",
    # Optional tenant ID (GUID) to force authentication against a specific tenant.
    [string]$tenantId = ""
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on licenses for tenant
Source - https://github.com/directorcia/Office365/blob/master/graph-devices-get.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Report-tenant-devices

Prerequisites = 1
1. Ensure the MS Graph module is installed

If you find value in this script please support the author of these scripts by:

- https://ko-fi.com/ciaops

or

- becoming a CIAOPS Patron: https://www.ciaops.com/patron

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"
$outputFile = "..\graph-devices.csv"
$scriptStart = Get-Date

# Retrieve all pages for a Graph collection endpoint by following @odata.nextLink.
function Get-GraphCollection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    # Use a generic list to avoid slow PowerShell array expansion in large tenants.
    $items = New-Object System.Collections.Generic.List[object]
    $nextLink = $Uri
    $pageNumber = 0

    while ($null -ne $nextLink) {
        $pageNumber++
        Write-Host -ForegroundColor $processmessagecolor ("Requesting Graph page {0} ..." -f $pageNumber)
        $response = Invoke-MgGraphRequest -Uri $nextLink -Method GET -OutputType PSObject
        $pageItems = 0
        if ($null -ne $response.value) {
            foreach ($entry in $response.value) {
                [void]$items.Add($entry)
                $pageItems++
            }
        }

        Write-Host -ForegroundColor $processmessagecolor ("Page {0} returned {1} device object(s). Running total: {2}" -f $pageNumber, $pageItems, $items.Count)

        # When nextLink is null, paging is complete.
        $nextLink = $response.'@odata.nextLink'
    }

    Write-Host -ForegroundColor $processmessagecolor ("Completed Graph paging after {0} page(s)." -f $pageNumber)
    return $items
}

# Query tenant metadata used for runtime verification and operator visibility.
function Get-TenantVerificationInfo {
    $tenantInfo = [pscustomobject]@{
        OrganizationId = $null
        OrganizationName = $null
        VerifiedDomains = @()
    }

    $orgInfo = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/organization?`$select=id,displayName,verifiedDomains" -Method GET -OutputType PSObject
    if ($orgInfo.value.Count -gt 0) {
        $tenantInfo.OrganizationId = $orgInfo.value[0].id
        $tenantInfo.OrganizationName = $orgInfo.value[0].displayName

        $domains = @()
        foreach ($domain in $orgInfo.value[0].verifiedDomains) {
            if ($domain.isVerified -eq $true -and -not [string]::IsNullOrWhiteSpace($domain.name)) {
                $domains += $domain.name
            }
        }

        $tenantInfo.VerifiedDomains = $domains | Sort-Object -Unique
    }

    return $tenantInfo
}

if ($debug) {
    # Start a transcript so all host output can be reviewed after execution.
    Write-Host "Script activity logged at .\graph-devices-get.txt"
    Start-Transcript ".\graph-devices-get.txt" | Out-Null ## Log file created in parent directory that is overwritten on each run
}

Clear-Host
Write-Host -ForegroundColor $systemmessagecolor "Tenant device report script - Started`n"
Write-Host -ForegroundColor $processmessagecolor "Connect to MS Graph"
$scopes = "Device.Read.All"
Write-Host -ForegroundColor $processmessagecolor "Required scope(s): $scopes"
Write-Host -ForegroundColor $processmessagecolor "Authentication mode: interactive login"
if (-not [string]::IsNullOrWhiteSpace($expectedTenantDomain)) {
    Write-Host -ForegroundColor $processmessagecolor "Expected tenant domain: $expectedTenantDomain"
}
if (-not [string]::IsNullOrWhiteSpace($tenantId)) {
    Write-Host -ForegroundColor $processmessagecolor "Requested tenant id: $tenantId"
}

# Fail early with a clear message if the Graph SDK is not available.
if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
    Write-Host -ForegroundColor $errormessagecolor "Microsoft Graph PowerShell SDK is not installed. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    exit 1
}

$confirmedContext = $false
do {
    try {
        # Clear existing session so login-based auth can switch account/tenant cleanly.
        if ($forceLogin) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }

        # Prompt for interactive auth and request only the scope needed by this report.
        if (-not [string]::IsNullOrWhiteSpace($tenantId)) {
            Connect-MgGraph -Scopes $scopes -TenantId $tenantId -ContextScope Process -NoWelcome | Out-Null
        }
        else {
            Connect-MgGraph -Scopes $scopes -ContextScope Process -NoWelcome | Out-Null
        }
    }
    catch {
        Write-Host -ForegroundColor $errormessagecolor "`nUnable to connect to Microsoft Graph."
        Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
        exit 1
    }

    $graphcontext = Get-MgContext
    Write-Host -ForegroundColor $processmessagecolor "Connected account =", $graphcontext.Account
    Write-Host -ForegroundColor $processmessagecolor "Connected tenant id =", $graphcontext.TenantId
    Write-Host -ForegroundColor $processmessagecolor "Graph profile in use =", $graphcontext.Environment

    if (-not [string]::IsNullOrWhiteSpace($tenantId) -and $graphcontext.TenantId -ne $tenantId) {
        Write-Host -ForegroundColor $errormessagecolor "Tenant mismatch detected. Expected $tenantId but connected to $($graphcontext.TenantId)."
        Disconnect-MgGraph | Out-Null
        exit 1
    }

    $tenantInfo = $null
    try {
        $tenantInfo = Get-TenantVerificationInfo
        if ($null -ne $tenantInfo) {
            Write-Host -ForegroundColor $processmessagecolor "Connected tenant name =", $tenantInfo.OrganizationName
            Write-Host -ForegroundColor $processmessagecolor "Connected tenant id (organization) =", $tenantInfo.OrganizationId
            if ($tenantInfo.VerifiedDomains.Count -gt 0) {
                Write-Host -ForegroundColor $processmessagecolor "Verified tenant domain(s) =", ($tenantInfo.VerifiedDomains -join ", ")
            }
        }
    }
    catch {
        Write-Host -ForegroundColor $warningmessagecolor "Unable to query organization profile for extra tenant verification."
    }

    if (-not [string]::IsNullOrWhiteSpace($expectedTenantDomain) -and $null -ne $tenantInfo) {
        $domainMatched = $false
        foreach ($domain in $tenantInfo.VerifiedDomains) {
            if ($domain.ToLowerInvariant() -eq $expectedTenantDomain.ToLowerInvariant()) {
                $domainMatched = $true
                break
            }
        }

        if (-not $domainMatched) {
            Write-Host -ForegroundColor $errormessagecolor "Tenant domain mismatch. Connected tenant does not contain verified domain '$expectedTenantDomain'."
            if ($prompt) {
                Write-Host -ForegroundColor $warningmessagecolor "Re-authenticating. Select an account in the correct tenant."
                Disconnect-MgGraph | Out-Null
                $confirmedContext = $false
                continue
            }

            Disconnect-MgGraph | Out-Null
            exit 1
        }
    }

    if (-not $prompt) {
        $confirmedContext = $true
    }
    else {
        do {
            $response = Read-Host -Prompt "`nUse this account and tenant? [Y/N]"
        } until (-not [string]::IsNullOrEmpty($response))

        if ($response -eq "Y" -or $response -eq "y") {
            $confirmedContext = $true
        }
        else {
            Write-Host -ForegroundColor $warningmessagecolor "Re-authenticating. Select the correct account/tenant in the sign-in window."
            Disconnect-MgGraph | Out-Null
            $confirmedContext = $false
        }
    }
} until ($confirmedContext)

if ($prompt) { Read-Host -Prompt "`n[PROMPT] -- Press Enter to continue" }

# Get all devices with paging support
# https://learn.microsoft.com/graph/api/device-list
$Url = "https://graph.microsoft.com/v1.0/devices?`$top=999"
Write-Host -ForegroundColor $processmessagecolor "Make Graph request for all devices"
try {
    # Collect every page of results so large tenants are fully represented.
    $results = Get-GraphCollection -Uri $Url
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`nFailed to retrieve devices from Microsoft Graph."
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    Disconnect-MgGraph | Out-Null
    exit 1
}

Write-Host -ForegroundColor $processmessagecolor ("Graph returned {0} raw device record(s)." -f $results.Count)

$devicesummary = New-Object System.Collections.Generic.List[object]
Write-Host -ForegroundColor $processmessagecolor "Transforming device records for report output ..."
if ($results.Count -eq 0) {
    Write-Host -ForegroundColor $warningmessagecolor "No devices were returned for this tenant/context."
}

$processed = 0
foreach ($result in $results) {
    # Keep only key device properties for readable console/CSV reporting.
    $devicesummary.Add([pscustomobject]@{
        Displayname     = $result.displayname
        DeviceID        = $result.DeviceId
        OperatingSystem = $result.OperatingSystem
        Trusttype       = $result.Trusttype
    }) | Out-Null

    $processed++
    if (($processed % 250) -eq 0) {
        Write-Host -ForegroundColor $processmessagecolor ("Processed {0} device record(s) ..." -f $processed)
    }
}

# Output the devices
Write-Host -ForegroundColor $processmessagecolor "Devices returned:", $devicesummary.Count
Write-Host -ForegroundColor $processmessagecolor "Displaying sorted results in console table ..."
# Sort by display name to make both console and CSV output easier to scan.
$devicesummary | Sort-Object DisplayName | Format-Table DisplayName, DeviceId, OperatingSystem, Trusttype

if ($csv) {
    # Optional export path for post-processing in Excel or other tooling.
    Write-Host -ForegroundColor $processmessagecolor "`nOutput to CSV", $outputFile
    $devicesummary | Export-Csv $outputFile -NoTypeInformation
    Write-Host -ForegroundColor $processmessagecolor "CSV export complete."
}

Write-Host -ForegroundColor $processmessagecolor "Disconnecting Graph session ..."
Disconnect-MgGraph | Out-Null

$elapsed = (Get-Date) - $scriptStart
Write-Host -ForegroundColor $processmessagecolor ("Elapsed time: {0:mm\:ss}" -f $elapsed)
Write-Host -ForegroundColor $systemmessagecolor "`nGraph devices script - Finished"
if ($debug) {
    # End transcript only when debug logging was enabled.
    Stop-Transcript | Out-Null
}
