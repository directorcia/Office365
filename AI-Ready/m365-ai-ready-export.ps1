<#
================================================================================
GITHUB COPILOT — IMPLEMENTATION BRIEF (read before completing this file)
================================================================================
GOAL: A READ-ONLY PowerShell script that extracts Microsoft 365 tenant
configuration and posture data to assess "AI / Copilot readiness", and writes
ONE consolidated JSON file for downstream analysis.

HARD RULES:
  - READ-ONLY. Use only Get-* / Connect-* / Disconnect-* cmdlets.
    NEVER emit Set-, New-, Remove-, Update-, Add-, Grant-, or Revoke- cmdlets.
  - RESILIENT. Wrap every collector in try/catch. On failure, record the error
    in $script:Report.Errors and CONTINUE. Never hard-fail the whole run.
  - NO SECRETS. Output config/IDs only — never tokens, passwords, or keys.
  - DETERMINISTIC JSON. Use stable, ordered key names across all tenants.
  - SAMPLING-AWARE. Honour -IncludeSampling / -SampleSize on heavy collectors
    (SharePoint sites, permissions). Note truncation in the output.
  - Prefer Microsoft Graph cmdlets; use ExchangeOnlineManagement, PnP.PowerShell,
    and SharePoint Online Management Shell only where Graph is insufficient.
  - PowerShell 7+. Use [ordered] hashtables for all JSON sections.
================================================================================

Documentation - https://github.com/directorcia/Office365/wiki/Microsoft-365-AI-Readiness-Export-%E2%80%90-Execution-and-Operations-Guide

#>

[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputPath = ".\AIReadiness_$(Get-Date -Format 'yyyyMMdd_HHmmss').json",
    [string]$GraphClientId = '04b07795-8ddb-461a-bbee-02f9e1bf7b46',
    [string]$PnPClientId,
    [string[]]$GraphScopes,
    [switch]$RequestRequiredGraphScopesUpFront,
    [switch]$ForceInteractiveTeamsPnPLogin,
    [switch]$SharePointCollectionChild,
    [string]$SharePointAdminUrl,
    [switch]$IncludeSampling,
    [switch]$IncludeDiagnostics,
    [switch]$IncludeDetailedData,
    [ValidateRange(1, 5000)]
    [int]$SampleSize = 200
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
function Write-Step {
    param(
        [Parameter(Mandatory)][string]$Message,
        [string]$Color = 'Cyan'
    )
    Write-Host $Message -ForegroundColor $Color
}

function Write-Phase {
    param([Parameter(Mandatory)][string]$Title)

    Write-Host ''
    Write-Host "===== $Title =====" -ForegroundColor Magenta
}

function Write-Reassurance {
    Write-Step 'READ-ONLY mode: this script only reads tenant settings and writes one local JSON report.' -Color Green
    Write-Step 'No tenant configuration is created, updated, or removed by this run.' -Color Green
}

function Resolve-OutputFilePath {
    param([Parameter(Mandatory)][string]$Path)

    $currentLocation = (Get-Location).Path
    $resolvedPath = if ([System.IO.Path]::IsPathRooted($Path)) {
        [System.IO.Path]::GetFullPath($Path)
    }
    else {
        [System.IO.Path]::GetFullPath((Join-Path -Path $currentLocation -ChildPath $Path))
    }

    $directory = Split-Path -Path $resolvedPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        $null = New-Item -ItemType Directory -Path $directory -Force
    }

    return $resolvedPath
}

function Convert-ToPowerShellLiteral {
    param([Parameter(Mandatory)][object]$Value)

    if ($null -eq $Value) { return '$null' }

    if ($Value -is [string]) {
        return "'" + ($Value -replace "'", "''") + "'"
    }

    if ($Value -is [bool] -or $Value -is [switch]) {
        return (if ([bool]$Value) { '$true' } else { '$false' })
    }

    return [string]$Value
}

function Get-CurrentScriptArgumentString {
    $parts = [System.Collections.Generic.List[string]]::new()

    foreach ($entry in ($PSBoundParameters.GetEnumerator() | Sort-Object Key)) {
        $key = [string]$entry.Key
        $value = $entry.Value

        if ($value -is [switch] -or $value -is [bool]) {
            if ([bool]$value) {
                $parts.Add("-$key")
            }
            continue
        }

        if ($value -is [System.Array]) {
            foreach ($item in @($value)) {
                $parts.Add("-$key")
                $parts.Add((Convert-ToPowerShellLiteral -Value $item))
            }
            continue
        }

        $parts.Add("-$key")
        $parts.Add((Convert-ToPowerShellLiteral -Value $value))
    }

    return ($parts -join ' ')
}

function Get-CleanRelaunchCommand {
    $scriptPathLiteral = Convert-ToPowerShellLiteral -Value $PSCommandPath
    $argsText = Get-CurrentScriptArgumentString
    $cmd = "pwsh -NoProfile -NoLogo -ExecutionPolicy Bypass -File $scriptPathLiteral"
    if (-not [string]::IsNullOrWhiteSpace($argsText)) {
        $cmd = "$cmd $argsText"
    }
    return $cmd
}

function Get-SharePointAdminUrl {
    $org = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/organization')) | Select-Object -First 1
    $initialDomain = @($org.VerifiedDomains | Where-Object { $_.Name -match '\.onmicrosoft\.com$' } | Select-Object -ExpandProperty Name -First 1)
    $tenantSlug = if ($initialDomain) { $initialDomain -replace '\.onmicrosoft\.com$', '' } else { $null }

    if ([string]::IsNullOrWhiteSpace($tenantSlug)) {
        return $null
    }

    return "https://$($tenantSlug.ToLowerInvariant())-admin.sharepoint.com"
}

function Get-PnPClientIdForConnection {
    if (-not [string]::IsNullOrWhiteSpace($PnPClientId)) {
        return $PnPClientId
    }

    foreach ($envName in @('ENTRAID_APP_ID', 'ENTRAID_CLIENT_ID', 'PNP_CLIENT_ID')) {
        $envValue = [Environment]::GetEnvironmentVariable($envName)
        if (-not [string]::IsNullOrWhiteSpace($envValue)) {
            return $envValue
        }
    }

    return $null
}

function Connect-PnPForSharePointCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AdminUrl
    )

    $effectivePnPClientId = Get-PnPClientIdForConnection
    $effectiveTenant = if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $TenantId } else { $null }
    $attemptErrors = [System.Collections.Generic.List[string]]::new()

    $baseParams = @{
        Url = $AdminUrl
        ErrorAction = 'Stop'
        WarningAction = 'SilentlyContinue'
        InformationAction = 'SilentlyContinue'
    }

    if (-not [string]::IsNullOrWhiteSpace($effectiveTenant)) {
        $baseParams['Tenant'] = $effectiveTenant
    }
    if (-not [string]::IsNullOrWhiteSpace($effectivePnPClientId)) {
        $baseParams['ClientId'] = $effectivePnPClientId
    }

    $attempts = @(
        @{ Name = 'Interactive'; ParameterName = 'Interactive' },
        @{ Name = 'DeviceLogin'; ParameterName = 'DeviceLogin' },
        @{ Name = 'OSLogin'; ParameterName = 'OSLogin' }
    )

    foreach ($attempt in $attempts) {
        $connectParams = @{}
        foreach ($key in $baseParams.Keys) {
            $connectParams[$key] = $baseParams[$key]
        }
        $connectParams[$attempt.ParameterName] = $true

        try {
            Connect-PnPOnline @connectParams
            return [ordered]@{
                Succeeded = $true
                Method = $attempt.Name
                PnPClientIdUsed = if (-not [string]::IsNullOrWhiteSpace($effectivePnPClientId)) { $effectivePnPClientId } else { 'not provided' }
            }
        }
        catch {
            $attemptErrors.Add("$($attempt.Name): $($_.Exception.Message)") | Out-Null
        }
    }

    $combinedReason = if ($attemptErrors.Count -gt 0) {
        $attemptErrors -join ' | '
    }
    else {
        'Unknown PnP authentication failure.'
    }

    $requiresClientIdHint = ($combinedReason -match 'Unable to connect using provided arguments') -or ($combinedReason -match 'ClientId')
    if ($requiresClientIdHint -and [string]::IsNullOrWhiteSpace($effectivePnPClientId)) {
        $combinedReason = "$combinedReason. PnP interactive auth may require an explicit Entra app id. Rerun with -PnPClientId <app-id> or set ENTRAID_APP_ID in the environment."
    }

    return [ordered]@{
        Succeeded = $false
        Reason = $combinedReason
        PnPClientIdUsed = if (-not [string]::IsNullOrWhiteSpace($effectivePnPClientId)) { $effectivePnPClientId } else { 'not provided' }
    }
}

function Invoke-SharePointChildCollector {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AdminUrl,
        [Parameter(Mandatory = $true)]
        [switch]$Sample,
        [Parameter(Mandatory = $true)]
        [int]$Max
    )

    $childResult = $null
    $connectFailed = $false

    try {
        Import-SilentPnPModule | Out-Null
    }
    catch {
        return [ordered]@{
            Status                = 'not collected'
            Reason                = "PnP.PowerShell could not be loaded in the child SharePoint process: $($_.Exception.Message)"
            ExternalSharingLevel  = 'not collected'
            DefaultSharingLinkType = 'not collected'
            RestrictedSearchEnabled = 'not collected'
            TotalSiteCount        = 'not collected'
            Truncated             = 'not collected'
            Sites                 = @()
            OneDriveCoveragePercent = 'not collected'
            InactiveSiteCount     = 'not collected'
            OversharingSignals    = [ordered]@{
                Status                    = 'not collected'
                CollectionMethod          = 'not collected'
                CollectionScope           = 'not collected'
                CollectionPathUsed        = 'not collected'
                CollectionReason          = 'not collected'
                RequiredForReliableCollection = 'not collected'
                DocumentationNote         = 'not collected'
                SampledAnyoneLinkCount    = 'not collected'
                SampledOrgWideLinkCount   = 'not collected'
                SitesWithAnyoneLinksCount = 'not collected'
            }
        }
    }

    $connectResult = Connect-PnPForSharePointCollection -AdminUrl $AdminUrl
    if (-not $connectResult.Succeeded) {
            $connectFailed = $true
            return [ordered]@{
                Status                = 'not collected'
                Reason                = "SharePoint / PnP connection failed in the isolated child process: $($connectResult.Reason)"
                ExternalSharingLevel  = 'not collected'
                DefaultSharingLinkType = 'not collected'
                RestrictedSearchEnabled = 'not collected'
                TotalSiteCount        = 'not collected'
                Truncated             = 'not collected'
                Sites                 = @()
                OneDriveCoveragePercent = 'not collected'
                InactiveSiteCount     = 'not collected'
                OversharingSignals    = [ordered]@{
                    Status                    = 'not collected'
                    CollectionMethod          = 'not collected'
                    CollectionScope           = 'not collected'
                    CollectionPathUsed        = 'not collected'
                    CollectionReason          = 'not collected'
                    RequiredForReliableCollection = 'not collected'
                    DocumentationNote         = 'not collected'
                    SampledAnyoneLinkCount    = 'not collected'
                    SampledOrgWideLinkCount   = 'not collected'
                    SitesWithAnyoneLinksCount = 'not collected'
                }
            }
    }

    if (-not $connectFailed) {
        $childResult = Get-SharePointOneDriveReadinessLocal -Sample:$Sample -Max $Max
    }

    return $childResult
}

function Import-SilentPnPModule {
    [CmdletBinding()]
    param()

    $alreadyLoaded = Get-Module -Name PnP.PowerShell | Select-Object -First 1
    if ($null -ne $alreadyLoaded) {
        return $alreadyLoaded
    }

    $candidateModules = @(Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending)
    if ($candidateModules.Count -eq 0) {
        throw 'PnP.PowerShell module is required.'
    }

    $preferredModule = $candidateModules | Where-Object { $_.Path -notlike '*WindowsPowerShell\Modules*' } | Select-Object -First 1
    if ($null -eq $preferredModule) {
        Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop | Out-Null
        $candidateModules = @(Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending)
        $preferredModule = $candidateModules | Where-Object { $_.Path -notlike '*WindowsPowerShell\Modules*' } | Select-Object -First 1
        if ($null -eq $preferredModule) {
            throw 'Unable to locate a PowerShell 7 compatible PnP.PowerShell installation.'
        }
    }

    Import-Module -Name $preferredModule.Path -ErrorAction Stop | Out-Null
    return (Get-Module -Name PnP.PowerShell | Select-Object -First 1)
}

function Get-PnPAssemblyConflictState {
    $tenantAssemblies = @(
        [AppDomain]::CurrentDomain.GetAssemblies() |
            Where-Object { $_.GetName().Name -eq 'Microsoft.Online.SharePoint.Client.Tenant' }
    )

    if ($tenantAssemblies.Count -eq 0) {
        return [ordered]@{
            ConflictDetected = $false
            Reason = 'not loaded'
            LoadedVersion = ''
            ExpectedVersion = '16.1.0.0'
            Location = ''
        }
    }

    $loadedVersion = [string]$tenantAssemblies[0].GetName().Version
    $location = ''
    try { $location = [string]$tenantAssemblies[0].Location } catch {}

    $conflictDetected = ($loadedVersion -ne '16.1.0.0')

    return [ordered]@{
        ConflictDetected = $conflictDetected
        Reason = if ($conflictDetected) { 'version mismatch' } else { 'ok' }
        LoadedVersion = $loadedVersion
        ExpectedVersion = '16.1.0.0'
        Location = $location
    }
}

function Write-PnPAssemblyStartupGuidance {
    $state = Get-PnPAssemblyConflictState
    if (-not $state.ConflictDetected) { return }

    $relaunchCommand = Get-CleanRelaunchCommand
    Write-Warning ("Startup check: SharePoint CSOM assembly conflict detected (Microsoft.Online.SharePoint.Client.Tenant loaded=$($state.LoadedVersion), expected=$($state.ExpectedVersion)). Relaunch clean with: $relaunchCommand")
}

$script:Report = [ordered]@{
    Metadata                = $null
    Licensing               = $null
    Identity                = $null
    ConditionalAccess       = $null
    DataGovernance          = $null
    SharePointOneDrive      = $null
    Exchange                = $null
    Teams                   = $null
    SearchIndex             = $null
    Apps                    = $null
    SecureScore             = $null
    AdoptionSignals         = $null
    IdentityAccessAdvanced  = $null
    DataProtectionAdvanced  = $null
    SharePointExposureAdvanced = $null
    TeamsAdvanced           = $null
    SearchSemanticAdvanced  = $null
    EndpointAppAdvanced     = $null
    AdoptionAdvanced        = $null
    AppGovernanceAdvanced   = $null
    ReadinessFlags          = $null
    ReadinessEvidence       = $null
    Recommendations         = $null
    PrerequisitesAndGaps    = $null
    CollectorTimings        = [ordered]@{}
    Errors                  = [System.Collections.Generic.List[object]]::new()
}
$script:GraphAccessToken = $null
$script:GraphAccessTokenExpiresAtUtc = $null
$script:MgContext = $null
$script:TeamsConnected = $false
$script:SharePointChildCollectionSucceeded = $false
$script:GraphMissingScopes = @()
$script:GraphProvidedScopes = @()
$script:GraphRestAllCache = @{}
$script:GraphUsersCache = @{}
$script:GraphReportRowsCache = @{}
$script:CopilotUsageCache = @{}
$script:GraphRestMaxPageCount = 500

function Initialize-GraphRuntimeState {
    $graphRestAllCacheVar = Get-Variable -Name GraphRestAllCache -Scope Script -ErrorAction SilentlyContinue
    if ($null -eq $graphRestAllCacheVar -or $null -eq $graphRestAllCacheVar.Value) {
        $script:GraphRestAllCache = @{}
    }

    $graphUsersCacheVar = Get-Variable -Name GraphUsersCache -Scope Script -ErrorAction SilentlyContinue
    if ($null -eq $graphUsersCacheVar -or $null -eq $graphUsersCacheVar.Value) {
        $script:GraphUsersCache = @{}
    }

    $graphReportRowsCacheVar = Get-Variable -Name GraphReportRowsCache -Scope Script -ErrorAction SilentlyContinue
    if ($null -eq $graphReportRowsCacheVar -or $null -eq $graphReportRowsCacheVar.Value) {
        $script:GraphReportRowsCache = @{}
    }

    $copilotUsageCacheVar = Get-Variable -Name CopilotUsageCache -Scope Script -ErrorAction SilentlyContinue
    if ($null -eq $copilotUsageCacheVar -or $null -eq $copilotUsageCacheVar.Value) {
        $script:CopilotUsageCache = @{}
    }

    $graphRestMaxPageCountVar = Get-Variable -Name GraphRestMaxPageCount -Scope Script -ErrorAction SilentlyContinue
    if ($null -eq $graphRestMaxPageCountVar -or $null -eq $graphRestMaxPageCountVar.Value) {
        $script:GraphRestMaxPageCount = 500
    }
}

function Get-GraphTokenClaims {
    param([Parameter(Mandatory)][string]$AccessToken)

    $parts = $AccessToken -split '\.'
    if ($parts.Count -lt 2) { return $null }

    $payload = $parts[1]
    # JWT payload uses base64url, so normalize before decoding.
    $payload = $payload.Replace('-', '+').Replace('_', '/')
    $padding = (4 - ($payload.Length % 4)) % 4
    $base64 = $payload + ('=' * $padding)

    try {
        $jsonBytes = [Convert]::FromBase64String($base64)
        $jsonText = [System.Text.Encoding]::UTF8.GetString($jsonBytes)
        return $jsonText | ConvertFrom-Json -AsHashtable
    }
    catch {
        return $null
    }
}

function Get-GraphTokenExpirationUtc {
    param([string]$AccessToken)

    if ([string]::IsNullOrWhiteSpace($AccessToken)) { return $null }

    $claims = Get-GraphTokenClaims -AccessToken $AccessToken
    if (-not $claims -or -not $claims.ContainsKey('exp')) { return $null }

    try {
        $exp = [int64]$claims['exp']
        return ([datetimeoffset]::FromUnixTimeSeconds($exp)).UtcDateTime
    }
    catch {
        return $null
    }
}

function Test-HashtableKey {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$Key
    )

    if ($null -eq $InputObject) { return $false }
    if ($InputObject -is [System.Collections.IDictionary]) {
        return $InputObject.Contains($Key)
    }

    return $null -ne $InputObject.PSObject.Properties[$Key]
}

function Get-ObjectPropertyValue {
    # Safe accessor for hashtable keys and object properties under StrictMode.
    # Returns $Default when the member is missing or its value is $null.
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()][object]$Default = $null
    )

    if ($null -eq $InputObject) { return $Default }

    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name) -and $null -ne $InputObject[$Name]) {
            return $InputObject[$Name]
        }
        return $Default
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        return $property.Value
    }

    return $Default
}

function Test-SectionCollected {
    param([AllowNull()][object]$SectionData)

    # Section-level heuristic only: this does not inspect nested fields for partial "not collected" values.
    if ($null -eq $SectionData) { return $false }

    if ($SectionData -is [string]) {
        return -not ([string]$SectionData -like 'not collected*')
    }

    if (Test-HashtableKey -InputObject $SectionData -Key 'Status') {
        $status = [string](Get-ObjectPropertyValue -InputObject $SectionData -Name 'Status' -Default '')
        if ($status -like 'not collected*') { return $false }
    }

    return $true
}

function Test-ValueCollected {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return $false }
    if ($Value -is [string]) {
        return -not ([string]$Value -like 'not collected*')
    }

    return $true
}

function Get-FirstIntPropertyValue {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string[]]$CandidateNames
    )

    if ($null -eq $InputObject) { return $null }

    foreach ($name in @($CandidateNames)) {
        if ($InputObject.PSObject.Properties.Name -contains $name) {
            $raw = $InputObject.$name
            if ($null -eq $raw -or [string]::IsNullOrWhiteSpace([string]$raw)) { continue }
            try {
                return [int]$raw
            }
            catch {
                continue
            }
        }
    }

    return $null
}

function Get-IntValue {
    param(
        [AllowNull()][object]$Value,
        [int]$Default = 0
    )

    if ($null -eq $Value) { return $Default }

    $parsedInt = 0
    if ([int]::TryParse([string]$Value, [ref]$parsedInt)) {
        return $parsedInt
    }

    $parsedDouble = 0.0
    if ([double]::TryParse([string]$Value, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedDouble)) {
        return [int][math]::Round($parsedDouble, 0)
    }

    if ([double]::TryParse([string]$Value, [ref]$parsedDouble)) {
        return [int][math]::Round($parsedDouble, 0)
    }

    return $Default
}

function Get-DoubleValue {
    param(
        [AllowNull()][object]$Value,
        [double]$Default = 0
    )

    if ($null -eq $Value) { return $Default }

    $parsed = 0.0
    if ([double]::TryParse([string]$Value, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }

    if ([double]::TryParse([string]$Value, [ref]$parsed)) {
        return $parsed
    }

    return $Default
}

function ConvertTo-NormalizedGraphPermissionName {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }

    $normalized = [string]$Value.Trim()
    $normalized = $normalized -replace '^https://graph\.microsoft\.com/', ''
    $normalized = $normalized -replace '^https://graph\.microsoft\.com$', ''
    $normalized = $normalized.ToLowerInvariant()

    return $normalized
}

function Get-RequiredGraphScopes {
    $scopes = @(
        'AuditLog.Read.All'
        'DeviceManagementManagedDevices.Read.All'
        'Directory.Read.All'
        'ExternalConnection.Read.All'
        'Group.Read.All'
        'InformationProtectionPolicy.Read.All'
        'Organization.Read.All'
        'Policy.Read.All'
        'Reports.Read.All'
        'SecurityEvents.Read.All'
    )

    return @($scopes | Sort-Object -Unique)
}

function Get-DefaultGraphScopes {
    return @(Get-RequiredGraphScopes)
}

function Test-GraphPermissionsAvailable {
    param([string[]]$RequiredPermissions)

    if ($null -eq $RequiredPermissions -or @($RequiredPermissions).Count -eq 0) {
        return $true
    }

    $missingPermissions = @(Get-GraphMissingScopes -RequiredScopes $RequiredPermissions)
    return $missingPermissions.Count -eq 0
}

function Get-DefaultScoringModel {
    return [ordered]@{
        IdentityAccess = [ordered]@{
            MfaRegisteredWeight = 40
            MfaRegisteredPercentThreshold = 80
            AdvancedMfaWeight = 30
            AdvancedMfaPercentThreshold = 20
            ConditionalAccessWeight = 30
            ConditionalAccessPolicyThreshold = 1
        }
        DataGovernance = [ordered]@{
            LabelsPublishedWeight = 34
            DlpEnforcedWeight = 33
            RetentionEnabledWeight = 33
        }
        ContentExposure = [ordered]@{
            OneDriveCoverageWeight = 34
            OneDriveCoverageThreshold = 70
            NoAnyoneLinksWeight = 33
            AnyoneLinksPresentWeight = 0
            RestrictedSearchWeight = 33
        }
        Adoption = [ordered]@{
            ActiveUserHighWeight = 70
            ActiveUserHighThreshold = 60
            ActiveUserMediumThreshold = 35
            ActiveUserMediumWeight = 40
            ActiveUserLowWeight = 15
            TrendReportedWeight = 30
        }
        EndpointReadiness = [ordered]@{
            ComplianceHighThreshold = 90
            ComplianceMediumThreshold = 75
            HighScore = 100
            MediumScore = 70
            LowScore = 35
        }
        Confidence = [ordered]@{
            EvaluatedSections = @(
                'Identity',
                'ConditionalAccess',
                'DataGovernance',
                'SharePointOneDrive',
                'AdoptionSignals',
                'IdentityAccessAdvanced',
                'DataProtectionAdvanced',
                'SharePointExposureAdvanced',
                'AppGovernanceAdvanced',
                'EndpointAppAdvanced'
            )
            HighThreshold = 85
            MediumThreshold = 60
        }
    }
}

function Get-GraphMissingScopes {
    param([string[]]$RequiredScopes)

    $providedScopes = New-Object System.Collections.Generic.List[string]
    $ctx = Get-MgContext -ErrorAction SilentlyContinue

    if ($ctx -and $ctx.PSObject.Properties.Name -contains 'Scopes' -and $null -ne $ctx.Scopes) {
        foreach ($scope in @($ctx.Scopes)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$scope)) { $providedScopes.Add([string]$scope) }
        }
    }

    $token = if ($script:GraphAccessToken) { $script:GraphAccessToken } else { if ($ctx -and $ctx.PSObject.Properties.Name -contains 'AccessToken') { $ctx.AccessToken } else { $null } }
    if (-not [string]::IsNullOrWhiteSpace($token)) {
        $claims = Get-GraphTokenClaims -AccessToken $token
        if ($claims) {
            if ($claims.ContainsKey('scp')) {
                foreach ($scope in @(($claims['scp'] -split ' '))) {
                    if (-not [string]::IsNullOrWhiteSpace($scope)) { $providedScopes.Add($scope) }
                }
            }
            if ($claims.ContainsKey('roles')) {
                foreach ($role in @($claims['roles'])) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$role)) { $providedScopes.Add([string]$role) }
                }
            }
        }
    }

    $providedScopes = @($providedScopes | ForEach-Object { ConvertTo-NormalizedGraphPermissionName -Value $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    $script:GraphProvidedScopes = @($providedScopes)
    $requiredScopeList = @($RequiredScopes | ForEach-Object { ConvertTo-NormalizedGraphPermissionName -Value $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

    return @($requiredScopeList | Where-Object { $_ -notin $providedScopes } | Sort-Object -Unique)
}

function Get-GraphAdminConsentUrl {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId
    )

    $targetTenant = if ([string]::IsNullOrWhiteSpace($TenantId)) { 'organizations' } else { $TenantId }
    return "https://login.microsoftonline.com/$targetTenant/v2.0/adminconsent?client_id=$ClientId"
}

function Get-GraphErrorPayload {
    param([Parameter(Mandatory)][object]$ErrorRecord)

    if ($null -eq $ErrorRecord) { return $null }

    $raw = $null
    if ($ErrorRecord.PSObject.Properties.Name -contains 'ErrorDetails' -and $null -ne $ErrorRecord.ErrorDetails) {
        $raw = [string]$ErrorRecord.ErrorDetails.Message
    }

    if ([string]::IsNullOrWhiteSpace($raw)) {
        $message = [string]$ErrorRecord.Exception.Message
        if ($message -match '\{.+\}') {
            $raw = $matches[0]
        }
    }

    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }

    try {
        return ($raw | ConvertFrom-Json -ErrorAction Stop)
    }
    catch {
        return $null
    }
}

function Get-GraphResourceServicePrincipalId {
    $graphAppId = '00000003-0000-0000-c000-000000000000'
    $sp = @(Get-GraphRestAll -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$graphAppId'") | Select-Object -First 1
    if ($sp) { return [string]$sp.id }
    return $null
}

function Get-GraphClientServicePrincipalId {
    param([Parameter(Mandatory)][string]$ClientAppId)

    $sp = @(Get-GraphRestAll -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$ClientAppId'") | Select-Object -First 1
    if ($sp) { return [string]$sp.id }
    return $null
}

function Confirm-GraphConsent {
    param([string[]]$RequiredScopes)

    if ([string]::IsNullOrWhiteSpace($GraphClientId)) { return $false }

    $clientServicePrincipalId = Get-GraphClientServicePrincipalId -ClientAppId $GraphClientId
    if ([string]::IsNullOrWhiteSpace($clientServicePrincipalId)) {
        Write-Step 'No client service principal exists in this tenant for the selected Graph client ID, so consent grant creation is skipped.' -Color DarkYellow
        return $false
    }

    $resourceId = Get-GraphResourceServicePrincipalId
    if ([string]::IsNullOrWhiteSpace($resourceId)) {
        Write-Warning 'Microsoft Graph resource principal could not be resolved, so consent status cannot be evaluated.'
        return $false
    }

    $scopeString = @($RequiredScopes |
        ForEach-Object { ConvertTo-NormalizedGraphPermissionName -Value $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique) -join ' '

    if ([string]::IsNullOrWhiteSpace($scopeString)) {
        return $false
    }

    try {
        $existing = @(Get-GraphRestAll -Uri "https://graph.microsoft.com/beta/oauth2PermissionGrants?`$filter=clientId eq '$clientServicePrincipalId' and resourceId eq '$resourceId'") | Select-Object -First 1
        if ($existing) {
            Write-Step 'A Graph consent grant for this app already exists. This script will not modify consent.' -Color Green
            return $true
        }

        $consentUrl = Get-GraphAdminConsentUrl -TenantId $TenantId -ClientId $GraphClientId
        Write-Step 'Missing Graph consent was detected. Read-only mode prevents tenant consent changes.' -Color DarkYellow
        Write-Step "Admin action required: grant consent manually here, then rerun this script: $consentUrl" -Color DarkYellow
        return $false
    }
    catch {
        Write-Warning "Graph consent status check failed: $($_.Exception.Message)"
        return $false
    }
}

function Test-GraphScopeValidation {
    param([string[]]$RequiredScopes)

    $missing = @(Get-GraphMissingScopes -RequiredScopes $RequiredScopes)
    $providedScopes = @($script:GraphProvidedScopes)

    if ($providedScopes.Count -eq 0) {
        Write-Verbose 'Graph scope validation could not confirm the granted scopes from the current session; some collectors may return not collected.'
        return $false
    }

    $script:GraphMissingScopes = @($missing | Sort-Object -Unique)
    if ($missing.Count -gt 0) {
        Write-Verbose ("Graph scope validation detected missing permissions: {0}. Partial data is expected for collectors that depend on those scopes." -f ($missing -join ', '))
        return $false
    }

    Write-Step ("Graph scope validation passed with {0} granted scope(s)." -f $providedScopes.Count) -Color Green
    return $true
}

# ----------------------------------------------------------------------------
# Helper: run a collector safely, log failures, keep going.
# ----------------------------------------------------------------------------
function Invoke-Collector {
    param(
        [Parameter(Mandatory)][string]$Section,
        [Parameter(Mandatory)][scriptblock]$Action
    )
    $script:CollectorStepIndex++
    if ($script:CollectorStepTotal -gt 0) {
        Write-Step ("[{0}/{1}] Reading section '{2}'..." -f $script:CollectorStepIndex, $script:CollectorStepTotal, $Section) -Color Cyan
    }
    else {
        Write-Step "Reading section '$Section'..." -Color Cyan
    }
    $startedUtc = (Get-Date).ToUniversalTime()
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $result = & $Action
        if ($null -eq $result) {
            $result = [ordered]@{ Status = 'not collected' }
        }
        $script:Report[$Section] = $result
        $stopwatch.Stop()
        $script:Report.CollectorTimings[$Section] = [ordered]@{
            Status      = 'completed'
            StartedUtc  = $startedUtc.ToString('o')
            FinishedUtc = (Get-Date).ToUniversalTime().ToString('o')
            DurationMs  = [int]$stopwatch.ElapsedMilliseconds
        }
        Write-Step "Completed section '$Section' in $([int]$stopwatch.ElapsedMilliseconds) ms." -Color Green
    }
    catch {
        $stopwatch.Stop()
        $details = $_ | Out-String -Width 200
        $script:Report.Errors.Add([ordered]@{
            Section     = $Section
            Message     = $_.Exception.Message
            Exception   = $_.Exception.GetType().FullName
            Detail      = $details.Trim()
            TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
        })
        $script:Report.CollectorTimings[$Section] = [ordered]@{
            Status      = 'failed'
            StartedUtc  = $startedUtc.ToString('o')
            FinishedUtc = (Get-Date).ToUniversalTime().ToString('o')
            DurationMs  = [int]$stopwatch.ElapsedMilliseconds
            Message     = $_.Exception.Message
        }
        Write-Warning "Collector failed for '$Section': $($_.Exception.Message)"
    }
}

# ----------------------------------------------------------------------------
# Prerequisites
# ----------------------------------------------------------------------------
function Test-RequiredModules {
    Write-Step 'Checking Graph auth availability...' -Color Yellow
    $hasGraphCmd = [bool](Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)
    $hasContext = [bool](Get-MgContext -ErrorAction SilentlyContinue)

    if (-not $hasGraphCmd) {
        Write-Warning 'Microsoft Graph PowerShell cmdlets are not available; using REST calls only.'
    }
    elseif (-not $hasContext) {
        Write-Step 'No existing Microsoft Graph session was found; REST collectors will continue and prompt only if needed.' -Color DarkYellow
    }

    return ($hasGraphCmd -or $hasContext)
}

# ----------------------------------------------------------------------------
# Graph REST helpers
# ----------------------------------------------------------------------------
function Get-GraphAccessToken {
    [CmdletBinding()]
    param(
        [string]$TenantId = 'organizations',
        [string]$ClientId = '04b07795-8ddb-461a-bbee-02f9e1bf7b46',
        [string[]]$Scopes = @('https://graph.microsoft.com/.default')
    )

    if ([string]::IsNullOrWhiteSpace($TenantId)) { $TenantId = 'organizations' }

    if ($script:GraphAccessToken -and $script:GraphAccessTokenExpiresAtUtc -and $script:GraphAccessTokenExpiresAtUtc -gt (Get-Date).ToUniversalTime().AddMinutes(1)) {
        return $script:GraphAccessToken
    }

    if ($script:GraphAccessToken -and -not $script:GraphAccessTokenExpiresAtUtc) {
        $cachedExpiresAt = Get-GraphTokenExpirationUtc -AccessToken $script:GraphAccessToken
        if ($cachedExpiresAt -and $cachedExpiresAt -gt (Get-Date).ToUniversalTime().AddMinutes(1)) {
            $script:GraphAccessTokenExpiresAtUtc = $cachedExpiresAt
            return $script:GraphAccessToken
        }
    }

    $deviceCodeUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode"
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $scopeString = $Scopes -join ' '

    $deviceCodeResponse = Invoke-RestMethod -Method Post -Uri $deviceCodeUrl -Body @{ client_id = $ClientId; scope = $scopeString } -ErrorAction Stop

    $verificationUri = [string]$deviceCodeResponse.verification_uri
    $userCode = [string]$deviceCodeResponse.user_code

    if (-not [string]::IsNullOrWhiteSpace($userCode)) {
        try {
            if (Get-Command Set-Clipboard -ErrorAction SilentlyContinue) {
                Set-Clipboard -Value $userCode
                Write-Step "Device code copied to the clipboard: $userCode" -Color Green
            }
            else {
                Write-Step "Clipboard command unavailable. Use this device code: $userCode" -Color Yellow
            }
        }
        catch {
            Write-Step "Could not copy device code to clipboard. Use this code manually: $userCode" -Color Yellow
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($verificationUri)) {
        try {
            Start-Process -FilePath $verificationUri | Out-Null
            Write-Step "Opened browser for authentication: $verificationUri" -Color Cyan
        }
        catch {
            Write-Step "Could not open browser automatically. Open this URL manually: $verificationUri" -Color Yellow
        }
    }

    $token = $null
    $pollIntervalSeconds = [math]::Max(1, [int]$deviceCodeResponse.interval)
    $body = @{ client_id = $ClientId; grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; device_code = $deviceCodeResponse.device_code; scope = $scopeString }
    $expires = (Get-Date).AddSeconds($deviceCodeResponse.expires_in)

    while ($null -eq $token -and (Get-Date) -lt $expires) {
        Start-Sleep -Seconds $pollIntervalSeconds
        try {
            $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ErrorAction Stop
            $token = $tokenResponse.access_token
        }
        catch {
            if ($_.Exception.Response -and ($_.Exception.Response.StatusCode -eq 400)) {
                $errorJson = Get-GraphErrorPayload -ErrorRecord $_
                if ($null -eq $errorJson -or $errorJson.error -ne 'authorization_pending') {
                    throw $_
                }
            }
            else {
                throw $_
            }
        }
    }

    if ($null -eq $token) { throw 'Failed to acquire access token via device code flow.' }

    $script:GraphAccessToken = $token
    $script:GraphAccessTokenExpiresAtUtc = Get-GraphTokenExpirationUtc -AccessToken $token
    return $token
}

function Add-GraphQueryString {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Query
    )

    if (-not $Query -or $Query.Count -eq 0) { return $Uri }

    $queryString = foreach ($key in $Query.Keys) {
        $encodedKey = [System.Uri]::EscapeDataString([string]$key)
        $encodedValue = [System.Uri]::EscapeDataString([string]$Query[$key])
        "${encodedKey}=${encodedValue}"
    }

    $separator = if ($Uri -match '\?') { '&' } else { '?' }
    return "${Uri}${separator}$($queryString -join '&')"
}

function Get-GraphRetryDelaySeconds {
    param(
        [AllowNull()][object]$ErrorRecord,
        [int]$Attempt = 0
    )

    $fallbackDelay = [math]::Min(60, [int][math]::Pow(2, ([math]::Max(0, $Attempt) + 1)))
    if ($null -eq $ErrorRecord) { return $fallbackDelay }

    $response = $null
    if ($ErrorRecord.PSObject.Properties.Name -contains 'Exception' -and $null -ne $ErrorRecord.Exception) {
        if ($ErrorRecord.Exception.PSObject.Properties.Name -contains 'Response') {
            $response = $ErrorRecord.Exception.Response
        }
    }

    if ($null -eq $response) { return $fallbackDelay }

    $headerCandidates = @()
    if ($response.PSObject.Properties.Name -contains 'Headers' -and $null -ne $response.Headers) {
        $headerCandidates += @($response.Headers)
    }

    foreach ($headers in @($headerCandidates)) {
        foreach ($headerName in @('Retry-After', 'retry-after')) {
            try {
                $headerValues = $null
                $headerValue = $null

                if ($headers -is [System.Collections.IDictionary]) {
                    $headerValue = $headers[$headerName]
                }
                elseif ($headers.PSObject.Methods.Name -contains 'TryGetValues') {
                    if (-not $headers.TryGetValues($headerName, [ref]$headerValues)) { continue }
                    $headerValue = @($headerValues) | Select-Object -First 1
                }

                if ($null -eq $headerValue) { continue }

                $seconds = 0
                if ([int]::TryParse([string]$headerValue, [ref]$seconds) -and $seconds -gt 0) {
                    return $seconds
                }

                $retryAt = [datetimeoffset]::MinValue
                if ([datetimeoffset]::TryParse([string]$headerValue, [ref]$retryAt)) {
                    $delaySeconds = [int][math]::Ceiling(($retryAt.UtcDateTime - (Get-Date).ToUniversalTime()).TotalSeconds)
                    if ($delaySeconds -gt 0) { return $delaySeconds }
                }
            }
            catch {
                continue
            }
        }

        foreach ($headerName in @('x-ms-retry-after-ms', 'X-MS-Retry-After-MS')) {
            try {
                $headerValues = $null
                $headerValue = $null

                if ($headers -is [System.Collections.IDictionary]) {
                    $headerValue = $headers[$headerName]
                }
                elseif ($headers.PSObject.Methods.Name -contains 'TryGetValues') {
                    if (-not $headers.TryGetValues($headerName, [ref]$headerValues)) { continue }
                    $headerValue = @($headerValues) | Select-Object -First 1
                }

                if ($null -eq $headerValue) { continue }

                $milliseconds = 0
                if ([int]::TryParse([string]$headerValue, [ref]$milliseconds) -and $milliseconds -gt 0) {
                    return [math]::Max(1, [int][math]::Ceiling($milliseconds / 1000.0))
                }
            }
            catch {
                continue
            }
        }
    }

    return $fallbackDelay
}

function Get-CachedGraphUsers {
    param(
        [switch]$IncludeSignInActivity
    )

    $baseCacheKey = 'base'
    if (-not $script:GraphUsersCache.ContainsKey($baseCacheKey)) {
        $baseUsers = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/users' -Query @{ '$select' = 'id,userPrincipalName,userType,accountEnabled,assignedPlans,mail' }))
        $script:GraphUsersCache[$baseCacheKey] = $baseUsers
    }

    if (-not $IncludeSignInActivity) {
        return @($script:GraphUsersCache[$baseCacheKey])
    }

    $signInCacheKey = 'withSignInActivity'
    if ($script:GraphUsersCache.ContainsKey($signInCacheKey)) {
        return @($script:GraphUsersCache[$signInCacheKey])
    }

    $signInUsers = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/users' -Query @{ '$select' = 'id,signInActivity' }))
    $signInById = @{}
    foreach ($user in @($signInUsers)) {
        if ($null -eq $user) { continue }
        if (-not ($user.PSObject.Properties.Name -contains 'Id')) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$user.Id)) { continue }

        $signInById[[string]$user.Id] = $user
    }

    $mergedUsers = foreach ($baseUser in @($script:GraphUsersCache[$baseCacheKey])) {
        $mergedUser = [pscustomobject]([ordered]@{})
        foreach ($property in @($baseUser.PSObject.Properties)) {
            $mergedUser | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value
        }

        $signInUser = $null
        if ($baseUser.PSObject.Properties.Name -contains 'Id' -and $signInById.ContainsKey([string]$baseUser.Id)) {
            $signInUser = $signInById[[string]$baseUser.Id]
        }

        $signInActivityValue = $null
        if ($signInUser -and $signInUser.PSObject.Properties.Name -contains 'SignInActivity') {
            $signInActivityValue = $signInUser.SignInActivity
        }

        $mergedUser | Add-Member -NotePropertyName 'SignInActivity' -NotePropertyValue $signInActivityValue
        $mergedUser
    }

    $script:GraphUsersCache[$signInCacheKey] = @($mergedUsers)
    return @($script:GraphUsersCache[$signInCacheKey])
}

function Get-GraphReportRows {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Query = @{}
    )

    $requestUri = Add-GraphQueryString -Uri $Uri -Query $Query
    if ($script:GraphReportRowsCache.ContainsKey($requestUri)) {
        return @($script:GraphReportRowsCache[$requestUri])
    }

    $token = Get-GraphAccessToken
    if (-not $token) { return @() }

    $maxAttempts = 5
    for ($attempt = 0; $attempt -lt $maxAttempts; $attempt++) {
        try {
            $response = Invoke-WebRequest -Method Get -Uri $requestUri -Headers @{ Authorization = "Bearer $token"; Accept = 'text/csv' } -ErrorAction Stop
            $content = [string]$response.Content
            if ([string]::IsNullOrWhiteSpace($content)) { return @() }

            $trimmed = $content.Trim()

            # Some report endpoints return a pre-authenticated download URL as the response body.
            if ($trimmed -match '^https?://') {
                $downloadResponse = Invoke-WebRequest -Method Get -Uri $trimmed -ErrorAction Stop
                $content = [string]$downloadResponse.Content
                if ([string]::IsNullOrWhiteSpace($content)) { return @() }
                $trimmed = $content.Trim()
            }

            $rows = if ($trimmed.StartsWith('{') -or $trimmed.StartsWith('[')) {
                try {
                    $jsonPayload = $trimmed | ConvertFrom-Json -ErrorAction Stop
                    $valueProperty = if ($null -ne $jsonPayload) { $jsonPayload.PSObject.Properties['value'] } else { $null }
                    if ($null -ne $valueProperty) {
                        # OData-style payloads wrap the report rows in a 'value' array.
                        @($valueProperty.Value)
                    }
                    else {
                        @($jsonPayload)
                    }
                }
                catch {
                    @()
                }
            }
            else {
                @($content | ConvertFrom-Csv)
            }

            $script:GraphReportRowsCache[$requestUri] = $rows
            return @($rows)
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response -and $_.Exception.Response.PSObject.Properties.Name -contains 'StatusCode') {
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { $statusCode = $null }
            }

            if (($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)) -and $attempt -lt ($maxAttempts - 1)) {
                $delaySeconds = Get-GraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt
                Start-Sleep -Seconds $delaySeconds
                continue
            }

            return @()
        }
    }

    return @()
}

function Get-RowValueNormalized {
    param(
        [Parameter(Mandatory)][object]$Row,
        [Parameter(Mandatory)][string[]]$CandidateNames
    )

    $propertyMap = @{}
    foreach ($property in @($Row.PSObject.Properties)) {
        $normalized = ([string]$property.Name).ToLowerInvariant() -replace '[^a-z0-9]', ''
        if (-not [string]::IsNullOrWhiteSpace($normalized)) {
            $propertyMap[$normalized] = $property.Value
        }
    }

    foreach ($candidate in @($CandidateNames)) {
        $normalizedCandidate = ([string]$candidate).ToLowerInvariant() -replace '[^a-z0-9]', ''
        if ($propertyMap.ContainsKey($normalizedCandidate)) {
            return $propertyMap[$normalizedCandidate]
        }
    }

    return $null
}

function Get-CopilotUsageSnapshot {
    param(
        [ValidateSet('D7', 'D30', 'D90')]
        [string]$Period = 'D30'
    )

    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('Reports.Read.All'))) {
        return [ordered]@{
            Status      = 'not collected'
            Reason      = 'missing Graph permission: Reports.Read.All'
            ActiveUsers = 'not collected'
            ReportDate  = 'not collected'
            RecordCount = 0
            Source      = 'not collected'
        }
    }

    if ($script:CopilotUsageCache.ContainsKey($Period)) {
        return $script:CopilotUsageCache[$Period]
    }

    $candidateReports = @(
        "https://graph.microsoft.com/v1.0/reports/getMicrosoft365CopilotUsageUserCounts(period='$Period')",
        "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserCounts(period='$Period')",
        "https://graph.microsoft.com/v1.0/reports/getMicrosoft365CopilotUsageUserDetail(period='$Period')",
        "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='$Period')"
    )

    foreach ($reportUri in $candidateReports) {
        $rows = @(Get-GraphReportRows -Uri $reportUri)
        if (@($rows).Count -eq 0) { continue }

        $latestRow = $rows |
            Sort-Object {
                $reportDateValue = Get-RowValueNormalized -Row $_ -CandidateNames @('reportDate', 'report refresh date')
                try { [datetime]$reportDateValue } catch { [datetime]::MinValue }
            } -Descending |
            Select-Object -First 1

        $activeUsersValue = Get-RowValueNormalized -Row $latestRow -CandidateNames @(
            'activeUsers',
            'active users',
            'microsoft365CopilotChatUsers',
            'microsoft 365 copilot chat users',
            'copilotStudioUsers',
            'copilot studio users',
            'totalUsers'
        )

        $activeUsers = $null
        if ($null -ne $activeUsersValue -and -not [string]::IsNullOrWhiteSpace([string]$activeUsersValue)) {
            try {
                $activeUsers = [int]$activeUsersValue
            }
            catch {
                $activeUsers = $null
            }
        }

        if ($null -eq $activeUsers -and $reportUri -match 'UserDetail') {
            # Detail reports are row-based (one user per row), so row count approximates active users.
            $activeUsers = @($rows).Count
        }

        $reportDate = Get-RowValueNormalized -Row $latestRow -CandidateNames @('reportDate', 'report refresh date')

        $result = [ordered]@{
            Status      = 'reported'
            Reason      = ''
            ActiveUsers = if ($null -ne $activeUsers) { [int]$activeUsers } else { 'not collected' }
            ReportDate  = if ($reportDate) { [string]$reportDate } else { 'not collected' }
            RecordCount = @($rows).Count
            Source      = if ($reportUri -match 'UserDetail') { 'microsoft365CopilotUsageUserDetail' } else { 'microsoft365CopilotUsageUserCounts' }
        }
        $script:CopilotUsageCache[$Period] = $result
        return $result
    }

    $result = [ordered]@{
        Status      = 'not collected'
        Reason      = 'Copilot usage report endpoint returned no data (not enabled, unsupported, or not licensed in this tenant).'
        ActiveUsers = 'not collected'
        ReportDate  = 'not collected'
        RecordCount = 0
        Source      = 'not collected'
    }
    $script:CopilotUsageCache[$Period] = $result
    return $result
}

function Get-GraphRest {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Query = @{}
    )

    $token = Get-GraphAccessToken
    if (-not $token) { return $null }

    $requestUri = Add-GraphQueryString -Uri $Uri -Query $Query

    $irmParams = @{
        Method      = 'Get'
        Uri         = $requestUri
        Headers     = @{ Authorization = "Bearer $token" }
        ErrorAction = 'Stop'
    }

    $maxAttempts = 5
    for ($attempt = 0; $attempt -lt $maxAttempts; $attempt++) {
        try {
            return Invoke-RestMethod @irmParams
        }
        catch {
            $message = $_.Exception.Message
            $statusCode = $null
            if ($_.Exception.Response -and $_.Exception.Response.PSObject.Properties.Name -contains 'StatusCode') {
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { $statusCode = $null }
            }

            if (($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)) -and $attempt -lt ($maxAttempts - 1)) {
                $delaySeconds = Get-GraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt
                Start-Sleep -Seconds $delaySeconds
                continue
            }

            $suppressWarning = (
                ($Uri -match 'identitySecurityDefaultsEnforcementPolicy' -and $message -match '403') -or
                ($Uri -match 'informationProtection/policy/labels' -and $message -match '400') -or
                ($Uri -match 'external/connections' -and $message -match '401') -or
                ($Uri -match 'reports/m365AppUserCounts' -and $message -match '400') -or
                ($Uri -match 'reports/getOffice365ActiveUserCounts' -and $message -match '403') -or
                ($Uri -match 'reports/office365ActiveUserDetail' -and $message -match '400') -or
                ($Uri -match 'deviceManagement/managedDevices' -and $message -match '403')
            )

            if (-not $suppressWarning) {
                Write-Warning "Graph REST call failed for ${Uri}: $message"
            }
            return $null
        }
    }

    return $null
}

function Get-GraphRestAll {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Query = @{},
        [switch]$BypassCache
    )

    Initialize-GraphRuntimeState

    $requestUri = Add-GraphQueryString -Uri $Uri -Query $Query
    if (-not $BypassCache -and $script:GraphRestAllCache.ContainsKey($requestUri)) {
        return @($script:GraphRestAllCache[$requestUri])
    }

    $items = New-Object System.Collections.Generic.List[object]
    $nextUri = $requestUri
    $pageCount = 0

    while ($nextUri) {
        $pageCount++
        if ($pageCount -gt [int]$script:GraphRestMaxPageCount) {
            Write-Warning "Graph REST pagination aborted for ${Uri}: exceeded max page count of $($script:GraphRestMaxPageCount)."
            break
        }

        $response = Get-GraphRest -Uri $nextUri -Query @{}
        if ($null -eq $response) { break }

        if ($response -is [System.Collections.IDictionary]) {
            if ($response.ContainsKey('value')) {
                foreach ($item in @($response['value'])) {
                    if ($null -ne $item) { $items.Add([pscustomobject]$item) }
                }
                $nextUri = if ($response.ContainsKey('@odata.nextLink')) { [string]$response['@odata.nextLink'] } else { $null }
            }
            else {
                if ($null -ne $response) { $items.Add([pscustomobject]$response) }
                break
            }
        }
        else {
            $props = @($response.PSObject.Properties | ForEach-Object { $_.Name })
            if ($props -contains 'value') {
                foreach ($item in @($response.value)) {
                    if ($null -ne $item) { $items.Add([pscustomobject]$item) }
                }
                $nextUri = if ($props -contains '@odata.nextLink') { [string]$response.'@odata.nextLink' } else { $null }
            }
            else {
                if ($null -ne $response) { $items.Add([pscustomobject]$response) }
                break
            }
        }

    }

    $results = @($items.ToArray())
    if (-not $BypassCache) {
        $script:GraphRestAllCache[$requestUri] = $results
    }

    return $results
}

function Get-UserRegistrationFeatureCount {
    param(
        [AllowNull()][object[]]$FeatureCounts,
        [Parameter(Mandatory)][string]$FeatureName
    )

    foreach ($featureCount in @($FeatureCounts)) {
        if ($null -eq $featureCount) { continue }
        if (-not ($featureCount.PSObject.Properties.Name -contains 'feature')) { continue }
        if ([string]$featureCount.feature -ne $FeatureName) { continue }

        if ($featureCount.PSObject.Properties.Name -contains 'userCount') {
            return Get-IntValue -Value $featureCount.userCount -Default 0
        }
    }

    return $null
}

function Get-UserRegistrationFeatureSummary {
    param(
        [string]$IncludedUserTypes = 'all',
        [string]$IncludedUserRoles = 'all'
    )

    $uri = "https://graph.microsoft.com/v1.0/reports/authenticationMethods/usersRegisteredByFeature(includedUserTypes='$IncludedUserTypes',includedUserRoles='$IncludedUserRoles')"
    return Get-GraphRest -Uri $uri
}

function Get-DirectoryUsersForReadiness {
    param(
        [switch]$IncludeSignInActivity
    )

    return @(Get-CachedGraphUsers -IncludeSignInActivity:$IncludeSignInActivity)
}

# ----------------------------------------------------------------------------
# Connections (read-only scopes)
# ----------------------------------------------------------------------------
function Connect-Services {
    Write-Step 'Connecting to Microsoft Graph and Microsoft 365 services...' -Color Yellow

    $teamsImportError = $null
    $pnpImportError = $null
    try { Import-Module MicrosoftTeams -ErrorAction Stop } catch { $teamsImportError = $_.Exception.Message }
    try { Import-Module PnP.PowerShell -ErrorAction Stop } catch { $pnpImportError = $_.Exception.Message }
    try { Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue } catch {}

    try {
        $isDefaultMsClient = ($GraphClientId -eq '04b07795-8ddb-461a-bbee-02f9e1bf7b46')
        $graphScopes = if ($GraphScopes -and $GraphScopes.Count -gt 0) {
            @($GraphScopes)
        }
        elseif ($RequestRequiredGraphScopesUpFront) {
            Get-DefaultGraphScopes
        }
        elseif ($isDefaultMsClient) {
            @('https://graph.microsoft.com/.default')
        }
        else {
            Get-DefaultGraphScopes
        }

        $existingContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($existingContext) {
            Write-Step 'Detected an existing Microsoft Graph session; continuing with the current token flow.' -Color Green
        }

        $existingToken = $null
        if ($script:GraphAccessToken) {
            $existingToken = $script:GraphAccessToken
        }
        elseif ($existingContext -and $existingContext.PSObject.Properties.Name -contains 'AccessToken') {
            # In Microsoft Graph PowerShell SDK v1, AccessToken may be available for reuse.
            # In SDK v2+, this property is not a reliable cached token source, so keep the fallback conservative.
            $existingToken = [string]$existingContext.AccessToken
        }

        $existingTokenExpiresAt = $null
        if (-not [string]::IsNullOrWhiteSpace($existingToken)) {
            $existingTokenExpiresAt = Get-GraphTokenExpirationUtc -AccessToken $existingToken
        }

        if ($RequestRequiredGraphScopesUpFront) {
            $existingToken = $null
            $existingTokenExpiresAt = $null
            $script:GraphAccessToken = $null
            $script:GraphAccessTokenExpiresAtUtc = $null
            Write-Step 'Upfront Graph scope mode is enabled; cached Graph tokens will not be reused.' -Color DarkYellow
        }

        if ($existingToken -and $existingTokenExpiresAt -and $existingTokenExpiresAt -gt (Get-Date).ToUniversalTime().AddMinutes(1)) {
            $shouldReuseExistingToken = $true

            if ($graphScopes -contains 'https://graph.microsoft.com/.default') {
                $shouldReuseExistingToken = $false
                Write-Step 'Cached Graph token will not be reused when the .default scope is requested; reauth will request the full delegated scope set.' -Color DarkYellow
            }
            else {
                $missingRequestedScopes = @(Get-GraphMissingScopes -RequiredScopes $graphScopes)
                if ($missingRequestedScopes.Count -gt 0) {
                    $shouldReuseExistingToken = $false
                    Write-Warning ("Cached Graph token is missing required scope(s): {0}. Reauth will request the full scope set." -f ($missingRequestedScopes -join ', '))
                }
            }

            if ($shouldReuseExistingToken) {
                $script:GraphAccessToken = $existingToken
                $script:GraphAccessTokenExpiresAtUtc = $existingTokenExpiresAt
                Write-Step 'Reusing the current Microsoft Graph access token from the existing session.' -Color Green
                Write-Step 'Microsoft Graph authentication is ready.' -Color Green
            }
            else {
                $script:GraphAccessToken = $null
                $script:GraphAccessTokenExpiresAtUtc = $null
            }
        }
        else {
            if ($RequestRequiredGraphScopesUpFront -and $isDefaultMsClient) {
                Write-Warning 'Upfront Graph scope mode works best with a custom app registration. The default Microsoft public client cannot reliably request this tenant-specific delegated scope set directly, so a custom GraphClientId is recommended.'
            }

            if ($isDefaultMsClient -and ($graphScopes -notcontains 'https://graph.microsoft.com/.default')) {
                Write-Warning 'The Microsoft first-party client ID cannot request tenant-specific delegated Graph scopes directly. Falling back to https://graph.microsoft.com/.default to avoid AADSTS65002.'
                $graphScopes = @('https://graph.microsoft.com/.default')
            }

            if ($graphScopes -contains 'https://graph.microsoft.com/.default') {
                Write-Step 'Using Graph default scope set for device-code authentication.' -Color DarkYellow
            }
            else {
                $consentUrl = Get-GraphAdminConsentUrl -TenantId $TenantId -ClientId $GraphClientId
                if ($RequestRequiredGraphScopesUpFront) {
                    Write-Step 'Requesting the exporter''s full delegated Graph scope set up front so consent can be granted before collection starts.' -Color DarkYellow
                }
                else {
                    Write-Step 'Requesting the required delegated Graph permissions so the tenant can consent to them if needed.' -Color DarkYellow
                }
                Write-Step "If sign-in fails due to missing consent, grant admin consent for the custom app registration: $consentUrl" -Color DarkYellow
            }

            Write-Step 'Starting device-code sign-in and opening the login page automatically...' -Color DarkYellow
            $null = Get-GraphAccessToken -TenantId $TenantId -ClientId $GraphClientId -Scopes $graphScopes
            Write-Step 'Connected to Microsoft Graph.' -Color Green
        }
    }
    catch {
        Write-Warning "Graph sign-in failed: $($_.Exception.Message)"
    }


    $script:TeamsConnected = $false
    try {
        $teamsModuleAvailable = [bool](Get-Module -ListAvailable -Name MicrosoftTeams -ErrorAction SilentlyContinue)
        $teamsConnectCommand = Get-Command Connect-MicrosoftTeams -ErrorAction SilentlyContinue
        $teamsCmdletAvailable = $null -ne $teamsConnectCommand

        if ($teamsCmdletAvailable) {
            $existingTeamsSession = $false
            if (-not $ForceInteractiveTeamsPnPLogin) {
                try {
                    if (Get-Command Get-CsTenant -ErrorAction SilentlyContinue) {
                        $null = Get-CsTenant -ErrorAction Stop
                        $existingTeamsSession = $true
                    }
                }
                catch {
                    $existingTeamsSession = $false
                }
            }

            if ($existingTeamsSession) {
                $script:TeamsConnected = $true
                Write-Step 'Reusing existing Microsoft Teams session.' -Color Green
            }
            else {
                if ($ForceInteractiveTeamsPnPLogin) {
                    Write-Step 'Force interactive mode enabled: starting a fresh Microsoft Teams sign-in attempt.' -Color DarkYellow
                }

            try {
                $teamsParams = @{ ErrorAction = 'Stop' }
                if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
                    $teamsParams['TenantId'] = $TenantId
                }

                $null = Connect-MicrosoftTeams @teamsParams -WarningAction SilentlyContinue -InformationAction SilentlyContinue 3>$null 6>$null
                $script:TeamsConnected = $true
                Write-Step 'Connected to Microsoft Teams for policy metadata.' -Color Green
            }
            catch {
                $script:TeamsConnected = $false
                Write-Warning "Teams sign-in was not available during startup: $($_.Exception.Message)"
            }
            }
        }
        else {
            if ($teamsModuleAvailable) {
                if (-not [string]::IsNullOrWhiteSpace($teamsImportError)) {
                    Write-Step ("MicrosoftTeams module is installed, but failed to import in this session: {0}" -f $teamsImportError) -Color DarkYellow
                }
                else {
                    Write-Step 'MicrosoftTeams module is installed, but Connect-MicrosoftTeams is not available in this session.' -Color DarkYellow
                }
            }
            else {
                Write-Step 'Microsoft Teams cmdlets are not available in this session because the MicrosoftTeams module is not installed.' -Color DarkYellow
            }
        }
    }
    catch {
        $script:TeamsConnected = $false
        Write-Warning "Teams connection check failed: $($_.Exception.Message)"
    }

    Write-Step 'SharePoint / PnP collection will run in a clean child PowerShell process.' -Color Cyan

    try {
        $exoConnected = $false
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $connections = @((Get-ConnectionInformation -ErrorAction SilentlyContinue))
            $exoConnected = @($connections | Where-Object {
                ($_.PSObject.Properties.Name -contains 'Name' -and [string]$_.Name -match 'ExchangeOnline') -or
                ($_.PSObject.Properties.Name -contains 'ConnectionUri' -and [string]$_.ConnectionUri -match 'outlook\.office365\.com')
            }).Count -gt 0
        }

        if (-not $exoConnected -and (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
            Write-Step 'Connecting to Exchange Online for mailbox and compliance signals...' -Color Cyan
            $exoParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
            if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
                $exoParams['Organization'] = $TenantId
            }
            Connect-ExchangeOnline @exoParams -WarningAction SilentlyContinue -InformationAction SilentlyContinue 3>$null 6>$null | Out-Null
            $exoConnected = $true
            Write-Step 'Connected to Exchange Online.' -Color Green
        }
        elseif ($exoConnected) {
            Write-Step 'Reusing existing Exchange Online session.' -Color Green
        }
        else {
            Write-Step 'ExchangeOnlineManagement cmdlets are not available in this session.' -Color DarkYellow
        }
    }
    catch {
        Write-Warning "Exchange Online sign-in was not available during startup: $($_.Exception.Message)"
    }

    try {
        $ippsConnected = $false
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $connections = @((Get-ConnectionInformation -ErrorAction SilentlyContinue))
            $ippsConnected = @($connections | Where-Object {
                ($_.PSObject.Properties.Name -contains 'Name' -and [string]$_.Name -match 'IPPSSession|SecurityCompliance') -or
                ($_.PSObject.Properties.Name -contains 'ConnectionUri' -and [string]$_.ConnectionUri -match 'ps\.compliance\.protection\.outlook\.com')
            }).Count -gt 0
        }

        if (-not $ippsConnected -and (Get-Command Connect-IPPSSession -ErrorAction SilentlyContinue)) {
            Write-Step 'Connecting to Microsoft Purview compliance session for DLP and retention signals...' -Color Cyan
            $ippsCommand = Get-Command Connect-IPPSSession -ErrorAction SilentlyContinue
            $ippsParams = @{ ErrorAction = 'Stop' }
            if ($ippsCommand -and $ippsCommand.Parameters.ContainsKey('ShowBanner')) {
                $ippsParams['ShowBanner'] = $false
            }
            if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
                $ippsParams['Organization'] = $TenantId
            }
            Connect-IPPSSession @ippsParams -WarningAction SilentlyContinue -InformationAction SilentlyContinue 3>$null 6>$null | Out-Null
            $ippsConnected = $true
            Write-Step 'Connected to Microsoft Purview compliance session.' -Color Green
        }
        elseif ($ippsConnected) {
            Write-Step 'Reusing existing Microsoft Purview compliance session.' -Color Green
        }
        else {
            Write-Step 'Compliance cmdlets are not available in this session.' -Color DarkYellow
        }
    }
    catch {
        Write-Warning "Compliance (IPPS) sign-in was not available during startup: $($_.Exception.Message)"
    }

    try { $script:MgContext = Get-MgContext -ErrorAction Stop } catch { $script:MgContext = $null }
}

function Disconnect-Services {
    Write-Step 'Disconnecting services and clearing sessions...' -Color Yellow
    try {
        if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
            $null = Disconnect-MgGraph -ErrorAction SilentlyContinue -InformationAction SilentlyContinue 6>$null
        }
    } catch {}
    try {
        if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
            $null = Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue -InformationAction SilentlyContinue 6>$null
        }
    } catch {}
    try {
        if (Get-Command Disconnect-IPPSSession -ErrorAction SilentlyContinue) {
            $null = Disconnect-IPPSSession -Confirm:$false -ErrorAction SilentlyContinue -InformationAction SilentlyContinue 6>$null
        }
    } catch {}
    try {
        if (Get-Command Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue) {
            $null = Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue -InformationAction SilentlyContinue 6>$null
        }
    } catch {}
    try {
        if (Get-Command Disconnect-PnPOnline -ErrorAction SilentlyContinue) {
            $null = Disconnect-PnPOnline -ErrorAction SilentlyContinue -InformationAction SilentlyContinue 6>$null
        }
    } catch {}
    try {
        if (Get-Command Disconnect-SPOService -ErrorAction SilentlyContinue) {
            $null = Disconnect-SPOService -ErrorAction SilentlyContinue -InformationAction SilentlyContinue 6>$null
        }
    } catch {}
}

# ============================================================================
# COLLECTORS — each returns an [ordered] hashtable; no side effects.
# ============================================================================

function Get-TenantMetadata {
    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('Directory.Read.All', 'Organization.Read.All'))) {
        return [ordered]@{
            Status               = 'not collected'
            Reason               = 'Missing Graph permission: Directory.Read.All and/or Organization.Read.All'
            TenantName           = 'not collected'
            TenantId             = 'not collected'
            DefaultDomain        = 'not collected'
            GeoLocation          = 'not collected'
            ScriptVersion        = '1.0'
            RunTimestampUtc      = (Get-Date).ToUniversalTime().ToString('o')
            ExecutingAdmin       = if ($script:MgContext) { $script:MgContext.Account } else { 'not collected' }
            GrantedScopes        = if ($script:MgContext -and $script:MgContext.Scopes) { @($script:MgContext.Scopes) } elseif (@($script:GraphProvidedScopes).Count -gt 0) { @($script:GraphProvidedScopes) } else { 'not collected' }
            ConsentStatus        = if (@($script:GraphMissingScopes).Count -gt 0) { 'partial' } else { 'complete' }
            MissingGraphPermissions = @($script:GraphMissingScopes)
            TotalUsers           = 'not collected'
            MemberUsers          = 'not collected'
            GuestUsers           = 'not collected'
            IncludeSampling      = [bool]$IncludeSampling
            IncludeDetailedData  = [bool]$IncludeDetailedData
            SampleSize           = [int]$SampleSize
        }
    }

    $org = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/organization')) | Select-Object -First 1
    $users = @(Get-DirectoryUsersForReadiness)
    $userCount = @($users).Count
    $memberCount = @($users | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'userType' -Default '')) -ne 'Guest' }).Count
    $guestCount = @($users | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'userType' -Default '')) -eq 'Guest' }).Count
    $defaultDomain = @($org.VerifiedDomains | Where-Object { $_.IsDefault -eq $true } | Select-Object -ExpandProperty Name -First 1)

    return [ordered]@{
        Status             = 'reported'
        Reason             = ''
        TenantName         = $org.DisplayName
        TenantId           = $org.Id
        DefaultDomain      = $defaultDomain
        GeoLocation        = $org.CountryLetterCode
        ScriptVersion      = '1.0'
        RunTimestampUtc    = (Get-Date).ToUniversalTime().ToString('o')
        ExecutingAdmin     = if ($script:MgContext) { $script:MgContext.Account } else { 'not collected' }
        GrantedScopes      = if ($script:MgContext -and $script:MgContext.Scopes) { @($script:MgContext.Scopes) } elseif (@($script:GraphProvidedScopes).Count -gt 0) { @($script:GraphProvidedScopes) } else { 'not collected' }
        ConsentStatus      = if (@($script:GraphMissingScopes).Count -gt 0) { 'partial' } else { 'complete' }
        MissingGraphPermissions = @($script:GraphMissingScopes)
        AuthenticationMode = 'interactive-device-code'
        UnattendedExecutionSupported = $false
        AuthenticationModeNote = 'This script currently supports interactive delegated Graph authentication only; app-only or certificate-based unattended execution is not implemented.'
        TotalUsers         = $userCount
        MemberUsers        = $memberCount
        GuestUsers         = $guestCount
        IncludeSampling    = [bool]$IncludeSampling
        IncludeDetailedData = [bool]$IncludeDetailedData
        SampleSize         = [int]$SampleSize
    }
}

function Get-LicensingReadiness {
    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('Directory.Read.All', 'Organization.Read.All'))) {
        return [ordered]@{
            Status = 'not collected'
            Reason = 'Missing Graph permission: Directory.Read.All and/or Organization.Read.All'
            Skus = @()
            CopilotSku = 'not collected'
            QualifyingBaseLicenses = @()
            PrereqReadyUserCount = 'not collected'
            PrereqReadyPercent = 'not collected'
        }
    }

    $skus = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus'))
    $skuList = @($skus | ForEach-Object {
        $enabledUnits = 0
        if ($_.PSObject.Properties.Name -contains 'EnabledUnits') {
            $enabledUnits = [int]$_.EnabledUnits
        }
        elseif ($_.PSObject.Properties.Name -contains 'PrepaidUnits' -and $null -ne $_.PrepaidUnits -and $_.PrepaidUnits.PSObject.Properties.Name -contains 'Enabled') {
            $enabledUnits = [int]$_.PrepaidUnits.Enabled
        }

        $consumedUnits = 0
        if ($_.PSObject.Properties.Name -contains 'ConsumedUnits' -and $null -ne $_.ConsumedUnits) {
            $consumedUnits = [int]$_.ConsumedUnits
        }

        [ordered]@{
            SkuPartNumber = if ([string]::IsNullOrWhiteSpace([string]$_.SkuPartNumber)) { [string]$_.SkuId } else { [string]$_.SkuPartNumber }
            EnabledUnits  = $enabledUnits
            ConsumedUnits = $consumedUnits
        }
    })
    $knownCopilotSkus = @(
        'MICROSOFT_365_COPILOT',
        'M365_COPILOT',
        'COPILOT_FOR_MICROSOFT_365',
        'COPILOT_STUDIO',
        'COPILOT_STUDIO_ATTACH',
        'M365COPILOT',
        'Microsoft_365_Copilot'
    )
    $knownCopilotSkusUpper = @($knownCopilotSkus | ForEach-Object { $_.ToUpperInvariant() })

    $copilotSkuMatches = @($skus |
        Where-Object {
            $_.PSObject.Properties.Name -contains 'SkuPartNumber' -and
            -not [string]::IsNullOrWhiteSpace([string]$_.SkuPartNumber) -and
            (
                ([string]$_.SkuPartNumber).ToUpperInvariant() -in $knownCopilotSkusUpper -or
                ([string]$_.SkuPartNumber) -match 'COPILOT'
            )
        } |
        Select-Object -ExpandProperty SkuPartNumber -Unique)

    $qualifying = @($skus | Where-Object { $_.SkuPartNumber -match 'ENTERPRISEPACK|STANDARDPACK|M365BUSINESSSTANDARD|M365BUSINESSPREMIUM|E3|E5' } | Select-Object -ExpandProperty SkuPartNumber)

    $users = @(Get-DirectoryUsersForReadiness)
    $ready = 0
    foreach ($u in $users) {
        $plans = @(Get-ObjectPropertyValue -InputObject $u -Name 'assignedPlans' -Default @())
        # Graph assignedPlans expose capabilityStatus + service (not provisioningStatus/servicePlanName).
        $enabledServices = @($plans |
            ForEach-Object {
                $capability = [string](Get-ObjectPropertyValue -InputObject $_ -Name 'capabilityStatus' -Default '')
                if ($capability -eq 'Enabled') {
                    [string](Get-ObjectPropertyValue -InputObject $_ -Name 'service' -Default '')
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

        $hasApps = @($enabledServices | Where-Object { $_ -match '^(MicrosoftOffice|OfficeForWeb)$' }).Count -gt 0
        $hasExchange = @($enabledServices | Where-Object { $_ -match '^exchange$' }).Count -gt 0
        $hasSharePoint = @($enabledServices | Where-Object { $_ -match '^SharePoint' }).Count -gt 0
        $hasTeams = @($enabledServices | Where-Object { $_ -match '^(TeamspaceAPI|MicrosoftTeams)' }).Count -gt 0

        if ($hasApps -and $hasExchange -and $hasSharePoint -and $hasTeams) {
            $ready++
        }
    }

    return [ordered]@{
        Status                = 'reported'
        Reason                = ''
        Skus                  = $skuList
        CopilotSku            = if (@($copilotSkuMatches).Count -gt 0) { @($copilotSkuMatches) } else { 'not collected' }
        QualifyingBaseLicenses = @($qualifying)
        PrereqReadyUserCount  = $ready
        PrereqReadyPercent   = if (@($users).Count -gt 0) { [math]::Round(($ready / @($users).Count) * 100, 2) } else { 0 }
    }
}

function Get-IdentityReadiness {
    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('AuditLog.Read.All', 'Directory.Read.All'))) {
        return [ordered]@{
            Status               = 'not collected'
            Reason               = 'Missing Graph permission: AuditLog.Read.All and/or Directory.Read.All'
            MfaCapableCount      = 'not collected'
            MfaRegisteredCount   = 'not collected'
            PasswordlessCount    = 'not collected'
            MfaPopulationUserCount = 'not collected'
            MfaPopulationSource  = 'not collected'
            MfaSummaryTotalUserCount = 'not collected'
            MfaSummaryUserTypes  = 'not collected'
            MfaSummaryUserRoles  = 'not collected'
            MfaCapableSummaryCount = 'not collected'
            GuestUsers           = 'not collected'
            StaleUsers           = 'not collected'
            GlobalAdminCount     = 'not collected'
            PrivilegedRoleCount  = 'not collected'
            SecurityDefaultsEnabled = 'not collected'
        }
    }

    $auth = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails'))
    $featureSummary = Get-UserRegistrationFeatureSummary -IncludedUserTypes 'all' -IncludedUserRoles 'all'
    $mfaCapable = @($auth | Where-Object { $_.IsMfaCapable -eq $true }).Count
    $mfaRegistered = @($auth | Where-Object { $_.IsMfaRegistered -eq $true }).Count
    $passwordless = @($auth | Where-Object { $_.IsPasswordlessCapable -eq $true }).Count

    $mfaPopulationUserCount = 'not collected'
    $mfaPopulationSource = 'not collected'
    if (@($auth).Count -gt 0) {
        $mfaPopulationUserCount = @($auth).Count
        $mfaPopulationSource = 'reports/authenticationMethods/userRegistrationDetails row count'
    }
    elseif ($featureSummary -and $featureSummary.PSObject.Properties.Name -contains 'totalUserCount') {
        $mfaPopulationUserCount = Get-IntValue -Value $featureSummary.totalUserCount -Default 0
        $mfaPopulationSource = 'reports/authenticationMethods/usersRegisteredByFeature totalUserCount fallback'
    }

    $mfaCapableSummaryCount = $null
    if ($featureSummary -and $featureSummary.PSObject.Properties.Name -contains 'userRegistrationFeatureCounts') {
        $mfaCapableSummaryCount = Get-UserRegistrationFeatureCount -FeatureCounts $featureSummary.userRegistrationFeatureCounts -FeatureName 'mfaCapable'
    }

    $roles = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/directoryRoles'))
    $targetRoleNames = @(
        'Global Administrator','Exchange Administrator','SharePoint Administrator',
        'Teams Administrator','Security Administrator','Compliance Administrator',
        'Conditional Access Administrator','Privileged Role Administrator'
    )
    $globalAdminCount = 0
    $privilegedRoleCount = 0

    foreach ($role in ($roles | Where-Object { $_.DisplayName -in $targetRoleNames })) {
        $members = @((Get-GraphRestAll -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($role.Id)/members" -Query @{ '$select' = 'id' }))
        $memberCount = @($members).Count

        if ($role.DisplayName -match 'Global Administrator') {
            $globalAdminCount = $memberCount
        }
        elseif ($role.DisplayName -match 'Exchange|SharePoint|Teams|Security|Compliance|Conditional Access|Privileged') {
            $privilegedRoleCount += $memberCount
        }
    }

    $users = @(Get-DirectoryUsersForReadiness -IncludeSignInActivity)
    $guestUsers = @($users | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'userType' -Default '')) -eq 'Guest' }).Count
    $staleUsers = 0
    foreach ($u in $users) {
        $last = $null
        if ($u.PSObject.Properties.Name -contains 'SignInActivity' -and $null -ne $u.SignInActivity) {
            if ($u.SignInActivity.PSObject.Properties.Name -contains 'LastSignInDateTime') {
                $last = $u.SignInActivity.LastSignInDateTime
            }
        }

        if ($last) {
            if ((Get-Date).ToUniversalTime() - [DateTime]$last -gt [TimeSpan]::FromDays(90)) { $staleUsers++ }
        }
    }

    $securityDefaults = Get-GraphRest -Uri 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy'
    $identityStatus = if (Test-GraphPermissionsAvailable -RequiredPermissions @('Policy.Read.All')) { 'reported' } else { 'partial' }
    $identityReason = if ($identityStatus -eq 'partial') { 'Security defaults could not be collected because Policy.Read.All is missing.' } else { '' }

    return [ordered]@{
        Status               = $identityStatus
        Reason               = $identityReason
        MfaCapableCount      = $mfaCapable
        MfaRegisteredCount   = $mfaRegistered
        PasswordlessCount    = $passwordless
        MfaPopulationUserCount = $mfaPopulationUserCount
        MfaPopulationSource  = $mfaPopulationSource
        MfaSummaryTotalUserCount = if ($featureSummary -and $featureSummary.PSObject.Properties.Name -contains 'totalUserCount') { Get-IntValue -Value $featureSummary.totalUserCount -Default 0 } else { 'not collected' }
        MfaSummaryUserTypes  = if ($featureSummary -and $featureSummary.PSObject.Properties.Name -contains 'userTypes') { [string]$featureSummary.userTypes } else { 'not collected' }
        MfaSummaryUserRoles  = if ($featureSummary -and $featureSummary.PSObject.Properties.Name -contains 'userRoles') { [string]$featureSummary.userRoles } else { 'not collected' }
        MfaCapableSummaryCount = if ($null -ne $mfaCapableSummaryCount) { $mfaCapableSummaryCount } else { 'not collected' }
        GuestUsers           = $guestUsers
        StaleUsers           = $staleUsers
        GlobalAdminCount     = $globalAdminCount
        PrivilegedRoleCount  = $privilegedRoleCount
        AdminRoleEnumerationScope = 'activated directory roles with active memberships only'
        AdminRoleEnumerationNote = 'Counts exclude PIM-eligible assignments and roles not currently activated in the tenant.'
        SecurityDefaultsEnabled = if ($securityDefaults) { [bool]$securityDefaults.IsEnabled } else { 'not collected' }
    }
}

function Get-ConditionalAccessReadiness {
    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('Policy.Read.All'))) {
        return [ordered]@{
            Status = 'not collected'
            Reason = 'Missing Graph permission: Policy.Read.All'
            TotalPolicies = 'not collected'
            EnabledPolicies = 'not collected'
            CoverageFlags = [ordered]@{
                RequireMfa = 'not collected'
                BlockLegacyAuth = 'not collected'
                DeviceCompliance = 'not collected'
                SignInRisk = 'not collected'
                HasPoliciesWithExclusions = 'not collected'
            }
        }
    }

    $policies = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies'))
    $enabledPolicies = @($policies | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'state' -Default '')) -eq 'enabled' })
    $enabled = @($enabledPolicies).Count

    $requireMfa = @($enabledPolicies | Where-Object {
        $grantControls = Get-ObjectPropertyValue -InputObject $_ -Name 'grantControls'
        @(Get-ObjectPropertyValue -InputObject $grantControls -Name 'builtInControls' -Default @()) -contains 'mfa'
    }).Count -gt 0

    $blockLegacyAuth = @($enabledPolicies | Where-Object {
        $grantControls = Get-ObjectPropertyValue -InputObject $_ -Name 'grantControls'
        $conditions = Get-ObjectPropertyValue -InputObject $_ -Name 'conditions'
        $clientAppTypes = @(Get-ObjectPropertyValue -InputObject $conditions -Name 'clientAppTypes' -Default @())
        # Some newer CA policies express this through authenticationFlows instead of clientAppTypes.
        $authenticationFlows = Get-ObjectPropertyValue -InputObject $conditions -Name 'authenticationFlows'
        $transferMethods = [string](Get-ObjectPropertyValue -InputObject $authenticationFlows -Name 'transferMethods' -Default '')

        (@(Get-ObjectPropertyValue -InputObject $grantControls -Name 'builtInControls' -Default @()) -contains 'block') -and
        (
            ($clientAppTypes -contains 'exchangeActiveSync' -or $clientAppTypes -contains 'other') -or
            (-not [string]::IsNullOrWhiteSpace($transferMethods))
        )
    }).Count -gt 0

    $deviceCompliance = @($enabledPolicies | Where-Object {
        $grantControls = Get-ObjectPropertyValue -InputObject $_ -Name 'grantControls'
        @(Get-ObjectPropertyValue -InputObject $grantControls -Name 'builtInControls' -Default @()) -contains 'compliantDevice'
    }).Count -gt 0

    $signInRisk = @($enabledPolicies | Where-Object {
        $conditions = Get-ObjectPropertyValue -InputObject $_ -Name 'conditions'
        @(Get-ObjectPropertyValue -InputObject $conditions -Name 'signInRiskLevels' -Default @()).Count -gt 0
    }).Count -gt 0

    $hasPoliciesWithExclusions = @($enabledPolicies | Where-Object {
        $conditions = Get-ObjectPropertyValue -InputObject $_ -Name 'conditions'
        $userConditions = Get-ObjectPropertyValue -InputObject $conditions -Name 'users'
        $appConditions = Get-ObjectPropertyValue -InputObject $conditions -Name 'applications'
        $excludedCount = @(Get-ObjectPropertyValue -InputObject $userConditions -Name 'excludeUsers' -Default @()).Count +
            @(Get-ObjectPropertyValue -InputObject $userConditions -Name 'excludeGroups' -Default @()).Count +
            @(Get-ObjectPropertyValue -InputObject $userConditions -Name 'excludeRoles' -Default @()).Count +
            @(Get-ObjectPropertyValue -InputObject $appConditions -Name 'excludeApplications' -Default @()).Count
        $excludedCount -gt 0
    }).Count -gt 0

    return [ordered]@{
        Status             = 'reported'
        Reason             = ''
        TotalPolicies      = @($policies).Count
        EnabledPolicies    = $enabled
        CoverageFlags      = [ordered]@{
            RequireMfa         = $requireMfa
            BlockLegacyAuth    = $blockLegacyAuth
            DeviceCompliance   = $deviceCompliance
            SignInRisk         = $signInRisk
            HasPoliciesWithExclusions = $hasPoliciesWithExclusions
        }
        CoverageNotes      = [ordered]@{
            BlockLegacyAuth = 'Heuristic signal based on enabled block policies targeting legacy-style client app types or related authentication flow hints; not an authoritative legacy-auth protection determination.'
        }
    }
}

function Get-DataGovernanceReadiness {
    $hasLabelPermission = Test-GraphPermissionsAvailable -RequiredPermissions @('InformationProtectionPolicy.Read.All')
    $labels = if ($hasLabelPermission) { @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/informationProtection/policy/labels')) } else { @() }
    # This endpoint exposes isActive rather than isPublished; fall back accordingly.
    $published = @($labels | Where-Object {
        $isPublished = Get-ObjectPropertyValue -InputObject $_ -Name 'isPublished'
        $isActive = Get-ObjectPropertyValue -InputObject $_ -Name 'isActive'
        ($isPublished -eq $true) -or ($null -eq $isPublished -and $isActive -eq $true)
    }).Count
    $autoLabel = @($labels | Where-Object {
        (Get-ObjectPropertyValue -InputObject $_ -Name 'isLabelAutoClassified') -eq $true -or
        (Get-ObjectPropertyValue -InputObject $_ -Name 'isDefault') -eq $true
    }).Count

    $hasIppsSession = $false
    if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        $connections = @((Get-ConnectionInformation -ErrorAction SilentlyContinue))
        $hasIppsSession = @($connections | Where-Object {
            ($_.PSObject.Properties.Name -contains 'Name' -and [string]$_.Name -match 'IPPSSession|SecurityCompliance') -or
            ($_.PSObject.Properties.Name -contains 'ConnectionUri' -and [string]$_.ConnectionUri -match 'ps\.compliance\.protection\.outlook\.com')
        }).Count -gt 0
    }

    $dlp = if ($hasIppsSession -and (Get-Command Get-DlpCompliancePolicy -ErrorAction SilentlyContinue)) { Get-DlpCompliancePolicy -ErrorAction SilentlyContinue | Select-Object Name, Mode, State } else { 'not collected' }
    $retention = if ($hasIppsSession -and (Get-Command Get-RetentionCompliancePolicy -ErrorAction SilentlyContinue)) { Get-RetentionCompliancePolicy -ErrorAction SilentlyContinue | Select-Object Name, Enabled } else { 'not collected' }
    $hasComplianceData = (Test-ValueCollected -Value $dlp) -or (Test-ValueCollected -Value $retention)

    $status = 'reported'
    $reason = ''
    if (-not $hasLabelPermission -and -not $hasComplianceData) {
        $status = 'not collected'
        if (-not $hasIppsSession) {
            $reason = 'Neither labels nor compliance policy data could be collected because no Purview compliance session was available.'
        }
        else {
            $reason = 'Neither labels nor compliance policy data could be collected.'
        }
    }
    elseif (-not $hasLabelPermission -or -not $hasComplianceData) {
        $status = 'partial'
        if (-not $hasIppsSession) {
            $reason = 'Only part of the data governance signals could be collected because no Purview compliance session was available.'
        }
        else {
            $reason = 'Only part of the data governance signals could be collected.'
        }
    }

    return [ordered]@{
        Status                    = $status
        Reason                    = $reason
        SensitivityLabelsPublished = if ($hasLabelPermission) { $published } else { 'not collected' }
        SensitivityLabelsTotal     = if ($hasLabelPermission) { @($labels).Count } else { 'not collected' }
        AutoLabelingIndicators     = if ($hasLabelPermission) { $autoLabel } else { 'not collected' }
        DlpPolicies                = $dlp
        RetentionPolicies          = $retention
    }
}

function Get-SharePointOneDriveReadiness {
    [CmdletBinding()]
    param([switch]$Sample, [int]$Max = 200)

    if ($script:SharePointChildCollectionSucceeded) {
        return Get-SharePointOneDriveReadinessLocal -Sample:$Sample -Max $Max
    }

    $adminUrl = Get-SharePointAdminUrl
    if ([string]::IsNullOrWhiteSpace($adminUrl)) {
        return [ordered]@{
            Status                = 'not collected'
            Reason                = 'Unable to determine the SharePoint admin URL for the child process.'
            ExternalSharingLevel  = 'not collected'
            DefaultSharingLinkType = 'not collected'
            RestrictedSearchEnabled = 'not collected'
            TotalSiteCount        = 'not collected'
            Truncated             = 'not collected'
            Sites                 = @()
            OneDriveCoveragePercent = 'not collected'
            InactiveSiteCount     = 'not collected'
            OversharingSignals    = [ordered]@{
                Status                    = 'not collected'
                CollectionMethod          = 'not collected'
                CollectionScope           = 'not collected'
                CollectionPathUsed        = 'not collected'
                CollectionReason          = 'not collected'
                RequiredForReliableCollection = 'not collected'
                DocumentationNote         = 'not collected'
                SampledAnyoneLinkCount    = 'not collected'
                SampledOrgWideLinkCount   = 'not collected'
                SitesWithAnyoneLinksCount = 'not collected'
            }
        }
    }

    $pwshPath = (Get-Command pwsh -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1)
    if ([string]::IsNullOrWhiteSpace($pwshPath) -or -not (Test-Path -LiteralPath $pwshPath)) {
        return [ordered]@{
            Status                = 'not collected'
            Reason                = "PowerShell 7 executable 'pwsh' not found for the SharePoint child process."
            ExternalSharingLevel  = 'not collected'
            DefaultSharingLinkType = 'not collected'
            RestrictedSearchEnabled = 'not collected'
            TotalSiteCount        = 'not collected'
            Truncated             = 'not collected'
            Sites                 = @()
            OneDriveCoveragePercent = 'not collected'
            InactiveSiteCount     = 'not collected'
            OversharingSignals    = [ordered]@{
                Status                    = 'not collected'
                CollectionMethod          = 'not collected'
                CollectionScope           = 'not collected'
                CollectionPathUsed        = 'not collected'
                CollectionReason          = 'not collected'
                RequiredForReliableCollection = 'not collected'
                DocumentationNote         = 'not collected'
                SampledAnyoneLinkCount    = 'not collected'
                SampledOrgWideLinkCount   = 'not collected'
                SitesWithAnyoneLinksCount = 'not collected'
            }
        }
    }

    $childArgs = @(
        '-NoProfile',
        '-NoLogo',
        '-ExecutionPolicy', 'Bypass',
        '-File', $PSCommandPath,
        '-SharePointCollectionChild',
        '-SharePointAdminUrl', $adminUrl
    )
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $childArgs += @('-TenantId', $TenantId) }
    if (-not [string]::IsNullOrWhiteSpace($PnPClientId)) { $childArgs += @('-PnPClientId', $PnPClientId) }
    if (-not [string]::IsNullOrWhiteSpace($GraphClientId)) { $childArgs += @('-GraphClientId', $GraphClientId) }
    if ($Sample) { $childArgs += '-IncludeSampling' }
    if ($Max -ne 200) { $childArgs += @('-SampleSize', [string]$Max) }

    $childOutput = & $pwshPath @childArgs 2>&1 | Out-String
    if ([string]::IsNullOrWhiteSpace($childOutput)) {
        return [ordered]@{
            Status                  = 'not collected'
            Reason                  = 'SharePoint child process returned no output.'
            ExternalSharingLevel    = 'not collected'
            DefaultSharingLinkType  = 'not collected'
            RestrictedSearchEnabled = 'not collected'
            TotalSiteCount          = 'not collected'
            Truncated               = 'not collected'
            Sites                   = @()
            OneDriveCoveragePercent = 'not collected'
            InactiveSiteCount       = 'not collected'
            OversharingSignals      = [ordered]@{
                Status                        = 'not collected'
                CollectionMethod              = 'not collected'
                CollectionScope               = 'not collected'
                CollectionPathUsed            = 'not collected'
                CollectionReason              = 'not collected'
                RequiredForReliableCollection = 'not collected'
                DocumentationNote             = 'not collected'
                SampledAnyoneLinkCount        = 'not collected'
                SampledOrgWideLinkCount       = 'not collected'
                SitesWithAnyoneLinksCount     = 'not collected'
            }
        }
    }

    try {
        $jsonStart = $childOutput.IndexOf('{')
        $jsonEnd = $childOutput.LastIndexOf('}')
        if ($jsonStart -lt 0 -or $jsonEnd -lt $jsonStart) {
            throw 'No JSON payload was found in the SharePoint child output.'
        }

        $jsonPayload = $childOutput.Substring($jsonStart, $jsonEnd - $jsonStart + 1)
        $result = $jsonPayload | ConvertFrom-Json -ErrorAction Stop
        $script:SharePointChildCollectionSucceeded = ([string]$result.Status -ne 'not collected')
        return $result
    }
    catch {
        return [ordered]@{
            Status                  = 'not collected'
            Reason                  = "SharePoint child process did not return valid JSON output: $($_.Exception.Message). Raw output: $childOutput"
            ExternalSharingLevel    = 'not collected'
            DefaultSharingLinkType  = 'not collected'
            RestrictedSearchEnabled = 'not collected'
            TotalSiteCount          = 'not collected'
            Truncated               = 'not collected'
            Sites                   = @()
            OneDriveCoveragePercent = 'not collected'
            InactiveSiteCount       = 'not collected'
            OversharingSignals      = [ordered]@{
                Status                        = 'not collected'
                CollectionMethod              = 'not collected'
                CollectionScope               = 'not collected'
                CollectionPathUsed            = 'not collected'
                CollectionReason              = 'not collected'
                RequiredForReliableCollection = 'not collected'
                DocumentationNote             = 'not collected'
                SampledAnyoneLinkCount        = 'not collected'
                SampledOrgWideLinkCount       = 'not collected'
                SitesWithAnyoneLinksCount     = 'not collected'
            }
        }
    }
}

function Get-SharePointOneDriveReadinessLocal {
    param([switch]$Sample, [int]$Max = 200)

    $tenant = $null
    $sites = @()

    try {
        if (Get-Command Get-PnPTenant -ErrorAction SilentlyContinue) { $tenant = Get-PnPTenant -ErrorAction Stop }
    } catch {}
    try {
        if (Get-Command Get-PnPTenantSite -ErrorAction SilentlyContinue) { $sites = @(Get-PnPTenantSite -IncludeOneDriveSites -ErrorAction Stop) }
    } catch {}

    if ($null -eq $tenant -and @($sites).Count -eq 0) {
        return [ordered]@{
            Status                = 'not collected'
            Reason                = 'SharePoint / PnP tenant session was unavailable or returned no data.'
            ExternalSharingLevel  = 'not collected'
            DefaultSharingLinkType = 'not collected'
            RestrictedSearchEnabled = 'not collected'
            TotalSiteCount        = 'not collected'
            Truncated             = 'not collected'
            Sites                 = @()
            OneDriveCoveragePercent = 'not collected'
            InactiveSiteCount     = 'not collected'
            OversharingSignals    = [ordered]@{
                Status                    = 'not collected'
                CollectionMethod          = 'not collected'
                CollectionScope           = 'not collected'
                CollectionPathUsed        = 'not collected'
                CollectionReason          = 'SharePoint / PnP tenant session was unavailable or returned no data.'
                RequiredForReliableCollection = 'not collected'
                DocumentationNote         = 'not collected'
                SampledAnyoneLinkCount    = 'not collected'
                SampledOrgWideLinkCount   = 'not collected'
                SitesWithAnyoneLinksCount = 'not collected'
            }
        }
    }

    $totalCount = @($sites).Count
    $siteList = @($sites)

    if ($Sample -and $totalCount -gt $Max) {
        $siteList = @($siteList | Select-Object -First $Max)
    }

    $oneDriveCoverage = 0
    $inactive = 0
    $anyoneLinkCount = 0
    $orgWideLinkCount = 0
    $sitesWithAnyoneLinks = 0
    $oversharingSignalAvailable = $false
    $oversharingPropertySource = 'none'
    $siteSummaries = foreach ($site in $siteList) {
        if ($site.LastContentModifiedDate -and ((Get-Date).ToUniversalTime() - [DateTime]$site.LastContentModifiedDate) -gt [TimeSpan]::FromDays(180)) { $inactive++ }
        if ($site.Url -match '/personal/') { $oneDriveCoverage++ }

        $siteAnyoneLinks = Get-FirstIntPropertyValue -InputObject $site -CandidateNames @(
            'AnyoneLinkCount',
            'AnyoneLinksCount',
            'AnonymousLinkCount',
            'AnonymousLinksCount'
        )
        if ($null -ne $siteAnyoneLinks) {
            $oversharingSignalAvailable = $true
            $oversharingPropertySource = 'PnP tenant site metadata counters (non-standard exposure)'
            $anyoneLinkCount += $siteAnyoneLinks
            if ($siteAnyoneLinks -gt 0) { $sitesWithAnyoneLinks++ }
        }

        $siteOrgWideLinks = Get-FirstIntPropertyValue -InputObject $site -CandidateNames @(
            'OrganizationLinkCount',
            'OrganizationLinksCount',
            'CompanyLinkCount',
            'CompanyLinksCount',
            'PeopleInOrganizationLinkCount',
            'PeopleInOrganizationLinksCount',
            'EveryoneExceptExternalUsersLinkCount'
        )
        if ($null -ne $siteOrgWideLinks) {
            $oversharingSignalAvailable = $true
            $oversharingPropertySource = 'PnP tenant site metadata counters (non-standard exposure)'
            $orgWideLinkCount += $siteOrgWideLinks
        }

        [ordered]@{
            Url                  = $site.Url
            StorageQuotaInMB    = $site.StorageQuota
            UsedStorageInMB     = $site.StorageUsageCurrent
            SharingCapability   = $site.SharingCapability
            AnyoneLinkEnabled   = $site.SharingCapability -in @('AnonymousAccess', 'ExternalUserAndGuestSharing')
            LastContentModified = $site.LastContentModifiedDate
        }
    }

    return [ordered]@{
        Status                = 'reported'
        Reason                = ''
        ExternalSharingLevel   = if ($tenant) { [string](Get-ObjectPropertyValue -InputObject $tenant -Name 'SharingCapability' -Default 'not collected') } else { 'not collected' }
        DefaultSharingLinkType = if ($tenant) { [string](Get-ObjectPropertyValue -InputObject $tenant -Name 'DefaultSharingLinkType' -Default 'not collected') } else { 'not collected' }
        RestrictedSearchEnabled = if ($tenant -and $null -ne (Get-ObjectPropertyValue -InputObject $tenant -Name 'RestrictedSharePointSearch')) { [bool](Get-ObjectPropertyValue -InputObject $tenant -Name 'RestrictedSharePointSearch') } else { 'not collected' }
        TotalSiteCount         = $totalCount
        Truncated              = [bool]($Sample -and $totalCount -gt $Max)
        Sites                  = @($siteSummaries)
        OneDrivePersonalSiteCount = $oneDriveCoverage
        OneDrivePersonalSiteSharePercent = if (@($siteList).Count -gt 0) { [math]::Round(($oneDriveCoverage / @($siteList).Count) * 100, 2) } else { 0 }
        OneDriveCoverageMetric = 'Deprecated alias: OneDriveCoveragePercent reflects the share of sampled SharePoint/PnP tenant sites whose URL is a OneDrive personal site, not licensed-user OneDrive provisioning coverage.'
        OneDriveCoveragePercent = if (@($siteList).Count -gt 0) { [math]::Round(($oneDriveCoverage / @($siteList).Count) * 100, 2) } else { 0 }
        InactiveSiteCount      = $inactive
        OversharingSignals     = [ordered]@{
            Status                     = if ($oversharingSignalAvailable) { 'reported' } else { 'not collected' }
            CollectionMethod           = if ($oversharingSignalAvailable) {
                'Best-effort from PnP tenant site metadata counters exposed by current cmdlet/runtime.'
            }
            else {
                'No oversharing counters collected from standard Get-PnPTenantSite output.'
            }
            CollectionScope            = if ($oversharingSignalAvailable) { 'sampled site metadata counters only (not guaranteed tenant-wide)' } else { 'standard PnP tenant site metadata only' }
            CollectionPathUsed         = if ($oversharingSignalAvailable) { $oversharingPropertySource } else { 'none' }
            CollectionReason           = if ($oversharingSignalAvailable) {
                'reported from tenant site metadata counters exposed by current cmdlet/runtime'
            }
            else {
                'standard Get-PnPTenantSite output does not normally include per-site sharing-link counters'
            }
            RequiredForReliableCollection = if ($oversharingSignalAvailable) {
                'SharePoint Advanced Management oversharing telemetry is still recommended for reliable tenant-wide results'
            }
            else {
                'SharePoint Advanced Management oversharing telemetry (recommended), or costly per-site sharing-link enumeration'
            }
            DocumentationNote          = if ($oversharingSignalAvailable) {
                'Counters were exposed in this run, but this is not guaranteed in standard PnP output across tenants.'
            }
            else {
                'This run did not collect oversharing link counters; absence here does not imply low oversharing risk.'
            }
            SampledAnyoneLinkCount     = if ($oversharingSignalAvailable) { $anyoneLinkCount } else { 'not collected' }
            SampledOrgWideLinkCount    = if ($oversharingSignalAvailable) { $orgWideLinkCount } else { 'not collected' }
            SitesWithAnyoneLinksCount  = if ($oversharingSignalAvailable) { $sitesWithAnyoneLinks } else { 'not collected' }
        }
    }
}

function Get-ExchangeReadiness {
    try {
        $mailboxes = @()
        $hasExoSession = $false

        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $hasExoSession = @((Get-ConnectionInformation -ErrorAction SilentlyContinue)).Count -gt 0
        }

        if ($hasExoSession -and (Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue)) {
            # ArchiveStatus and LitigationHoldEnabled are not in the default EXO minimum property set.
            $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails, ArchiveStatus, LitigationHoldEnabled -ErrorAction SilentlyContinue)
            $shared = @($mailboxes | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'RecipientTypeDetails' -Default '')) -eq 'SharedMailbox' }).Count
            $archive = @($mailboxes | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'ArchiveStatus' -Default '')) -eq 'Active' }).Count
            $litigation = @($mailboxes | Where-Object { (Get-ObjectPropertyValue -InputObject $_ -Name 'LitigationHoldEnabled') -eq $true }).Count

            return [ordered]@{
                Status               = 'reported'
                Reason               = ''
                TotalMailboxes       = @($mailboxes).Count
                UserMailboxes        = @($mailboxes | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'RecipientTypeDetails' -Default '')) -in @('UserMailbox', 'SharedMailbox') }).Count
                SharedMailboxes      = $shared
                ArchiveEnabledCount  = $archive
                LitigationHoldCount  = $litigation
                ExchangeOnlineReady  = @($mailboxes).Count -gt 0
            }
        }

        if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('Directory.Read.All'))) {
            return [ordered]@{
                Status               = 'not collected'
                Reason               = 'Exchange Online session was unavailable and Graph directory read access is missing.'
                TotalMailboxes       = 'not collected'
                UserMailboxes        = 'not collected'
                SharedMailboxes      = 'not collected'
                ArchiveEnabledCount  = 'not collected'
                LitigationHoldCount  = 'not collected'
                ExchangeOnlineReady  = 'not collected'
            }
        }

        $users = @(Get-DirectoryUsersForReadiness)
        # Graph omits null-valued properties, so access mail defensively under StrictMode.
        $mailboxCount = @($users | Where-Object { -not [string]::IsNullOrWhiteSpace([string](Get-ObjectPropertyValue -InputObject $_ -Name 'mail' -Default '')) }).Count

        return [ordered]@{
            Status               = 'partial'
            Reason               = 'Mailbox count was estimated from Graph users because Exchange Online session details were unavailable.'
            TotalMailboxes       = $mailboxCount
            UserMailboxes        = $mailboxCount
            SharedMailboxes      = 'not collected'
            ArchiveEnabledCount  = 'not collected'
            LitigationHoldCount  = 'not collected'
            ExchangeOnlineReady  = $mailboxCount -gt 0
        }
    }
    catch {
        return [ordered]@{
            Status               = 'not collected'
            Reason               = 'Exchange mailbox data could not be collected.'
            TotalMailboxes       = 'not collected'
            UserMailboxes        = 'not collected'
            SharedMailboxes      = 'not collected'
            ArchiveEnabledCount  = 'not collected'
            LitigationHoldCount  = 'not collected'
            ExchangeOnlineReady  = 'not collected'
        }
    }
}

function Get-TeamsReadiness {
    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('Group.Read.All'))) {
        return [ordered]@{
            Status                = 'not collected'
            Reason                = 'Missing Graph permission: Group.Read.All'
            TeamCount             = 'not collected'
            PrivateTeamCount      = 'not collected'
            PublicTeamCount       = 'not collected'
            StaleTeamCount        = 'not collected'
            MeetingPolicyState    = 'not collected'
            RecordingEnabled      = 'not collected'
            TranscriptionEnabled  = 'not collected'
        }
    }

    $teams = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/groups' -Query @{ '$filter' = "resourceProvisioningOptions/any(x:x eq 'Team')"; '$select' = 'id,displayName,visibility,renewedDateTime' }))
    $teamCount = @($teams).Count
    $privateTeams = @($teams | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'visibility' -Default '')) -eq 'Private' }).Count
    $publicTeams = @($teams | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'visibility' -Default '')) -eq 'Public' }).Count
    $staleTeams = @($teams | Where-Object {
        $renewed = Get-ObjectPropertyValue -InputObject $_ -Name 'renewedDateTime'
        $renewed -and ((Get-Date).ToUniversalTime() - [DateTime]$renewed) -gt [TimeSpan]::FromDays(180)
    }).Count

    $meetingPolicy = $null
    try {
        if ($script:TeamsConnected -and (Get-Command Get-CsTeamsMeetingPolicy -ErrorAction SilentlyContinue)) {
            $meetingPolicy = Get-CsTeamsMeetingPolicy -Identity Global -ErrorAction SilentlyContinue
            if (-not $meetingPolicy) {
                $meetingPolicy = @(Get-CsTeamsMeetingPolicy -ErrorAction SilentlyContinue) | Select-Object -First 1
            }
        }
    }
    catch {
        $meetingPolicy = $null
    }

    $hasAllowRecording = ($null -ne $meetingPolicy -and $meetingPolicy.PSObject.Properties.Name -contains 'AllowRecording')
    $hasAllowTranscription = ($null -ne $meetingPolicy -and $meetingPolicy.PSObject.Properties.Name -contains 'AllowTranscription')

    $recordingEnabled = if ($hasAllowRecording) { [bool]$meetingPolicy.AllowRecording } else { 'not collected' }
    $transcriptionEnabled = if ($hasAllowTranscription) { [bool]$meetingPolicy.AllowTranscription } else { 'not collected' }
    $meetingPolicyState = if ($hasAllowRecording -or $hasAllowTranscription) {
        $allowRecordingValue = if ($hasAllowRecording) { [bool]$meetingPolicy.AllowRecording } else { $false }
        $allowTranscriptionValue = if ($hasAllowTranscription) { [bool]$meetingPolicy.AllowTranscription } else { $false }
        if ($allowRecordingValue -or $allowTranscriptionValue) { 'enabled' } else { 'disabled' }
    }
    else {
        'not collected'
    }

    $teamsStatus = if ($hasAllowRecording -or $hasAllowTranscription) { 'reported' } else { 'partial' }
    $teamsReason = if ($teamsStatus -eq 'partial') { 'Teams meeting policy metadata was not collected.' } else { '' }

    return [ordered]@{
        Status                = $teamsStatus
        Reason                = $teamsReason
        TeamCount             = $teamCount
        PrivateTeamCount      = $privateTeams
        PublicTeamCount       = $publicTeams
        StaleTeamCount        = $staleTeams
        MeetingPolicyState    = $meetingPolicyState
        RecordingEnabled      = $recordingEnabled
        TranscriptionEnabled  = $transcriptionEnabled
    }
}

function Get-SearchIndexReadiness {
    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('ExternalConnection.Read.All'))) {
        return [ordered]@{
            Status = 'not collected'
            Reason = 'Missing Graph permission: ExternalConnection.Read.All'
            RestrictedSearchEnabled = 'not collected'
            ExternalConnectionsCount = 'not collected'
            ReadyExternalConnectionsCount = 'not collected'
            SearchSchemaCustomization = 'not collected'
            SemanticIndexIndicators = [ordered]@{
                ConnectorCount = 'not collected'
                ReadyConnectorCount = 'not collected'
                SearchState = 'not collected'
            }
        }
    }

    $connections = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/external/connections'))
    $readyConnectors = @($connections | Where-Object { $_.PSObject.Properties.Name -contains 'state' -and $_.state -match 'ready|active' }).Count

    return [ordered]@{
        Status = 'reported'
        Reason = ''
        RestrictedSearchEnabled = 'not collected'
        ExternalConnectionsCount = @($connections).Count
        ReadyExternalConnectionsCount = $readyConnectors
        SearchSchemaCustomization = 'not collected'
        SemanticIndexIndicators = [ordered]@{
            ConnectorCount = @($connections).Count
            ReadyConnectorCount = $readyConnectors
            SearchState = if (@($connections).Count -gt 0) { 'available' } else { 'not collected' }
        }
    }
}

function Get-AppsAndEndpointReadiness {
    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('Reports.Read.All'))) {
        return [ordered]@{
            Status = 'not collected'
            Reason = 'Missing Graph permission: Reports.Read.All'
            M365AppsUpdateChannel = 'not collected'
            VersionReadiness      = 'not collected'
            BrowserSignals        = 'not collected'
            EndpointSignals       = 'not collected'
            ReportCount           = 'not collected'
        }
    }

    $m365Apps = @()
    try {
        # getM365AppUserCounts is the supported report endpoint (returns CSV rows).
        $m365Apps = @(Get-GraphReportRows -Uri "https://graph.microsoft.com/v1.0/reports/getM365AppUserCounts(period='D30')")
    }
    catch {
        $m365Apps = @()
    }

    return [ordered]@{
        Status                = if (@($m365Apps).Count -gt 0) { 'reported' } else { 'not collected' }
        Reason                = if (@($m365Apps).Count -gt 0) { '' } else { 'M365 Apps report endpoint returned no data.' }
        M365AppsUpdateChannel = if (@($m365Apps).Count -gt 0) { 'reported' } else { 'not collected' }
        VersionReadiness      = if (@($m365Apps).Count -gt 0) { 'reported' } else { 'not collected' }
        BrowserSignals        = 'not collected'
        EndpointSignals       = 'not collected'
        ReportCount           = @($m365Apps).Count
    }
}

function Get-SecureScoreReadiness {
    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('SecurityEvents.Read.All'))) {
        return [ordered]@{
            Status = 'not collected'
            Reason = 'Missing Graph permission: SecurityEvents.Read.All'
            CurrentScore = 'not collected'
            MaxScore = 'not collected'
            ScorePercent = 'not collected'
            CurrentScoreDate = 'not collected'
            LicensedUserCount = 'not collected'
            EnabledServices = @()
            VendorInformation = 'not collected'
            ScoreHistory = @()
            ControlScoresSummary = [ordered]@{
                TotalControls = 'not collected'
                ControlsBelowMaxScore = 'not collected'
                TopControlGaps = @()
            }
            ControlProfilesSummary = [ordered]@{
                TotalProfiles = 'not collected'
                OpenProfiles = 'not collected'
                TopOpenProfiles = @()
            }
            ControlProfilesDetailedIncluded = [bool]$IncludeDetailedData
            ControlProfiles = @()
        }
    }

    $historyCount = if ($IncludeDetailedData) { 10 } else { 5 }
    $scores = @(
        Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/security/secureScores' -Query @{ '$top' = [string]$historyCount; '$orderby' = 'createdDateTime desc' }
    )

    if (@($scores).Count -eq 0) {
        return [ordered]@{
            Status = 'not collected'
            Reason = 'Secure Score endpoint returned no data.'
            CurrentScore = 'not collected'
            MaxScore = 'not collected'
            ScorePercent = 'not collected'
            CurrentScoreDate = 'not collected'
            LicensedUserCount = 'not collected'
            EnabledServices = @()
            VendorInformation = 'not collected'
            ScoreHistory = @()
            ControlScoresSummary = [ordered]@{
                TotalControls = 'not collected'
                ControlsBelowMaxScore = 'not collected'
                TopControlGaps = @()
            }
            ControlProfilesSummary = [ordered]@{
                TotalProfiles = 'not collected'
                OpenProfiles = 'not collected'
                TopOpenProfiles = @()
            }
            ControlProfilesDetailedIncluded = [bool]$IncludeDetailedData
            ControlProfiles = @()
        }
    }

    $sortedScores = @($scores | Sort-Object {
        $createdDateTime = Get-ObjectPropertyValue -InputObject $_ -Name 'createdDateTime'
        try { [datetime]$createdDateTime } catch { [datetime]::MinValue }
    } -Descending)
    $latestScore = $sortedScores | Select-Object -First 1

    $currentScore = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $latestScore -Name 'currentScore') -Default 0
    $maxScore = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $latestScore -Name 'maxScore') -Default 0
    $scorePercent = if ($maxScore -gt 0) { [math]::Round(($currentScore / $maxScore) * 100, 2) } else { 'not collected' }

    $scoreHistory = @($sortedScores | ForEach-Object {
        [ordered]@{
            CreatedDateTime = Get-ObjectPropertyValue -InputObject $_ -Name 'createdDateTime' -Default 'not collected'
            CurrentScore = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'currentScore') -Default 0
            MaxScore = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'maxScore') -Default 0
            ScorePercent = if ((Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'maxScore') -Default 0) -gt 0) {
                [math]::Round((
                    (Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'currentScore') -Default 0) /
                    (Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'maxScore') -Default 0)
                ) * 100, 2)
            }
            else {
                'not collected'
            }
        }
    })

    $controlScores = @(Get-ObjectPropertyValue -InputObject $latestScore -Name 'controlScores' -Default @())
    $controlGaps = @($controlScores | ForEach-Object {
        $controlScoreValue = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'score') -Default 0
        $controlMaxScoreValue = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'maxScore') -Default 0
        [ordered]@{
            ControlName = Get-ObjectPropertyValue -InputObject $_ -Name 'controlName' -Default 'not collected'
            Score = $controlScoreValue
            MaxScore = $controlMaxScoreValue
            Gap = [math]::Round(($controlMaxScoreValue - $controlScoreValue), 2)
        }
    } | Where-Object { $_.Gap -gt 0 } | Sort-Object Gap -Descending)

    $controlProfiles = @(
        Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles' -Query @{ '$top' = '200' }
    )
    $addressedStates = @('completed', 'ignored', 'thirdParty', 'riskAccepted', 'alternativeMitigation')
    $openControlProfiles = @($controlProfiles | Where-Object {
        $tier = [string](Get-ObjectPropertyValue -InputObject $_ -Name 'tier' -Default '')
        if ($tier -eq 'informational') { return $false }

        $controlStateUpdates = @(Get-ObjectPropertyValue -InputObject $_ -Name 'controlStateUpdates' -Default @())
        @($controlStateUpdates | Where-Object {
            [string](Get-ObjectPropertyValue -InputObject $_ -Name 'state' -Default '') -in $addressedStates
        }).Count -eq 0
    } | Sort-Object {
        Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'rank') -Default 99999
    })

    $topOpenProfiles = @($openControlProfiles | Select-Object -First 10 | ForEach-Object {
        [ordered]@{
            Id = Get-ObjectPropertyValue -InputObject $_ -Name 'id' -Default 'not collected'
            Title = Get-ObjectPropertyValue -InputObject $_ -Name 'title' -Default 'not collected'
            Rank = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'rank') -Default 0
            MaxScore = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'maxScore') -Default 0
            ControlCategory = Get-ObjectPropertyValue -InputObject $_ -Name 'controlCategory' -Default 'not collected'
            ImplementationCost = Get-ObjectPropertyValue -InputObject $_ -Name 'implementationCost' -Default 'not collected'
            UserImpact = Get-ObjectPropertyValue -InputObject $_ -Name 'userImpact' -Default 'not collected'
            Service = Get-ObjectPropertyValue -InputObject $_ -Name 'service' -Default 'not collected'
            Remediation = Get-ObjectPropertyValue -InputObject $_ -Name 'remediation' -Default 'not collected'
            ActionUrl = Get-ObjectPropertyValue -InputObject $_ -Name 'actionUrl' -Default 'not collected'
        }
    })

    $controlProfilesPayload = if ($IncludeDetailedData) {
        @($controlProfiles | ForEach-Object {
            [ordered]@{
                Id = Get-ObjectPropertyValue -InputObject $_ -Name 'id' -Default 'not collected'
                Title = Get-ObjectPropertyValue -InputObject $_ -Name 'title' -Default 'not collected'
                Rank = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'rank') -Default 0
                MaxScore = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'maxScore') -Default 0
                ControlCategory = Get-ObjectPropertyValue -InputObject $_ -Name 'controlCategory' -Default 'not collected'
                Service = Get-ObjectPropertyValue -InputObject $_ -Name 'service' -Default 'not collected'
                ImplementationCost = Get-ObjectPropertyValue -InputObject $_ -Name 'implementationCost' -Default 'not collected'
                UserImpact = Get-ObjectPropertyValue -InputObject $_ -Name 'userImpact' -Default 'not collected'
                Threats = @(Get-ObjectPropertyValue -InputObject $_ -Name 'threats' -Default @())
                Tier = Get-ObjectPropertyValue -InputObject $_ -Name 'tier' -Default 'not collected'
                Remediation = Get-ObjectPropertyValue -InputObject $_ -Name 'remediation' -Default 'not collected'
                RemediationImpact = Get-ObjectPropertyValue -InputObject $_ -Name 'remediationImpact' -Default 'not collected'
                ActionUrl = Get-ObjectPropertyValue -InputObject $_ -Name 'actionUrl' -Default 'not collected'
                Deprecated = [bool](Get-ObjectPropertyValue -InputObject $_ -Name 'deprecated' -Default $false)
            }
        })
    }
    else {
        @()
    }

    $status = 'reported'
    $reason = ''
    if (@($controlProfiles).Count -eq 0) {
        $status = 'partial'
        $reason = 'Secure Score data was collected, but supporting control-profile data returned no results.'
    }

    return [ordered]@{
        Status = $status
        Reason = $reason
        CurrentScore = $currentScore
        MaxScore = $maxScore
        ScorePercent = $scorePercent
        CurrentScoreDate = Get-ObjectPropertyValue -InputObject $latestScore -Name 'createdDateTime' -Default 'not collected'
        LicensedUserCount = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $latestScore -Name 'licensedUserCount') -Default 0
        EnabledServices = @(Get-ObjectPropertyValue -InputObject $latestScore -Name 'enabledServices' -Default @())
        VendorInformation = Get-ObjectPropertyValue -InputObject $latestScore -Name 'vendorInformation' -Default 'not collected'
        ScoreHistory = $scoreHistory
        ControlScoresSummary = [ordered]@{
            TotalControls = @($controlScores).Count
            ControlsBelowMaxScore = @($controlGaps).Count
            TopControlGaps = @($controlGaps | Select-Object -First 10)
        }
        ControlProfilesSummary = [ordered]@{
            TotalProfiles = @($controlProfiles).Count
            OpenProfiles = @($openControlProfiles).Count
            TopOpenProfiles = $topOpenProfiles
        }
        ControlProfilesDetailedIncluded = [bool]$IncludeDetailedData
        ControlProfiles = $controlProfilesPayload
    }
}

function Get-Office365ActiveUserCountsLatestRow {
    # Retrieves the latest row of the Office 365 active user counts report (CSV-based endpoint).
    param([ValidateSet('D7', 'D30', 'D90')][string]$Period = 'D30')

    $rows = @(Get-GraphReportRows -Uri "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserCounts(period='$Period')")
    if (@($rows).Count -eq 0) { return $null }

    return $rows |
        Sort-Object {
            $reportDateValue = Get-RowValueNormalized -Row $_ -CandidateNames @('reportDate', 'report date', 'report refresh date')
            try { [datetime]$reportDateValue } catch { [datetime]::MinValue }
        } -Descending |
        Select-Object -First 1
}

function Get-ActiveUserCountFromRow {
    param(
        [AllowNull()][object]$Row,
        [Parameter(Mandatory)][string[]]$CandidateNames
    )

    if ($null -eq $Row) { return 'not collected' }
    $value = Get-RowValueNormalized -Row $Row -CandidateNames $CandidateNames
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return 0 }
    return Get-IntValue -Value $value -Default 0
}

function Get-AdoptionSignals {
    $latest = $null
    try {
        $latest = Get-Office365ActiveUserCountsLatestRow -Period 'D30'
    }
    catch {
        $latest = $null
    }

    $copilot = Get-CopilotUsageSnapshot -Period 'D30'

    if ($null -eq $latest) {
        return [ordered]@{
            Status                  = if ($copilot.Status -eq 'reported') { 'partial' } else { 'not collected' }
            Reason                  = if ($copilot.Status -eq 'reported') { 'Office 365 active user report was unavailable.' } else { 'Office 365 active user report was unavailable.' }
            ActiveUsers30Days      = 'not collected'
            TeamsActiveUsers       = 'not collected'
            SharePointActiveUsers  = 'not collected'
            OneDriveActiveUsers    = 'not collected'
            OutlookActiveUsers     = 'not collected'
            ReportSource           = 'not collected'
            CopilotActiveUsers30Days = $copilot.ActiveUsers
            CopilotUsageStatus      = $copilot.Status
            CopilotReportSource     = $copilot.Source
            CopilotReportDate       = $copilot.ReportDate
            CopilotCollectionReason = if ([string]::IsNullOrWhiteSpace($copilot.Reason)) { 'reported' } else { $copilot.Reason }
        }
    }

    return [ordered]@{
        Status                  = 'reported'
        Reason                  = ''
        # Per Microsoft Graph, this report returns active-user counts for a report date within the
        # requested report period. This script uses the latest available row as the current adoption snapshot.
        ActiveUsers30Days      = Get-ActiveUserCountFromRow -Row $latest -CandidateNames @('office365', 'office 365')
        TeamsActiveUsers       = Get-ActiveUserCountFromRow -Row $latest -CandidateNames @('teams')
        SharePointActiveUsers  = Get-ActiveUserCountFromRow -Row $latest -CandidateNames @('sharePoint', 'sharepoint')
        OneDriveActiveUsers    = Get-ActiveUserCountFromRow -Row $latest -CandidateNames @('oneDrive', 'onedrive')
        OutlookActiveUsers     = Get-ActiveUserCountFromRow -Row $latest -CandidateNames @('exchange')
        ReportSource           = 'office365ActiveUserCounts'
        CopilotActiveUsers30Days = $copilot.ActiveUsers
        CopilotUsageStatus      = $copilot.Status
        CopilotReportSource     = $copilot.Source
        CopilotReportDate       = $copilot.ReportDate
        CopilotCollectionReason = if ([string]::IsNullOrWhiteSpace($copilot.Reason)) { 'reported' } else { $copilot.Reason }
    }
}

function Get-IdentityAccessAdvanced {
    $identity = $script:Report.Identity
    $metadata = $script:Report.Metadata

    if (-not (Test-SectionCollected -SectionData $identity)) {
        return [ordered]@{
            Status               = 'not collected'
            Reason               = 'Identity baseline data was not collected.'
            AdminRoleObjectsCount = 'not collected'
            MfaCapableUsers      = 'not collected'
            MfaRegisteredUsers   = 'not collected'
            PasswordlessCapableUsers = 'not collected'
            MfaPopulationUserCount = 'not collected'
            MfaPopulationSource  = 'not collected'
            MfaRegistrationPercent = 'not collected'
            PhishingResistantMfaCoverage = 'not collected'
            PimSignal            = 'not collected'
        }
    }

    $mfaCapable = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $identity -Name 'MfaCapableCount') -Default 0
    $mfaRegistered = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $identity -Name 'MfaRegisteredCount') -Default 0
    $passwordless = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $identity -Name 'PasswordlessCount') -Default 0
    $globalAdmins = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $identity -Name 'GlobalAdminCount') -Default 0
    $privilegedRoles = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $identity -Name 'PrivilegedRoleCount') -Default 0
    $adminRoleCount = $globalAdmins + $privilegedRoles
    $mfaPopulationUserCount = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $identity -Name 'MfaPopulationUserCount') -Default 0
    $fallbackTotalUsers = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $metadata -Name 'TotalUsers') -Default 0
    $mfaDenominator = if ($mfaPopulationUserCount -gt 0) { $mfaPopulationUserCount } else { $fallbackTotalUsers }

    return [ordered]@{
        Status               = 'partial'
        Reason               = 'Advanced MFA and PIM signals were not collected.'
        AdminRoleObjectsCount = $adminRoleCount
        MfaCapableUsers       = $mfaCapable
        MfaRegisteredUsers    = $mfaRegistered
        PasswordlessCapableUsers = $passwordless
        MfaPopulationUserCount = if ($mfaDenominator -gt 0) { $mfaDenominator } else { 'not collected' }
        MfaPopulationSource  = Get-ObjectPropertyValue -InputObject $identity -Name 'MfaPopulationSource' -Default 'not collected'
        MfaRegistrationPercent = if ($mfaDenominator -gt 0) { [math]::Round(($mfaRegistered / $mfaDenominator) * 100, 2) } else { 0 }
        PhishingResistantMfaCoverage = 'not collected'
        PimSignal = 'not collected'
        AdminRoleCoverageNote = 'Admin role counts are derived from activated directory roles and active memberships only; PIM-eligible assignments are not included.'
    }
}

function Get-DataProtectionAdvanced {
    $governance = $script:Report.DataGovernance

    if (-not (Test-SectionCollected -SectionData $governance)) {
        return [ordered]@{
            Status                 = 'not collected'
            Reason                 = 'Data governance baseline data was not collected.'
            LabelPoliciesPublished = 'not collected'
            LabelPoliciesTotal     = 'not collected'
            DlpPoliciesEnforced    = 'not collected'
            DlpPoliciesTotal       = 'not collected'
            RetentionPoliciesEnabled = 'not collected'
            RetentionPoliciesTotal = 'not collected'
            eDiscoverySignal       = 'not collected'
            InsiderRiskSignal      = 'not collected'
        }
    }

    $published = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $governance -Name 'SensitivityLabelsPublished') -Default 0
    $labelsTotal = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $governance -Name 'SensitivityLabelsTotal') -Default 0

    $dlpPolicies = @()
    $dlpPoliciesValue = Get-ObjectPropertyValue -InputObject $governance -Name 'DlpPolicies'
    $dlpPoliciesCollected = ($null -ne $dlpPoliciesValue) -and ($dlpPoliciesValue -isnot [string])
    if ($dlpPoliciesCollected) {
        $dlpPolicies = @($dlpPoliciesValue)
    }
    $dlpEnforced = if ($dlpPoliciesCollected) { @($dlpPolicies | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'Mode' -Default '')) -match 'Enforce|Enable' }).Count } else { 'not collected' }

    $retentionPolicies = @()
    $retentionPoliciesValue = Get-ObjectPropertyValue -InputObject $governance -Name 'RetentionPolicies'
    $retentionPoliciesCollected = ($null -ne $retentionPoliciesValue) -and ($retentionPoliciesValue -isnot [string])
    if ($retentionPoliciesCollected) {
        $retentionPolicies = @($retentionPoliciesValue)
    }
    $retentionEnabled = if ($retentionPoliciesCollected) { @($retentionPolicies | Where-Object { (Get-ObjectPropertyValue -InputObject $_ -Name 'Enabled') -eq $true }).Count } else { 'not collected' }

    $status = 'partial'

    return [ordered]@{
        Status                 = $status
        Reason                 = 'eDiscovery and insider risk signals were not collected.'
        LabelPoliciesPublished = $published
        LabelPoliciesTotal     = $labelsTotal
        DlpPoliciesEnforced    = $dlpEnforced
        DlpPoliciesTotal       = if ($dlpPoliciesCollected) { @($dlpPolicies).Count } else { 'not collected' }
        RetentionPoliciesEnabled = $retentionEnabled
        RetentionPoliciesTotal = if ($retentionPoliciesCollected) { @($retentionPolicies).Count } else { 'not collected' }
        eDiscoverySignal       = 'not collected'
        InsiderRiskSignal      = 'not collected'
    }
}

function Get-SharePointExposureAdvanced {
    $sharePoint = $script:Report.SharePointOneDrive
    $sharePointCollected = $sharePoint -and (Test-SectionCollected -SectionData $sharePoint) -and (Test-HashtableKey -InputObject $sharePoint -Key 'ExternalSharingLevel') -and (Test-ValueCollected -Value $sharePoint.ExternalSharingLevel)

    if (-not $sharePointCollected) {
        return [ordered]@{
            Status = 'not collected'
            Reason = 'SharePoint / OneDrive baseline data was not collected.'
            SampledSites = 'not collected'
            ExternalSharingSites = 'not collected'
            AnyoneLinkSites = 'not collected'
            InactiveSites = 'not collected'
            OversharedContentSignal = 'not collected'
            SampledAnyoneLinkCount = 'not collected'
            SampledOrgWideLinkCount = 'not collected'
            UnlabeledContentSignal = 'not collected'
            Truncated = 'not collected'
        }
    }

    $sampleSites = @()
    if ($sharePoint -and (Test-HashtableKey -InputObject $sharePoint -Key 'Sites') -and $sharePoint.Sites -isnot [string]) {
        $sampleSites = @($sharePoint.Sites)
    }

    $externalSharingSites = @($sampleSites | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'SharingCapability' -Default '')) -in @('ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'AnonymousAccess') }).Count
    $anyoneLinkSites = @($sampleSites | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'SharingCapability' -Default '')) -eq 'AnonymousAccess' }).Count
    $inactiveSites = @($sampleSites | Where-Object {
        # Site summaries store this as LastContentModified.
        $lastModified = Get-ObjectPropertyValue -InputObject $_ -Name 'LastContentModified'
        $lastModified -and ((Get-Date).ToUniversalTime() - [DateTime]$lastModified) -gt [TimeSpan]::FromDays(180)
    }).Count

    $oversharing = $null
    if ($sharePoint -and (Test-HashtableKey -InputObject $sharePoint -Key 'OversharingSignals')) {
        $oversharing = $sharePoint.OversharingSignals
    }

    $oversharingStatus = [string](Get-ObjectPropertyValue -InputObject $oversharing -Name 'Status' -Default 'not collected')
    $sampledAnyoneLinks = Get-ObjectPropertyValue -InputObject $oversharing -Name 'SampledAnyoneLinkCount' -Default 'not collected'
    $sampledOrgWideLinks = Get-ObjectPropertyValue -InputObject $oversharing -Name 'SampledOrgWideLinkCount' -Default 'not collected'

    $oversharedContentSignal = 'not collected'
    if ($oversharingStatus -eq 'reported') {
        $anyoneCountNumeric = Get-IntValue -Value $sampledAnyoneLinks -Default 0
        $orgWideCountNumeric = Get-IntValue -Value $sampledOrgWideLinks -Default 0

        if (($anyoneCountNumeric + $orgWideCountNumeric) -gt 0) {
            $oversharedContentSignal = 'elevated'
        }
        else {
            $oversharedContentSignal = 'low'
        }
    }

    return [ordered]@{
        Status = 'partial'
        Reason = 'Some oversharing and unlabeled content signals were not collected.'
        SampledSites = @($sampleSites).Count
        ExternalSharingSites = $externalSharingSites
        AnyoneLinkSites = $anyoneLinkSites
        InactiveSites = $inactiveSites
        OversharedContentSignal = $oversharedContentSignal
        SampledAnyoneLinkCount = $sampledAnyoneLinks
        SampledOrgWideLinkCount = $sampledOrgWideLinks
        UnlabeledContentSignal = 'not collected'
        Truncated = if ($sharePoint -and (Test-HashtableKey -InputObject $sharePoint -Key 'Truncated') -and (Test-ValueCollected -Value $sharePoint.Truncated)) { [bool]$sharePoint.Truncated } else { $false }
    }
}

function Get-TeamsAdvanced {
    $teams = $script:Report.Teams

    if (-not (Test-SectionCollected -SectionData $teams)) {
        return [ordered]@{
            Status                   = 'not collected'
            Reason                   = 'Teams baseline data was not collected.'
            TotalTeams               = 'not collected'
            PublicTeams              = 'not collected'
            StaleTeams               = 'not collected'
            GuestAccessPolicySignal  = 'not collected'
            ExternalAccessPolicySignal = 'not collected'
            OwnerlessTeamsSignal     = 'not collected'
        }
    }

    $totalTeams = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $teams -Name 'TeamCount') -Default 0
    $publicTeams = if (Test-HashtableKey -InputObject $teams -Key 'PublicTeamCount') { Get-IntValue -Value $teams.PublicTeamCount -Default 0 } else { 'not collected' }
    $staleTeams = if (Test-HashtableKey -InputObject $teams -Key 'StaleTeamCount') { Get-IntValue -Value $teams.StaleTeamCount -Default 0 } else { 'not collected' }

    return [ordered]@{
        Status = 'partial'
        Reason = 'Guest access, external access, and ownerless team signals were not collected.'
        TotalTeams = $totalTeams
        PublicTeams = $publicTeams
        StaleTeams  = $staleTeams
        GuestAccessPolicySignal = 'not collected'
        ExternalAccessPolicySignal = 'not collected'
        OwnerlessTeamsSignal = 'not collected'
    }
}

function Get-SearchSemanticAdvanced {
    $searchIndex = $script:Report.SearchIndex

    if (-not (Test-SectionCollected -SectionData $searchIndex)) {
        return [ordered]@{
            Status                     = 'not collected'
            Reason                     = 'Search index baseline data was not collected.'
            ConnectorCount             = 'not collected'
            ReadyConnectorCount        = 'not collected'
            ConnectorHealthSignal      = 'not collected'
            SharePointIndexabilitySignal = 'not collected'
        }
    }

    $connectorCount = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $searchIndex -Name 'ExternalConnectionsCount') -Default 0
    $readyConnectors = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $searchIndex -Name 'ReadyExternalConnectionsCount') -Default 0

    return [ordered]@{
        Status = 'partial'
        Reason = 'SharePoint indexability signal was not collected.'
        ConnectorCount = $connectorCount
        ReadyConnectorCount = $readyConnectors
        ConnectorHealthSignal = if ($connectorCount -gt 0) { 'partial' } else { 'not collected' }
        SharePointIndexabilitySignal = 'not collected'
    }
}

function Get-EndpointAppAdvanced {
    $hasDeviceReadPermission = Test-GraphPermissionsAvailable -RequiredPermissions @('DeviceManagementManagedDevices.Read.All')
    $managedDevices = @()
    if ($hasDeviceReadPermission) {
        $managedDevices = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices' -Query @{ '$select' = 'id,complianceState,operatingSystem' }))
    }
    $managedDevicesReadable = $hasDeviceReadPermission -and (@($managedDevices).Count -gt 0)
    $compliant = @($managedDevices | Where-Object { $_.PSObject.Properties.Name -contains 'ComplianceState' -and $_.ComplianceState -eq 'compliant' }).Count
    $windows = @($managedDevices | Where-Object { $_.PSObject.Properties.Name -contains 'OperatingSystem' -and $_.OperatingSystem -match 'Windows' }).Count

    $appsReport = @(Get-GraphReportRows -Uri "https://graph.microsoft.com/v1.0/reports/getM365AppUserCounts(period='D30')")

    if (-not $managedDevicesReadable -and @($appsReport).Count -eq 0) {
        return [ordered]@{
            Status                = 'not collected'
            Reason                = 'Neither device compliance nor M365 Apps report data was collected.'
            ManagedDeviceCount    = 'not collected'
            CompliantDeviceCount  = 'not collected'
            WindowsDeviceCount    = 'not collected'
            DeviceCompliancePercent = 'not collected'
            M365AppReportAvailable = 'not collected'
            DeviceManagementSignal = if (-not $hasDeviceReadPermission) { 'not collected (missing Graph permission: DeviceManagementManagedDevices.Read.All)' } else { 'not collected (requires DeviceManagement managedDevices read permissions)' }
            BrowserReadinessSignal = 'not collected'
        }
    }

    return [ordered]@{
        Status = 'partial'
        Reason = 'Browser readiness signal was not collected.'
        ManagedDeviceCount = if ($managedDevicesReadable) { @($managedDevices).Count } else { 'not collected' }
        CompliantDeviceCount = if ($managedDevicesReadable) { $compliant } else { 'not collected' }
        WindowsDeviceCount = if ($managedDevicesReadable) { $windows } else { 'not collected' }
        DeviceCompliancePercent = if ($managedDevicesReadable -and @($managedDevices).Count -gt 0) { [math]::Round(($compliant / @($managedDevices).Count) * 100, 2) } else { 'not collected' }
        M365AppReportAvailable = @($appsReport).Count -gt 0
        DeviceManagementSignal = if ($managedDevicesReadable) { 'reported' } elseif (-not $hasDeviceReadPermission) { 'not collected (missing Graph permission: DeviceManagementManagedDevices.Read.All)' } else { 'not collected (requires DeviceManagement managedDevices read permissions)' }
        BrowserReadinessSignal = 'not collected'
    }
}

function Get-AdoptionAdvanced {
    $latestD7 = $null
    $latestD30 = $null
    $latestD90 = $null
    try { $latestD7 = Get-Office365ActiveUserCountsLatestRow -Period 'D7' } catch { $latestD7 = $null }
    try { $latestD30 = Get-Office365ActiveUserCountsLatestRow -Period 'D30' } catch { $latestD30 = $null }
    try { $latestD90 = Get-Office365ActiveUserCountsLatestRow -Period 'D90' } catch { $latestD90 = $null }
    $copilotD7 = Get-CopilotUsageSnapshot -Period 'D7'
    $copilotD30 = Get-CopilotUsageSnapshot -Period 'D30'
    $copilotD90 = Get-CopilotUsageSnapshot -Period 'D90'

    $status = 'reported'
    $reason = ''
    if (-not $latestD7 -and -not $latestD30 -and -not $latestD90 -and $copilotD7.Status -ne 'reported' -and $copilotD30.Status -ne 'reported' -and $copilotD90.Status -ne 'reported') {
        $status = 'not collected'
        $reason = 'No adoption reports were collected.'
    }
    elseif (-not $latestD7 -or -not $latestD30 -or -not $latestD90 -or $copilotD7.Status -ne 'reported' -or $copilotD30.Status -ne 'reported' -or $copilotD90.Status -ne 'reported') {
        $status = 'partial'
        $reason = 'Only part of the adoption trend data was collected.'
    }

    return [ordered]@{
        Status = $status
        Reason = $reason
        ActiveUsersD7  = if ($latestD7) { Get-ActiveUserCountFromRow -Row $latestD7 -CandidateNames @('office365', 'office 365') } else { 'not collected' }
        ActiveUsersD30 = if ($latestD30) { Get-ActiveUserCountFromRow -Row $latestD30 -CandidateNames @('office365', 'office 365') } else { 'not collected' }
        ActiveUsersD90 = if ($latestD90) { Get-ActiveUserCountFromRow -Row $latestD90 -CandidateNames @('office365', 'office 365') } else { 'not collected' }
        TrendSignal = if ($latestD7 -and $latestD30 -and $latestD90) { 'reported' } elseif ($latestD7 -or $latestD30 -or $latestD90) { 'partial' } else { 'not collected' }
        CopilotActiveUsersD7 = $copilotD7.ActiveUsers
        CopilotActiveUsersD30 = $copilotD30.ActiveUsers
        CopilotActiveUsersD90 = $copilotD90.ActiveUsers
        CopilotTrendSignal = if ($copilotD7.Status -eq 'reported' -and $copilotD30.Status -eq 'reported' -and $copilotD90.Status -eq 'reported') { 'reported' } elseif ($copilotD7.Status -eq 'reported' -or $copilotD30.Status -eq 'reported' -or $copilotD90.Status -eq 'reported') { 'partial' } else { 'not collected' }
        CopilotReportSource = if ($copilotD30.Status -eq 'reported') { $copilotD30.Source } elseif ($copilotD7.Status -eq 'reported') { $copilotD7.Source } else { $copilotD90.Source }
        CopilotCollectionReason = if ($copilotD30.Status -eq 'not collected' -and -not [string]::IsNullOrWhiteSpace($copilotD30.Reason)) { $copilotD30.Reason } elseif ($copilotD7.Status -eq 'not collected' -and -not [string]::IsNullOrWhiteSpace($copilotD7.Reason)) { $copilotD7.Reason } elseif ($copilotD90.Status -eq 'not collected' -and -not [string]::IsNullOrWhiteSpace($copilotD90.Reason)) { $copilotD90.Reason } else { 'reported' }
    }
}

function Get-AppGovernanceAdvanced {
    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('Directory.Read.All'))) {
        return [ordered]@{
            Status = 'not collected'
            Reason = 'Missing Graph permission: Directory.Read.All'
            ServicePrincipalCount = 'not collected'
            OAuthGrantCount = 'not collected'
            HighRiskGrantCount = 'not collected'
            ApplicationPermissionGrantCount = 'not collected'
            HighRiskApplicationPermissionGrantCount = 'not collected'
            HighRiskApplicationPermissions = @()
            ApplicationPermissionSignal = 'not collected'
            StaleGrantSignal = 'not collected'
            OwnerlessAppSignal = 'not collected'
        }
    }

    $servicePrincipals = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Query @{ '$select' = 'id,appId,displayName,accountEnabled' }))
    $oauthGrants = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/oauth2PermissionGrants'))

    $highScopePattern = 'Mail\.ReadWrite|Files\.ReadWrite\.All|Sites\.FullControl\.All|Directory\.ReadWrite\.All'
    $highScopeGrants = @(
        $oauthGrants | Where-Object {
            $_.PSObject.Properties.Name -contains 'scope' -and
            [string]$_.scope -match $highScopePattern
        }
    ).Count

    $graphResourceQuery = @{
        '$filter' = "appId eq '00000003-0000-0000-c000-000000000000'"
        '$select' = 'id,appId,displayName,appRoles'
    }
    $graphResourceSp = @(
        Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Query $graphResourceQuery | Select-Object -First 1
    )
    $applicationPermissionAssignments = @()
    $appRoleValueById = @{}

    if ($graphResourceSp -and $graphResourceSp[0] -and $graphResourceSp[0].PSObject.Properties.Name -contains 'id' -and -not [string]::IsNullOrWhiteSpace([string]$graphResourceSp[0].id)) {
        foreach ($role in @($graphResourceSp[0].appRoles)) {
            if ($null -eq $role) { continue }
            $allowedMemberTypes = @($role.AllowedMemberTypes)
            if (@($allowedMemberTypes | Where-Object { $_ -eq 'Application' }).Count -eq 0) { continue }

            $roleId = [string]$role.Id
            if ([string]::IsNullOrWhiteSpace($roleId)) { continue }
            $appRoleValueById[$roleId.ToLowerInvariant()] = [string]$role.Value
        }

        $applicationPermissionAssignments = @(
            Get-GraphRestAll -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($graphResourceSp[0].id)/appRoleAssignedTo" -Query @{ '$select' = 'id,principalId,principalDisplayName,appRoleId,createdDateTime' }
        )
    }

    $highRiskApplicationPermissions = @(
        'Mail.ReadWrite',
        'Mail.Send',
        'Files.ReadWrite.All',
        'Files.Read.All',
        'Sites.FullControl.All',
        'Sites.ReadWrite.All',
        'Directory.ReadWrite.All'
    )

    $highRiskAppPermissionGrantCount = 0
    $highRiskAppPermissionValues = @{}
    foreach ($assignment in @($applicationPermissionAssignments)) {
        if ($null -eq $assignment) { continue }
        if (-not ($assignment.PSObject.Properties.Name -contains 'appRoleId')) { continue }

        $assignmentRoleId = [string]$assignment.appRoleId
        if ([string]::IsNullOrWhiteSpace($assignmentRoleId)) { continue }

        $roleValue = $null
        $normalizedRoleId = $assignmentRoleId.ToLowerInvariant()
        if ($appRoleValueById.ContainsKey($normalizedRoleId)) {
            $roleValue = [string]$appRoleValueById[$normalizedRoleId]
        }

        if ([string]::IsNullOrWhiteSpace($roleValue)) { continue }
        if ($roleValue -in $highRiskApplicationPermissions) {
            $highRiskAppPermissionGrantCount++
            $highRiskAppPermissionValues[$roleValue] = $true
        }
    }

    return [ordered]@{
        Status = 'partial'
        Reason = 'Ownerless app and stale grant signals were not collected.'
        ServicePrincipalCount = @($servicePrincipals).Count
        OAuthGrantCount = @($oauthGrants).Count
        HighRiskGrantCount = $highScopeGrants
        ApplicationPermissionGrantCount = if (@($applicationPermissionAssignments).Count -gt 0) { @($applicationPermissionAssignments).Count } else { 'not collected' }
        HighRiskApplicationPermissionGrantCount = if (@($applicationPermissionAssignments).Count -gt 0) { $highRiskAppPermissionGrantCount } else { 'not collected' }
        HighRiskApplicationPermissions = if (@($applicationPermissionAssignments).Count -gt 0) { @($highRiskAppPermissionValues.Keys | Sort-Object) } else { @() }
        ApplicationPermissionSignal = if (@($applicationPermissionAssignments).Count -gt 0) { 'reported' } else { 'not collected (missing app role assignment read permissions or no application grants available)' }
        StaleGrantSignal = 'not collected'
        OwnerlessAppSignal = 'not collected'
    }
}

function Get-ReadinessFlags {
    $lic = $script:Report.Licensing
    $id = $script:Report.Identity
    $ca = $script:Report.ConditionalAccess
    $gov = $script:Report.DataGovernance
    $spo = $script:Report.SharePointOneDrive
    $ad = $script:Report.AdoptionSignals

    $licPrereqReadyPercentValue = Get-ObjectPropertyValue -InputObject $lic -Name 'PrereqReadyPercent' -Default 'not collected'
    $licPercent = if (Test-ValueCollected -Value $licPrereqReadyPercentValue) { Get-DoubleValue -Value $licPrereqReadyPercentValue -Default 0 } else { 0 }
    $licPercentCollected = Test-ValueCollected -Value $licPrereqReadyPercentValue
    $totalUsersValue = Get-ObjectPropertyValue -InputObject $script:Report.Metadata -Name 'TotalUsers' -Default 'not collected'
    $totalUsers = if (Test-ValueCollected -Value $totalUsersValue) { Get-IntValue -Value $totalUsersValue -Default 0 } else { 0 }
    $mfaRegisteredCountValue = Get-ObjectPropertyValue -InputObject $id -Name 'MfaRegisteredCount' -Default 'not collected'
    $mfaRegisteredCount = if (Test-ValueCollected -Value $mfaRegisteredCountValue) { Get-IntValue -Value $mfaRegisteredCountValue -Default 0 } else { 0 }
    $mfaRegisteredCountCollected = Test-ValueCollected -Value $mfaRegisteredCountValue
    $mfaPopulationUserCountValue = Get-ObjectPropertyValue -InputObject $id -Name 'MfaPopulationUserCount' -Default 'not collected'
    $mfaPopulationUserCount = if (Test-ValueCollected -Value $mfaPopulationUserCountValue) { Get-IntValue -Value $mfaPopulationUserCountValue -Default 0 } else { 0 }
    $activeUsers30DaysValue = Get-ObjectPropertyValue -InputObject $ad -Name 'ActiveUsers30Days' -Default 'not collected'
    $mfaDenominator = if ($mfaPopulationUserCount -gt 0) { $mfaPopulationUserCount } else { $totalUsers }
    $mfaPercent = if ($mfaDenominator -gt 0 -and $mfaRegisteredCountCollected) { [math]::Round(($mfaRegisteredCount / $mfaDenominator) * 100, 2) } else { 'not collected' }
    # NOTE: getOffice365ActiveUserCounts returns daily active-user counts by workload within the
    # requested reporting period. This script uses the latest reported day as a simple adoption
    # signal, not a tenant-wide unique-user measure across the entire period.
    $activeUserPercent = if ($totalUsers -gt 0 -and (Test-ValueCollected -Value $activeUsers30DaysValue)) { [math]::Round((([double](Get-IntValue -Value $activeUsers30DaysValue -Default 0)) / $totalUsers) * 100, 2) } else { 'not collected' }

    $caEnabledPoliciesValue = Get-ObjectPropertyValue -InputObject $ca -Name 'EnabledPolicies' -Default 'not collected'
    $caEnabledPoliciesCollected = Test-ValueCollected -Value $caEnabledPoliciesValue
    $caEnabledPolicies = if ($caEnabledPoliciesCollected) { Get-IntValue -Value $caEnabledPoliciesValue -Default 0 } else { 'not collected' }
    $govPublishedValue = Get-ObjectPropertyValue -InputObject $gov -Name 'SensitivityLabelsPublished' -Default 'not collected'
    $govPublished = if ((Test-SectionCollected -SectionData $gov) -and (Test-ValueCollected -Value $govPublishedValue)) { Get-IntValue -Value $govPublishedValue -Default 0 } else { 'not collected' }
    $spoRestricted = Get-ObjectPropertyValue -InputObject $spo -Name 'RestrictedSearchEnabled' -Default 'not collected'
    $spoOneDrivePercentValue = Get-ObjectPropertyValue -InputObject $spo -Name 'OneDrivePersonalSiteSharePercent' -Default 'not collected'
    $spoOneDrivePercentCollected = Test-ValueCollected -Value $spoOneDrivePercentValue
    $spoOneDrivePercent = if ($spoOneDrivePercentCollected) { Get-DoubleValue -Value $spoOneDrivePercentValue -Default 0 } else { 'not collected' }
    $spoExternalSharing = [string](Get-ObjectPropertyValue -InputObject $spo -Name 'ExternalSharingLevel' -Default 'not collected')

    return [ordered]@{
        LicencePrereqPercent      = if ($licPercentCollected) { $licPercent } else { 'not collected' }
        MfaRegisteredPercent      = $mfaPercent
        MfaPopulationUserCount    = if ($mfaDenominator -gt 0) { $mfaDenominator } else { 'not collected' }
        MfaPopulationSource       = Get-ObjectPropertyValue -InputObject $id -Name 'MfaPopulationSource' -Default 'not collected'
        CaPoliciesEnabled         = $caEnabledPolicies
        SensitivityLabelsPublished = if (Test-ValueCollected -Value $govPublished) { ($govPublished -gt 0) } else { 'not collected' }
        RestrictedSearchEnabled   = if (Test-ValueCollected -Value $spoRestricted) { [bool]$spoRestricted } else { 'not collected' }
        OneDrivePersonalSiteSharePercent = $spoOneDrivePercent
        OneDriveCoveragePercent   = $spoOneDrivePercent
        ExternalSharingLevel      = $spoExternalSharing
        ActiveUsersInReportPeriodPercent = $activeUserPercent
        ActiveUserPercentByWorkload = $activeUserPercent
    }
}

function Get-ReadinessEvidence {
    $flags = $script:Report.ReadinessFlags
    $identity = $script:Report.IdentityAccessAdvanced
    $gov = $script:Report.DataProtectionAdvanced
    $spo = $script:Report.SharePointExposureAdvanced
    $apps = $script:Report.AppGovernanceAdvanced
    $endpoint = $script:Report.EndpointAppAdvanced
    $scoringModelVar = Get-Variable -Name ScoringModel -Scope Script -ErrorAction SilentlyContinue
    if ($null -eq $scoringModelVar -or $null -eq $scoringModelVar.Value) {
        $script:ScoringModel = Get-DefaultScoringModel
    }
    $scoring = $script:ScoringModel

    $advancedMfaCoverage = Get-ObjectPropertyValue -InputObject $identity -Name 'PhishingResistantMfaCoverage' -Default 'not collected'

    $mfaRegisteredPercentValue = Get-ObjectPropertyValue -InputObject $flags -Name 'MfaRegisteredPercent' -Default 'not collected'
    $caEnabledPoliciesValue = Get-ObjectPropertyValue -InputObject $flags -Name 'CaPoliciesEnabled' -Default 'not collected'
    $mfaRegisteredPercent = if (Test-ValueCollected -Value $mfaRegisteredPercentValue) { Get-DoubleValue -Value $mfaRegisteredPercentValue -Default 0 } else { $null }
    $caEnabledPolicies = if (Test-ValueCollected -Value $caEnabledPoliciesValue) { Get-IntValue -Value $caEnabledPoliciesValue -Default 0 } else { $null }
    $identityAdvancedMfaPercent = if (Test-ValueCollected -Value $advancedMfaCoverage) { Get-DoubleValue -Value $advancedMfaCoverage -Default 0 } else { $null }

    $identityScoreRaw = 0
    $identityMeasuredWeight = 0
    $identityMeasured = $false
    if ($null -ne $mfaRegisteredPercent) {
        $identityMeasured = $true
        $identityMeasuredWeight += [int]$scoring.IdentityAccess.MfaRegisteredWeight
        if ($mfaRegisteredPercent -ge [double]$scoring.IdentityAccess.MfaRegisteredPercentThreshold) { $identityScoreRaw += [int]$scoring.IdentityAccess.MfaRegisteredWeight }
    }
    if ($null -ne $identityAdvancedMfaPercent) {
        $identityMeasured = $true
        $identityMeasuredWeight += [int]$scoring.IdentityAccess.AdvancedMfaWeight
        if ($identityAdvancedMfaPercent -ge [double]$scoring.IdentityAccess.AdvancedMfaPercentThreshold) { $identityScoreRaw += [int]$scoring.IdentityAccess.AdvancedMfaWeight }
    }
    if ($null -ne $caEnabledPolicies) {
        $identityMeasured = $true
        $identityMeasuredWeight += [int]$scoring.IdentityAccess.ConditionalAccessWeight
        if ($caEnabledPolicies -ge [int]$scoring.IdentityAccess.ConditionalAccessPolicyThreshold) { $identityScoreRaw += [int]$scoring.IdentityAccess.ConditionalAccessWeight }
    }
    $identityScore = if ($identityMeasuredWeight -gt 0) { [math]::Round(($identityScoreRaw / $identityMeasuredWeight) * 100, 2) } else { $null }

    $governanceScoreRaw = 0
    $governanceMeasuredWeight = 0
    $labelsPublishedValue = Get-ObjectPropertyValue -InputObject $flags -Name 'SensitivityLabelsPublished' -Default 'not collected'
    $labelsPublished = if (Test-ValueCollected -Value $labelsPublishedValue) { [bool]$labelsPublishedValue } else { $null }
    $dlpEnforcedValue = Get-ObjectPropertyValue -InputObject $gov -Name 'DlpPoliciesEnforced' -Default 'not collected'
    $dlpEnforced = if (Test-ValueCollected -Value $dlpEnforcedValue) { Get-IntValue -Value $dlpEnforcedValue -Default 0 } else { $null }
    $retentionEnabledValue = Get-ObjectPropertyValue -InputObject $gov -Name 'RetentionPoliciesEnabled' -Default 'not collected'
    $retentionEnabled = if (Test-ValueCollected -Value $retentionEnabledValue) { Get-IntValue -Value $retentionEnabledValue -Default 0 } else { $null }
    $governanceMeasured = $false
    if ($null -ne $labelsPublished) {
        $governanceMeasured = $true
        $governanceMeasuredWeight += [int]$scoring.DataGovernance.LabelsPublishedWeight
        if ($labelsPublished) { $governanceScoreRaw += [int]$scoring.DataGovernance.LabelsPublishedWeight }
    }
    if ($null -ne $dlpEnforced) {
        $governanceMeasured = $true
        $governanceMeasuredWeight += [int]$scoring.DataGovernance.DlpEnforcedWeight
        if ($dlpEnforced -gt 0) { $governanceScoreRaw += [int]$scoring.DataGovernance.DlpEnforcedWeight }
    }
    if ($null -ne $retentionEnabled) {
        $governanceMeasured = $true
        $governanceMeasuredWeight += [int]$scoring.DataGovernance.RetentionEnabledWeight
        if ($retentionEnabled -gt 0) { $governanceScoreRaw += [int]$scoring.DataGovernance.RetentionEnabledWeight }
    }
    $governanceScore = if ($governanceMeasuredWeight -gt 0) { [math]::Round(($governanceScoreRaw / $governanceMeasuredWeight) * 100, 2) } else { $null }

    $contentScoreRaw = 0
    $contentMeasuredWeight = 0
    $oneDriveCoveragePercentValue = Get-ObjectPropertyValue -InputObject $flags -Name 'OneDrivePersonalSiteSharePercent' -Default (Get-ObjectPropertyValue -InputObject $flags -Name 'OneDriveCoveragePercent' -Default 'not collected')
    $oneDriveCoveragePercent = if (Test-ValueCollected -Value $oneDriveCoveragePercentValue) { Get-DoubleValue -Value $oneDriveCoveragePercentValue -Default 0 } else { $null }
    $anyoneLinkSitesValue = Get-ObjectPropertyValue -InputObject $spo -Name 'AnyoneLinkSites' -Default 'not collected'
    $anyoneLinkSites = if (Test-ValueCollected -Value $anyoneLinkSitesValue) { Get-IntValue -Value $anyoneLinkSitesValue -Default 0 } else { $null }
    $restrictedSearchEnabledValue = Get-ObjectPropertyValue -InputObject $flags -Name 'RestrictedSearchEnabled' -Default 'not collected'
    $restrictedSearchEnabled = if (Test-ValueCollected -Value $restrictedSearchEnabledValue) { [bool]$restrictedSearchEnabledValue } else { $null }
    $contentMeasured = $false
    if ($null -ne $oneDriveCoveragePercent) {
        $contentMeasured = $true
        $contentMeasuredWeight += [int]$scoring.ContentExposure.OneDriveCoverageWeight
        if ($oneDriveCoveragePercent -ge [double]$scoring.ContentExposure.OneDriveCoverageThreshold) { $contentScoreRaw += [int]$scoring.ContentExposure.OneDriveCoverageWeight }
    }
    if ($null -ne $anyoneLinkSites) {
        $contentMeasured = $true
        $contentMeasuredWeight += [int]$scoring.ContentExposure.NoAnyoneLinksWeight
        if ($anyoneLinkSites -eq 0) { $contentScoreRaw += [int]$scoring.ContentExposure.NoAnyoneLinksWeight } else { $contentScoreRaw += [int]$scoring.ContentExposure.AnyoneLinksPresentWeight }
    }
    if ($null -ne $restrictedSearchEnabled) {
        $contentMeasured = $true
        $contentMeasuredWeight += [int]$scoring.ContentExposure.RestrictedSearchWeight
        if ($restrictedSearchEnabled) { $contentScoreRaw += [int]$scoring.ContentExposure.RestrictedSearchWeight }
    }
    $contentScore = if ($contentMeasuredWeight -gt 0) { [math]::Round(($contentScoreRaw / $contentMeasuredWeight) * 100, 2) } else { $null }

    $adoptionScoreRaw = 0
    $adoptionMeasuredWeight = 0
    $activeUserPercentValue = Get-ObjectPropertyValue -InputObject $flags -Name 'ActiveUsersInReportPeriodPercent' -Default (Get-ObjectPropertyValue -InputObject $flags -Name 'ActiveUserPercentByWorkload' -Default 'not collected')
    $activeUserPercent = if (Test-ValueCollected -Value $activeUserPercentValue) { Get-DoubleValue -Value $activeUserPercentValue -Default 0 } else { $null }
    $adoptionTrendSignal = [string](Get-ObjectPropertyValue -InputObject $script:Report.AdoptionAdvanced -Name 'TrendSignal' -Default 'not collected')
    $adoptionMeasured = $false
    if ($null -ne $activeUserPercent) {
        $adoptionMeasured = $true
        $adoptionMeasuredWeight += [int]$scoring.Adoption.ActiveUserHighWeight
        if ($activeUserPercent -ge [double]$scoring.Adoption.ActiveUserHighThreshold) { $adoptionScoreRaw += [int]$scoring.Adoption.ActiveUserHighWeight }
        elseif ($activeUserPercent -ge [double]$scoring.Adoption.ActiveUserMediumThreshold) { $adoptionScoreRaw += [int]$scoring.Adoption.ActiveUserMediumWeight }
        else { $adoptionScoreRaw += [int]$scoring.Adoption.ActiveUserLowWeight }
    }
    if ($adoptionTrendSignal -eq 'reported') {
        $adoptionMeasured = $true
        $adoptionMeasuredWeight += [int]$scoring.Adoption.TrendReportedWeight
        $adoptionScoreRaw += [int]$scoring.Adoption.TrendReportedWeight
    }
    $adoptionScore = if ($adoptionMeasuredWeight -gt 0) { [math]::Round(($adoptionScoreRaw / $adoptionMeasuredWeight) * 100, 2) } else { $null }

    $highRiskGrantCountValue = Get-ObjectPropertyValue -InputObject $apps -Name 'HighRiskGrantCount' -Default 'not collected'
    $highRiskGrantCount = if (Test-ValueCollected -Value $highRiskGrantCountValue) { Get-IntValue -Value $highRiskGrantCountValue -Default 0 } else { 'not collected' }
    $deviceCompliancePercentValue = Get-ObjectPropertyValue -InputObject $endpoint -Name 'DeviceCompliancePercent' -Default 'not collected'
    $deviceCompliancePercent = if (Test-ValueCollected -Value $deviceCompliancePercentValue) { Get-DoubleValue -Value $deviceCompliancePercentValue -Default 0 } else { $null }
    $governanceRisk = if (-not (Test-ValueCollected -Value $highRiskGrantCount)) { 'not collected' } elseif ($highRiskGrantCount -gt 0) { 'elevated' } else { 'moderate' }
    $deviceScore = $null
    if ($null -ne $deviceCompliancePercent) {
        if ($deviceCompliancePercent -ge [double]$scoring.EndpointReadiness.ComplianceHighThreshold) {
            $deviceScore = [int]$scoring.EndpointReadiness.HighScore
        }
        elseif ($deviceCompliancePercent -ge [double]$scoring.EndpointReadiness.ComplianceMediumThreshold) {
            $deviceScore = [int]$scoring.EndpointReadiness.MediumScore
        }
        else {
            $deviceScore = [int]$scoring.EndpointReadiness.LowScore
        }
    }

    $availableDomainScores = New-Object System.Collections.Generic.List[double]
    if ($identityMeasured) { $availableDomainScores.Add([double]$identityScore) | Out-Null }
    if ($governanceMeasured) { $availableDomainScores.Add([double]$governanceScore) | Out-Null }
    if ($contentMeasured) { $availableDomainScores.Add([double]$contentScore) | Out-Null }
    if ($adoptionMeasured) { $availableDomainScores.Add([double]$adoptionScore) | Out-Null }
    if ($null -ne $deviceScore) { $availableDomainScores.Add([double]$deviceScore) | Out-Null }

    $overall = if ($availableDomainScores.Count -gt 0) {
        [math]::Round((($availableDomainScores | Measure-Object -Sum).Sum / $availableDomainScores.Count), 2)
    }
    else {
        'not collected'
    }

    $evaluatedSections = @($scoring.Confidence.EvaluatedSections)
    $collectedSections = New-Object System.Collections.Generic.List[string]
    $notCollectedSections = New-Object System.Collections.Generic.List[string]

    foreach ($sectionName in $evaluatedSections) {
        if (-not $script:Report.Contains($sectionName)) {
            $notCollectedSections.Add($sectionName) | Out-Null
            continue
        }

        $sectionData = $script:Report[$sectionName]
        if (Test-SectionCollected -SectionData $sectionData) {
            $collectedSections.Add($sectionName) | Out-Null
        }
        else {
            $notCollectedSections.Add($sectionName) | Out-Null
        }
    }

    $completenessPercent = if (@($evaluatedSections).Count -gt 0) {
        [math]::Round((@($collectedSections).Count / @($evaluatedSections).Count) * 100, 2)
    }
    else {
        0
    }

    $confidenceLevel = if ($completenessPercent -ge [double]$scoring.Confidence.HighThreshold) {
        'high'
    }
    elseif ($completenessPercent -ge [double]$scoring.Confidence.MediumThreshold) {
        'medium'
    }
    else {
        'low'
    }

    $confidenceAdjustedScore = if (Test-ValueCollected -Value $overall) {
        [math]::Round(((Get-DoubleValue -Value $overall -Default 0) * $completenessPercent) / 100, 2)
    }
    else {
        'not collected'
    }

    return [ordered]@{
        ScoringModel = [ordered]@{
            IdentityAccess = $scoring.IdentityAccess
            DataGovernance = $scoring.DataGovernance
            ContentExposure = $scoring.ContentExposure
            Adoption = $scoring.Adoption
            EndpointReadiness = $scoring.EndpointReadiness
        }
        DomainScores = [ordered]@{
            IdentityAccess    = if ($identityMeasured) { $identityScore } else { 'not collected' }
            DataGovernance    = if ($governanceMeasured) { $governanceScore } else { 'not collected' }
            ContentExposure   = if ($contentMeasured) { $contentScore } else { 'not collected' }
            Adoption          = if ($adoptionMeasured) { $adoptionScore } else { 'not collected' }
            EndpointReadiness = if ($null -ne $deviceScore) { $deviceScore } else { 'not collected' }
        }
        DomainCompleteness = [ordered]@{
            IdentityAccess = if ($identityMeasuredWeight -gt 0) { $identityMeasuredWeight } else { 'not collected' }
            DataGovernance = if ($governanceMeasuredWeight -gt 0) { $governanceMeasuredWeight } else { 'not collected' }
            ContentExposure = if ($contentMeasuredWeight -gt 0) { $contentMeasuredWeight } else { 'not collected' }
            Adoption = if ($adoptionMeasuredWeight -gt 0) { $adoptionMeasuredWeight } else { 'not collected' }
            EndpointReadiness = if ($null -ne $deviceScore) { 100 } else { 'not collected' }
        }
        IdentityScoreInputs = [ordered]@{
            MfaRegisteredPercent = $mfaRegisteredPercent
            AdvancedMfaCoveragePercent = if ($advancedMfaCoverage -ne 'not collected') { $identityAdvancedMfaPercent } else { 'not collected' }
            AdvancedMfaCoverageSource = if ($advancedMfaCoverage -ne 'not collected') { 'PhishingResistantMfaCoverage' } else { 'not collected' }
            CaEnabledPolicies = if ($null -ne $caEnabledPolicies) { $caEnabledPolicies } else { 'not collected' }
            MeasuredWeightCoveragePercent = if ($identityMeasuredWeight -gt 0) { $identityMeasuredWeight } else { 'not collected' }
            MeasuredWeightPercent = if ($identityMeasuredWeight -gt 0) { $identityMeasuredWeight } else { 'not collected' }
        }
        ScoredDomainCount = $availableDomainScores.Count
        OverallScore = $overall
        ConfidenceAdjustedScore = $confidenceAdjustedScore
        DataCompleteness = [ordered]@{
            CompletenessScope = 'section-level'
            EvaluatedSectionCount = @($evaluatedSections).Count
            CollectedSectionCount = @($collectedSections).Count
            CompletenessPercent = $completenessPercent
            ConfidenceLevel = $confidenceLevel
            NotCollectedSections = @($notCollectedSections)
        }
        RiskSignals = [ordered]@{
            AppGovernanceRisk = $governanceRisk
            HighRiskAppGrants = $highRiskGrantCount
            AnyoneLinkSites   = if ($null -ne $anyoneLinkSites) { $anyoneLinkSites } else { 'not collected' }
        }
    }
}

function Add-GapEntry {
    param(
        [Parameter(Mandatory)][hashtable]$Groups,
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Section,
        [Parameter(Mandatory)][string]$FieldPath,
        [string]$Reason = '',
        [string[]]$MissingGraphPermissions = @()
    )

    if (-not $Groups.ContainsKey($Category)) {
        $Groups[$Category] = [System.Collections.Generic.List[object]]::new()
    }

    $Groups[$Category].Add([ordered]@{
        Section = $Section
        FieldPath = $FieldPath
        Reason = $Reason
        MissingGraphPermissions = @($MissingGraphPermissions)
    }) | Out-Null
}

function Get-NotCollectedFieldPaths {
    param(
        [AllowNull()][object]$InputObject,
        [string]$CurrentPath = ''
    )

    $results = New-Object System.Collections.Generic.List[string]
    if ($null -eq $InputObject) { return @($results) }

    if ($InputObject -is [string]) {
        if ([string]$InputObject -like 'not collected*') {
            $results.Add($CurrentPath) | Out-Null
        }
        return @($results)
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in @($InputObject.Keys)) {
            $childPath = if ([string]::IsNullOrWhiteSpace($CurrentPath)) { [string]$key } else { "$CurrentPath.$key" }
            foreach ($childResult in @(Get-NotCollectedFieldPaths -InputObject $InputObject[$key] -CurrentPath $childPath)) {
                $results.Add($childResult) | Out-Null
            }
        }
        return @($results)
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
        $index = 0
        foreach ($item in @($InputObject)) {
            $childPath = if ([string]::IsNullOrWhiteSpace($CurrentPath)) { "[$index]" } else { "$CurrentPath[$index]" }
            foreach ($childResult in @(Get-NotCollectedFieldPaths -InputObject $item -CurrentPath $childPath)) {
                $results.Add($childResult) | Out-Null
            }
            $index++
        }
        return @($results)
    }

    foreach ($property in @($InputObject.PSObject.Properties)) {
        $childPath = if ([string]::IsNullOrWhiteSpace($CurrentPath)) { [string]$property.Name } else { "$CurrentPath.$($property.Name)" }
        foreach ($childResult in @(Get-NotCollectedFieldPaths -InputObject $property.Value -CurrentPath $childPath)) {
            $results.Add($childResult) | Out-Null
        }
    }

    return @($results)
}

function Get-GapRootCauseCategory {
    param(
        [string]$SectionName,
        [AllowNull()][object]$SectionData,
        [hashtable]$ModuleAvailability,
        [hashtable]$SessionAvailability
    )

    $reason = [string](Get-ObjectPropertyValue -InputObject $SectionData -Name 'Reason' -Default '')

    if ($reason -match 'Missing Graph permission|missing Graph permission') { return 'MissingGraphConsent' }
    if ($reason -match 'not enabled, unsupported, or not licensed|feature/licensing|not licensed') { return 'FeatureOrLicensingUnavailable' }
    if ($reason -match 'returned no data|returned no results') { return 'EndpointReturnedNoData' }
    if ($reason -match 'session was unavailable|no .* session was available|session details were unavailable') { return 'ServiceSessionUnavailable' }

    switch ($SectionName) {
        'SharePointOneDrive' {
            if (-not $SessionAvailability.SharePointConnected) {
                if (-not $ModuleAvailability.SharePointCmdletsAvailable) { return 'ModuleMissing' }
                return 'ServiceSessionUnavailable'
            }
        }
        'Exchange' {
            if (-not $SessionAvailability.ExchangeConnected) {
                if (-not $ModuleAvailability.ExchangeCmdletsAvailable) { return 'ModuleMissing' }
                return 'ServiceSessionUnavailable'
            }
        }
        'Teams' {
            if (-not $SessionAvailability.TeamsConnected) {
                if (-not $ModuleAvailability.TeamsCmdletsAvailable) { return 'ModuleMissing' }
                return 'ServiceSessionUnavailable'
            }
        }
        'DataGovernance' {
            if (-not $SessionAvailability.ComplianceConnected) {
                if (-not $ModuleAvailability.ComplianceCmdletsAvailable) { return 'ModuleMissing' }
                return 'ServiceSessionUnavailable'
            }
        }
    }

    return 'Other'
}

function Get-PrerequisitesAndGaps {
    $moduleAvailability = [ordered]@{
        GraphCmdletsAvailable = [bool](Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)
        TeamsCmdletsAvailable = [bool](Get-Command Connect-MicrosoftTeams -ErrorAction SilentlyContinue)
        SharePointCmdletsAvailable = [bool](Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)
        ExchangeCmdletsAvailable = [bool](Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)
        ComplianceCmdletsAvailable = [bool](Get-Command Connect-IPPSSession -ErrorAction SilentlyContinue)
    }

    $sessionAvailability = [ordered]@{
        GraphAuthenticated = -not [string]::IsNullOrWhiteSpace([string]$script:GraphAccessToken) -or $null -ne $script:MgContext
        TeamsConnected = [bool]$script:TeamsConnected
        SharePointConnected = [bool]$script:SharePointChildCollectionSucceeded
        ExchangeConnected = $false
        ComplianceConnected = $false
    }

    try {
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $connections = @((Get-ConnectionInformation -ErrorAction SilentlyContinue))
            $sessionAvailability.ExchangeConnected = @($connections | Where-Object {
                ($_.PSObject.Properties.Name -contains 'Name' -and [string]$_.Name -match 'ExchangeOnline') -or
                ($_.PSObject.Properties.Name -contains 'ConnectionUri' -and [string]$_.ConnectionUri -match 'outlook\.office365\.com')
            }).Count -gt 0
            $sessionAvailability.ComplianceConnected = @($connections | Where-Object {
                ($_.PSObject.Properties.Name -contains 'Name' -and [string]$_.Name -match 'IPPSSession|SecurityCompliance') -or
                ($_.PSObject.Properties.Name -contains 'ConnectionUri' -and [string]$_.ConnectionUri -match 'ps\.compliance\.protection\.outlook\.com')
            }).Count -gt 0
        }
    }
    catch {}

    $groupMap = @{}
    $excludedSections = @('CollectorTimings', 'Errors', 'PrerequisitesAndGaps')
    foreach ($sectionName in @($script:Report.Keys | Where-Object { $_ -notin $excludedSections })) {
        $sectionData = $script:Report[$sectionName]
        $fieldPaths = @(Get-NotCollectedFieldPaths -InputObject $sectionData)
        if (@($fieldPaths).Count -eq 0) { continue }

        $category = Get-GapRootCauseCategory -SectionName $sectionName -SectionData $sectionData -ModuleAvailability $moduleAvailability -SessionAvailability $sessionAvailability
        $reason = [string](Get-ObjectPropertyValue -InputObject $sectionData -Name 'Reason' -Default '')
        foreach ($fieldPath in @($fieldPaths | Sort-Object -Unique)) {
            Add-GapEntry -Groups $groupMap -Category $category -Section $sectionName -FieldPath $fieldPath -Reason $reason -MissingGraphPermissions @($script:GraphMissingScopes)
        }
    }

    $orderedCategories = @(
        'MissingGraphConsent',
        'ServiceSessionUnavailable',
        'ModuleMissing',
        'EndpointReturnedNoData',
        'FeatureOrLicensingUnavailable',
        'Other'
    )

    $categoryOutput = [ordered]@{}
    foreach ($categoryName in $orderedCategories) {
        $entries = if ($groupMap.ContainsKey($categoryName)) { @($groupMap[$categoryName]) } else { @() }
        $categoryOutput[$categoryName] = [ordered]@{
            Count = @($entries).Count
            Fields = $entries
        }
    }

    $totalGapFields = @($orderedCategories | ForEach-Object { $categoryOutput[$_].Count } | Measure-Object -Sum).Sum

    return [ordered]@{
        Status = 'reported'
        Reason = ''
        Summary = [ordered]@{
            TotalNotCollectedFields = [int]$totalGapFields
            CategoryCounts = [ordered]@{
                MissingGraphConsent = $categoryOutput.MissingGraphConsent.Count
                ServiceSessionUnavailable = $categoryOutput.ServiceSessionUnavailable.Count
                ModuleMissing = $categoryOutput.ModuleMissing.Count
                EndpointReturnedNoData = $categoryOutput.EndpointReturnedNoData.Count
                FeatureOrLicensingUnavailable = $categoryOutput.FeatureOrLicensingUnavailable.Count
                Other = $categoryOutput.Other.Count
            }
        }
        CurrentPrerequisites = [ordered]@{
            MissingGraphPermissions = @($script:GraphMissingScopes)
            GrantedGraphPermissions = @($script:GraphProvidedScopes)
            ModuleAvailability = $moduleAvailability
            SessionAvailability = $sessionAvailability
        }
        Categories = $categoryOutput
    }
}

function Get-Recommendations {
    $flags = $script:Report.ReadinessFlags
    $evidence = $script:Report.ReadinessEvidence
    $gov = $script:Report.DataProtectionAdvanced
    $spo = $script:Report.SharePointExposureAdvanced
    $apps = $script:Report.AppGovernanceAdvanced

    $topRisks = New-Object System.Collections.Generic.List[string]
    $quickWins = New-Object System.Collections.Generic.List[string]

    $overallScore = Get-ObjectPropertyValue -InputObject $evidence -Name 'OverallScore' -Default 'not collected'

    $mfaRegisteredPercentValue = Get-ObjectPropertyValue -InputObject $flags -Name 'MfaRegisteredPercent' -Default 'not collected'
    $mfaRegisteredPercent = if (Test-ValueCollected -Value $mfaRegisteredPercentValue) { Get-DoubleValue -Value $mfaRegisteredPercentValue -Default 0 } else { $null }
    $labelsPublishedValue = Get-ObjectPropertyValue -InputObject $flags -Name 'SensitivityLabelsPublished' -Default 'not collected'
    $labelsPublished = if (Test-ValueCollected -Value $labelsPublishedValue) { [bool]$labelsPublishedValue } else { $null }
    $activeUserPercentValue = Get-ObjectPropertyValue -InputObject $flags -Name 'ActiveUsersInReportPeriodPercent' -Default (Get-ObjectPropertyValue -InputObject $flags -Name 'ActiveUserPercentByWorkload' -Default 'not collected')
    $activeUserPercent = if (Test-ValueCollected -Value $activeUserPercentValue) { Get-DoubleValue -Value $activeUserPercentValue -Default 0 } else { $null }
    $dlpEnforcedValue = Get-ObjectPropertyValue -InputObject $gov -Name 'DlpPoliciesEnforced' -Default 'not collected'
    $dlpEnforced = if (Test-ValueCollected -Value $dlpEnforcedValue) { Get-IntValue -Value $dlpEnforcedValue -Default 0 } else { $null }
    $anyoneLinkSitesValue = Get-ObjectPropertyValue -InputObject $spo -Name 'AnyoneLinkSites' -Default 'not collected'
    $anyoneLinkSites = if (Test-ValueCollected -Value $anyoneLinkSitesValue) { Get-IntValue -Value $anyoneLinkSitesValue -Default 0 } else { $null }
    $highRiskGrantCountValue = Get-ObjectPropertyValue -InputObject $apps -Name 'HighRiskGrantCount' -Default 'not collected'
    $highRiskGrantCount = if (Test-ValueCollected -Value $highRiskGrantCountValue) { Get-IntValue -Value $highRiskGrantCountValue -Default 0 } else { $null }

    if ($null -ne $mfaRegisteredPercent -and $mfaRegisteredPercent -lt 80) {
        $topRisks.Add('MFA registration is below 80% for users; identity posture is insufficient for broad AI rollout.') | Out-Null
        $quickWins.Add('Increase MFA registration coverage with targeted campaigns and registration policy enforcement.') | Out-Null
    }

    if ($null -ne $labelsPublished -and -not $labelsPublished) {
        $topRisks.Add('Sensitivity labels are not broadly published; data classification coverage is low.') | Out-Null
        $quickWins.Add('Publish baseline sensitivity labels for Exchange, SharePoint, Teams, and Groups.') | Out-Null
    }

    if ($null -ne $dlpEnforced -and $dlpEnforced -eq 0) {
        $topRisks.Add('No enforced DLP policy was detected; risk of unintended data disclosure remains high.') | Out-Null
        $quickWins.Add('Move at least one high-value DLP policy from test mode to enforce mode.') | Out-Null
    }

    if ($null -ne $anyoneLinkSites -and $anyoneLinkSites -gt 0) {
        $topRisks.Add("$($anyoneLinkSites) sampled SharePoint/OneDrive sites allow anonymous links.") | Out-Null
        $quickWins.Add('Tighten external sharing defaults and reduce anonymous link usage in high-value sites.') | Out-Null
    }

    if ($null -ne $highRiskGrantCount -and $highRiskGrantCount -gt 0) {
        $topRisks.Add("Detected $($highRiskGrantCount) high-risk OAuth grants with elevated scopes.") | Out-Null
        $quickWins.Add('Review and remove unnecessary high-scope app grants; enforce app governance approvals.') | Out-Null
    }

    if ($null -ne $activeUserPercent -and $activeUserPercent -lt 30) {
        $quickWins.Add('Run user enablement and scenario-led training to improve 30-day workload adoption.') | Out-Null
    }

    return [ordered]@{
        TopRisks = @($topRisks)
        QuickWins = @($quickWins)
        OverallReadiness = if (-not (Test-ValueCollected -Value $overallScore)) { 'unknown' } elseif ((Get-DoubleValue -Value $overallScore -Default 0) -ge 75) { 'strong' } elseif ((Get-DoubleValue -Value $overallScore -Default 0) -ge 50) { 'moderate' } else { 'needs improvement' }
    }
}

# ============================================================================
# MAIN
# ============================================================================
if ($SharePointCollectionChild) {
    try {
        $WarningPreference = 'SilentlyContinue'
        $InformationPreference = 'SilentlyContinue'
        $VerbosePreference = 'SilentlyContinue'
        $DebugPreference = 'SilentlyContinue'
        $ProgressPreference = 'SilentlyContinue'

        if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
            throw 'SharePointAdminUrl is required in child mode.'
        }

        $null = Import-SilentPnPModule
        $connectResult = Connect-PnPForSharePointCollection -AdminUrl $SharePointAdminUrl
        if (-not $connectResult.Succeeded) {
            throw $connectResult.Reason
        }

        $childResult = Get-SharePointOneDriveReadinessLocal -Sample:$IncludeSampling -Max $SampleSize
        $script:SharePointChildCollectionSucceeded = ([string]$childResult.Status -ne 'not collected')
        $childResult | ConvertTo-Json -Depth 20 -Compress
    }
    catch {
        $script:SharePointChildCollectionSucceeded = $false
        [ordered]@{
            Status                = 'not collected'
            Reason                = "SharePoint child process failed: $($_.Exception.Message)"
            ExternalSharingLevel  = 'not collected'
            DefaultSharingLinkType = 'not collected'
            RestrictedSearchEnabled = 'not collected'
            TotalSiteCount        = 'not collected'
            Truncated             = 'not collected'
            Sites                 = @()
            OneDriveCoveragePercent = 'not collected'
            InactiveSiteCount     = 'not collected'
            OversharingSignals    = [ordered]@{
                Status                    = 'not collected'
                CollectionMethod          = 'not collected'
                CollectionScope           = 'not collected'
                CollectionPathUsed        = 'not collected'
                CollectionReason          = 'not collected'
                RequiredForReliableCollection = 'not collected'
                DocumentationNote         = 'not collected'
                SampledAnyoneLinkCount    = 'not collected'
                SampledOrgWideLinkCount   = 'not collected'
                SitesWithAnyoneLinksCount = 'not collected'
            }
        } | ConvertTo-Json -Depth 20 -Compress
    }

    exit 0
}
try {
    $resolvedOutputPath = Resolve-OutputFilePath -Path $OutputPath
    Initialize-GraphRuntimeState
    Write-Phase 'Start'
    Write-Step "Starting AI readiness export to '$resolvedOutputPath'..." -Color Yellow
    Write-Reassurance
    Write-PnPAssemblyStartupGuidance
    if ($IncludeSampling) { Write-Step "Sampling is enabled with SampleSize=$SampleSize." -Color Yellow }

    Write-Phase 'Connection Checks'
    $modulesOk = Test-RequiredModules
    if (-not $modulesOk) {
        Write-Warning 'Some optional PowerShell modules are unavailable; continuing with Graph REST and available service connectors.'
    }
    Connect-Services

    Write-Phase 'Permission Validation'
    $requiredScopes = Get-RequiredGraphScopes
    $script:GraphMissingScopes = @()
    $scopeValidationOk = Test-GraphScopeValidation -RequiredScopes $requiredScopes
    if (-not $scopeValidationOk) {
        if ($GraphClientId -eq '04b07795-8ddb-461a-bbee-02f9e1bf7b46') {
            Write-Step 'Using the default Microsoft client ID with .default scope requires pre-consented Graph delegated permissions in this tenant. Missing scopes must be granted by an admin (or use a custom app registration with the required delegated scopes).' -Color DarkYellow
        }
        $null = Confirm-GraphConsent -RequiredScopes $requiredScopes
        Write-Verbose 'Graph scope validation found partial consent in this tenant; some collectors will return not collected instead of failing.'
    }

    Write-Phase 'Data Collection'
    Write-Step 'Starting collector run...' -Color Yellow
    $script:CollectorStepIndex = 0
    $collectorDefinitions = @(
        [ordered]@{ Section = 'Metadata'; Action = { Get-TenantMetadata } }
        [ordered]@{ Section = 'Licensing'; Action = { Get-LicensingReadiness } }
        [ordered]@{ Section = 'Identity'; Action = { Get-IdentityReadiness } }
        [ordered]@{ Section = 'ConditionalAccess'; Action = { Get-ConditionalAccessReadiness } }
        [ordered]@{ Section = 'DataGovernance'; Action = { Get-DataGovernanceReadiness } }
        [ordered]@{ Section = 'SharePointOneDrive'; Action = { Get-SharePointOneDriveReadiness -Sample:$IncludeSampling -Max $SampleSize } }
        [ordered]@{ Section = 'Exchange'; Action = { Get-ExchangeReadiness } }
        [ordered]@{ Section = 'Teams'; Action = { Get-TeamsReadiness } }
        [ordered]@{ Section = 'SearchIndex'; Action = { Get-SearchIndexReadiness } }
        [ordered]@{ Section = 'Apps'; Action = { Get-AppsAndEndpointReadiness } }
        [ordered]@{ Section = 'SecureScore'; Action = { Get-SecureScoreReadiness } }
        [ordered]@{ Section = 'AdoptionSignals'; Action = { Get-AdoptionSignals } }
        [ordered]@{ Section = 'IdentityAccessAdvanced'; Action = { Get-IdentityAccessAdvanced } }
        [ordered]@{ Section = 'DataProtectionAdvanced'; Action = { Get-DataProtectionAdvanced } }
        [ordered]@{ Section = 'SharePointExposureAdvanced'; Action = { Get-SharePointExposureAdvanced } }
        [ordered]@{ Section = 'TeamsAdvanced'; Action = { Get-TeamsAdvanced } }
        [ordered]@{ Section = 'SearchSemanticAdvanced'; Action = { Get-SearchSemanticAdvanced } }
        [ordered]@{ Section = 'EndpointAppAdvanced'; Action = { Get-EndpointAppAdvanced } }
        [ordered]@{ Section = 'AdoptionAdvanced'; Action = { Get-AdoptionAdvanced } }
        [ordered]@{ Section = 'AppGovernanceAdvanced'; Action = { Get-AppGovernanceAdvanced } }
    )

    $script:CollectorStepTotal = @($collectorDefinitions).Count
    foreach ($collector in @($collectorDefinitions)) {
        Invoke-Collector -Section ([string]$collector.Section) -Action $collector.Action
    }

    # ReadinessFlags runs last — it summarises the sections above.
    Write-Phase 'Summary'
    Write-Step 'Generating readiness summary flags...' -Color Cyan
    # Run summary builders through Invoke-Collector so a summary failure never blocks the JSON export.
    $script:CollectorStepTotal = $script:CollectorStepIndex + 4
    Invoke-Collector -Section 'ReadinessFlags' -Action { Get-ReadinessFlags }
    Invoke-Collector -Section 'ReadinessEvidence' -Action { Get-ReadinessEvidence }
    Invoke-Collector -Section 'Recommendations' -Action { Get-Recommendations }
    Invoke-Collector -Section 'PrerequisitesAndGaps' -Action { Get-PrerequisitesAndGaps }

    Write-Phase 'Output'
    Write-Step "Writing JSON export to '$resolvedOutputPath'..." -Color Cyan
    $exportReport = [ordered]@{}
    foreach ($key in @($script:Report.Keys)) {
        if (-not $IncludeDiagnostics -and $key -eq 'CollectorTimings') {
            continue
        }

        $exportReport[$key] = $script:Report[$key]
    }
    $exportReport | ConvertTo-Json -Depth 15 | Out-File -FilePath $resolvedOutputPath -Encoding utf8NoBOM
    Write-Step "AI readiness export complete: $resolvedOutputPath" -Color Green
    $timingRows = @($script:Report.CollectorTimings.GetEnumerator() |
        Sort-Object { $_.Value.DurationMs } -Descending |
        Select-Object -First 5 |
        ForEach-Object {
            [PSCustomObject]@{
                Section    = $_.Key
                Status     = $_.Value.Status
                DurationMs = $_.Value.DurationMs
            }
        })

    if ($timingRows.Count -gt 0) {
        Write-Host ''
        Write-Host 'Top collector durations (ms):' -ForegroundColor Cyan
        $timingRows | Format-Table -AutoSize
    }

    if ($script:Report.ReadinessEvidence) {
        Write-Host ''
        Write-Host ("Overall AI readiness score: {0}" -f (Get-ObjectPropertyValue -InputObject $script:Report.ReadinessEvidence -Name 'OverallScore' -Default 'not collected')) -ForegroundColor Cyan
    }
    $reportTopRisks = @(Get-ObjectPropertyValue -InputObject $script:Report.Recommendations -Name 'TopRisks' -Default @())
    if (@($reportTopRisks).Count -gt 0) {
        Write-Host 'Top risk:' -ForegroundColor Yellow
        Write-Host (" - {0}" -f $reportTopRisks[0]) -ForegroundColor Yellow
    }
    $reportQuickWins = @(Get-ObjectPropertyValue -InputObject $script:Report.Recommendations -Name 'QuickWins' -Default @())
    if (@($reportQuickWins).Count -gt 0) {
        Write-Host 'Top quick win:' -ForegroundColor Green
        Write-Host (" - {0}" -f $reportQuickWins[0]) -ForegroundColor Green
    }

    Write-Host ("Sections collected: {0} | Errors: {1}" -f `
        (($script:Report.Keys | Where-Object { $script:Report[$_] -and $_ -ne 'Errors' }).Count), `
        $script:Report.Errors.Count)
    Write-Step 'Run complete. Reminder: tenant settings were read-only throughout this export.' -Color Green
}
catch {
    Write-Error "Fatal: $($_.Exception.Message)"
}
finally {
    Disconnect-Services
}