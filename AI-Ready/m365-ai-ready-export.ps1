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
    [string[]]$GraphScopes,
    [switch]$IncludeSampling,
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

$script:Report = [ordered]@{
    Metadata           = $null
    Licensing          = $null
    Identity           = $null
    ConditionalAccess  = $null
    DataGovernance     = $null
    SharePointOneDrive = $null
    Exchange           = $null
    Teams              = $null
    SearchIndex        = $null
    Apps               = $null
    AdoptionSignals    = $null
    IdentityAccessAdvanced = $null
    DataProtectionAdvanced = $null
    SharePointExposureAdvanced = $null
    TeamsAdvanced      = $null
    SearchSemanticAdvanced = $null
    EndpointAppAdvanced = $null
    AdoptionAdvanced   = $null
    AppGovernanceAdvanced = $null
    ReadinessFlags     = $null
    ReadinessEvidence  = $null
    Recommendations    = $null
    CollectorTimings   = [ordered]@{}
    Errors             = [System.Collections.Generic.List[object]]::new()
}
$script:GraphAccessToken = $null
$script:GraphAccessTokenExpiresAtUtc = $null
$script:TeamsConnected = $false
$script:GraphMissingScopes = @()
$script:GraphProvidedScopes = @()
$script:GraphRestAllCache = @{}
$script:CollectorStepIndex = 0
$script:CollectorStepTotal = 0
$script:ScoringModel = [ordered]@{
    # Explicit scoring model to avoid hard-coded magic numbers in Get-ReadinessEvidence.
    IdentityAccess = [ordered]@{
        MfaRegisteredPercentThreshold = 80
        MfaRegisteredWeight = 50
        AdvancedMfaPercentThreshold = 80
        AdvancedMfaWeight = 30
        ConditionalAccessPolicyThreshold = 3
        ConditionalAccessWeight = 20
    }
    DataGovernance = [ordered]@{
        LabelsPublishedWeight = 35
        DlpEnforcedWeight = 35
        RetentionEnabledWeight = 30
    }
    ContentExposure = [ordered]@{
        OneDriveCoverageThreshold = 60
        OneDriveCoverageWeight = 40
        NoAnyoneLinksWeight = 30
        AnyoneLinksPresentWeight = 10
        RestrictedSearchWeight = 30
    }
    Adoption = [ordered]@{
        ActiveUserHighThreshold = 60
        ActiveUserMediumThreshold = 30
        ActiveUserHighWeight = 60
        ActiveUserMediumWeight = 40
        ActiveUserLowWeight = 20
        TrendReportedWeight = 40
    }
    EndpointReadiness = [ordered]@{
        ComplianceHighThreshold = 80
        ComplianceMediumThreshold = 60
        HighScore = 100
        MediumScore = 70
        LowScore = 40
    }
    Confidence = [ordered]@{
        HighThreshold = 85
        MediumThreshold = 60
        EvaluatedSections = @(
            'Licensing',
            'Identity',
            'ConditionalAccess',
            'DataGovernance',
            'SharePointOneDrive',
            'Exchange',
            'Teams',
            'SearchIndex',
            'Apps',
            'AdoptionSignals',
            'IdentityAccessAdvanced',
            'DataProtectionAdvanced',
            'SharePointExposureAdvanced',
            'TeamsAdvanced',
            'SearchSemanticAdvanced',
            'EndpointAppAdvanced',
            'AdoptionAdvanced',
            'AppGovernanceAdvanced'
        )
    }
}

function Get-DefaultGraphScopes {
    return @(
        'https://graph.microsoft.com/User.Read.All',
        'https://graph.microsoft.com/Directory.Read.All',
        'https://graph.microsoft.com/Group.Read.All',
        'https://graph.microsoft.com/Organization.Read.All',
        'https://graph.microsoft.com/Policy.Read.All',
        'https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All',
        'https://graph.microsoft.com/Reports.Read.All',
        'https://graph.microsoft.com/AuditLog.Read.All',
        'https://graph.microsoft.com/ExternalConnection.Read.All',
        'https://graph.microsoft.com/InformationProtectionPolicy.Read.All'
    )
}

function Get-RequiredGraphScopes {
    # Baseline read-only permissions required for this script's core collectors.
    return @(
        'https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All',
        'https://graph.microsoft.com/Directory.Read.All',
        'https://graph.microsoft.com/ExternalConnection.Read.All',
        'https://graph.microsoft.com/Group.Read.All',
        'https://graph.microsoft.com/InformationProtectionPolicy.Read.All',
        'https://graph.microsoft.com/Organization.Read.All',
        'https://graph.microsoft.com/Policy.Read.All',
        'https://graph.microsoft.com/Reports.Read.All'
    )
}

function Test-GraphPermissionsAvailable {
    param([string[]]$RequiredPermissions)

    $missing = @($script:GraphMissingScopes)
    if (@($missing).Count -eq 0) {
        return $true
    }

    $requiredNormalized = @($RequiredPermissions |
        ForEach-Object { Normalize-GraphPermissionName -Value $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique)
    $missingNormalized = @($missing |
        ForEach-Object { Normalize-GraphPermissionName -Value $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique)

    return @($requiredNormalized | Where-Object { $_ -in $missingNormalized }).Count -eq 0
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

    return @($InputObject.Keys) -contains $Key
}

function Test-SectionCollected {
    param([AllowNull()][object]$SectionData)

    # Section-level heuristic only: this does not inspect nested fields for partial "not collected" values.
    if ($null -eq $SectionData) { return $false }

    if ($SectionData -is [string]) {
        return -not ([string]$SectionData -like 'not collected*')
    }

    if ($SectionData.PSObject.Properties.Name -contains 'Status') {
        $status = [string]$SectionData.Status
        if ($status -like 'not collected*') { return $false }
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

function Normalize-GraphPermissionName {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }

    $normalized = [string]$Value.Trim()
    $normalized = $normalized -replace '^https://graph\.microsoft\.com/', ''
    $normalized = $normalized -replace '^https://graph\.microsoft\.com$', ''
    $normalized = $normalized.ToLowerInvariant()

    return $normalized
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

    $providedScopes = @($providedScopes | ForEach-Object { Normalize-GraphPermissionName -Value $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    $script:GraphProvidedScopes = @($providedScopes)
    $requiredScopeList = @($RequiredScopes | ForEach-Object { Normalize-GraphPermissionName -Value $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

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

function Ensure-GraphConsent {
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
        ForEach-Object { Normalize-GraphPermissionName -Value $_ } |
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

function Get-GraphReportRows {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Query = @{}
    )

    $token = Get-GraphAccessToken
    if (-not $token) { return @() }

    $requestUri = Add-GraphQueryString -Uri $Uri -Query $Query

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

        if ($trimmed.StartsWith('{') -or $trimmed.StartsWith('[')) {
            try {
                $jsonPayload = $trimmed | ConvertFrom-Json -ErrorAction Stop
                if ($jsonPayload -is [System.Collections.IEnumerable] -and -not ($jsonPayload -is [string])) {
                    return @($jsonPayload)
                }

                return @($jsonPayload)
            }
            catch {
                return @()
            }
        }

        return @($content | ConvertFrom-Csv)
    }
    catch {
        return @()
    }
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

        return [ordered]@{
            Status      = 'reported'
            Reason      = ''
            ActiveUsers = if ($null -ne $activeUsers) { [int]$activeUsers } else { 'not collected' }
            ReportDate  = if ($reportDate) { [string]$reportDate } else { 'not collected' }
            RecordCount = @($rows).Count
            Source      = if ($reportUri -match 'UserDetail') { 'microsoft365CopilotUsageUserDetail' } else { 'microsoft365CopilotUsageUserCounts' }
        }
    }

    return [ordered]@{
        Status      = 'not collected'
        Reason      = 'Copilot usage report endpoint returned no data (not enabled, unsupported, or not licensed in this tenant).'
        ActiveUsers = 'not collected'
        ReportDate  = 'not collected'
        RecordCount = 0
        Source      = 'not collected'
    }
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

    try {
        return Invoke-RestMethod @irmParams
    }
    catch {
        $message = $_.Exception.Message
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

function Get-GraphRestAll {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Query = @{},
        [switch]$BypassCache
    )

    $requestUri = Add-GraphQueryString -Uri $Uri -Query $Query
    if (-not $BypassCache -and $script:GraphRestAllCache.ContainsKey($requestUri)) {
        return @($script:GraphRestAllCache[$requestUri])
    }

    $items = New-Object System.Collections.Generic.List[object]
    $nextUri = $requestUri

    while ($nextUri) {
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

# ----------------------------------------------------------------------------
# Connections (read-only scopes)
# ----------------------------------------------------------------------------
function Connect-Services {
    Write-Step 'Connecting to Microsoft Graph and Microsoft 365 services...' -Color Yellow

    try { Import-Module PnP.PowerShell -ErrorAction SilentlyContinue } catch {}
    try { Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue } catch {}

    try {
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
            $isDefaultMsClient = ($GraphClientId -eq '04b07795-8ddb-461a-bbee-02f9e1bf7b46')
            $graphScopes = if ($GraphScopes -and $GraphScopes.Count -gt 0) {
                @($GraphScopes)
            }
            elseif ($isDefaultMsClient) {
                @('https://graph.microsoft.com/.default')
            }
            else {
                Get-DefaultGraphScopes
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
                Write-Step 'Requesting the required delegated Graph permissions so the tenant can consent to them if needed.' -Color DarkYellow
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
        if (Get-Command Connect-MicrosoftTeams -ErrorAction SilentlyContinue) {
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
        else {
            Write-Step 'Microsoft Teams cmdlets are not available in this session.' -Color DarkYellow
        }
    }
    catch {
        $script:TeamsConnected = $false
        Write-Warning "Teams connection check failed: $($_.Exception.Message)"
    }

    try {
        $pnpConnected = $false
        if (Get-Command Get-PnPConnection -ErrorAction SilentlyContinue) {
            $pnpConnected = $null -ne (Get-PnPConnection -ErrorAction SilentlyContinue)
        }

        if (-not $pnpConnected -and (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) {
            Write-Step 'Connecting to SharePoint / PnP for site metadata...' -Color Cyan
            $org = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/organization')) | Select-Object -First 1
            $initialDomain = @($org.VerifiedDomains | Where-Object { $_.Name -match '\.onmicrosoft\.com$' } | Select-Object -ExpandProperty Name -First 1)
            $tenantSlug = if ($initialDomain) { $initialDomain -replace '\.onmicrosoft\.com$', '' } else { $null }
            $adminUrl = if ($tenantSlug) {
                "https://$($tenantSlug.ToLowerInvariant())-admin.sharepoint.com"
            }
            else { $null }
            if ($adminUrl -and $adminUrl -match '^https://.+') {
                Connect-PnPOnline -Url $adminUrl -Interactive -ErrorAction Stop -WarningAction SilentlyContinue -InformationAction SilentlyContinue 3>$null 6>$null
                $pnpConnected = $true
                Write-Step 'Connected to SharePoint / PnP.' -Color Green
            }
        }
        elseif ($pnpConnected) {
            Write-Step 'Reusing the existing SharePoint / PnP session.' -Color Green
        }
        else {
            Write-Step 'SharePoint / PnP cmdlets are not available in this session.' -Color DarkYellow
        }
    }
    catch {
        Write-Verbose "SharePoint / PnP connection skipped or unavailable: $($_.Exception.Message)"
    }

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
    $org = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/organization')) | Select-Object -First 1
    $users = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/users' -Query @{ '$select' = 'id,userPrincipalName,userType,accountEnabled' }))
    $userCount = @($users).Count
    $memberCount = @($users | Where-Object { $_.UserType -ne 'Guest' }).Count
    $guestCount = @($users | Where-Object { $_.UserType -eq 'Guest' }).Count
    $defaultDomain = @($org.VerifiedDomains | Where-Object { $_.IsDefault -eq $true } | Select-Object -ExpandProperty Name -First 1)

    return [ordered]@{
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
        TotalUsers         = $userCount
        MemberUsers        = $memberCount
        GuestUsers         = $guestCount
        IncludeSampling    = [bool]$IncludeSampling
        IncludeDetailedData = [bool]$IncludeDetailedData
        SampleSize         = [int]$SampleSize
    }
}

function Get-LicensingReadiness {
    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('Organization.Read.All'))) {
        return [ordered]@{
            Status = 'not collected'
            Reason = 'Missing Graph permission: Organization.Read.All'
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
            ([string]$_.SkuPartNumber).ToUpperInvariant() -in $knownCopilotSkusUpper
        } |
        Select-Object -ExpandProperty SkuPartNumber -Unique)

    $qualifying = @($skus | Where-Object { $_.SkuPartNumber -match 'ENTERPRISEPACK|STANDARDPACK|M365BUSINESSSTANDARD|M365BUSINESSPREMIUM|E3|E5' } | Select-Object -ExpandProperty SkuPartNumber)

    $users = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/users' -Query @{ '$select' = 'id,assignedPlans' }))
    $ready = 0
    foreach ($u in $users) {
        $plans = @($u.AssignedPlans)
        $servicePlanNames = @($plans |
            Where-Object {
                ($_.PSObject.Properties.Name -contains 'ProvisioningStatus') -and
                ([string]$_.ProvisioningStatus -eq 'Success')
            } |
            ForEach-Object {
                if ($_.PSObject.Properties.Name -contains 'ServicePlanName') {
                    [string]$_.ServicePlanName
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

        $hasApps = @($servicePlanNames | Where-Object {
            $_ -in @('OFFICESUBSCRIPTION', 'OFFICE_BUSINESS', 'MICROSOFTOFFICE', 'M365APPS')
        }).Count -gt 0
        $hasExchange = $servicePlanNames -match 'EXCHANGE|EXCHANGESTANDARD|EXCHANGEENTERPRISE'
        $hasSharePoint = $servicePlanNames -match 'SHAREPOINT|SHAREPOINTENTERPRISE|SPB|SPO'
        $hasTeams = $servicePlanNames -match 'TEAMS|TEAMSPACEAPI'

        if ($hasApps -and $hasExchange.Count -gt 0 -and $hasSharePoint.Count -gt 0 -and $hasTeams.Count -gt 0) {
            $ready++
        }
    }

    return [ordered]@{
        Skus                  = $skuList
        CopilotSku            = if (@($copilotSkuMatches).Count -gt 0) { @($copilotSkuMatches) } else { 'not collected' }
        QualifyingBaseLicenses = @($qualifying)
        PrereqReadyUserCount  = $ready
        PrereqReadyPercent   = if (@($users).Count -gt 0) { [math]::Round(($ready / @($users).Count) * 100, 2) } else { 0 }
    }
}

function Get-IdentityReadiness {
    $auth = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails'))
    $featureSummary = Get-UserRegistrationFeatureSummary -IncludedUserTypes 'all' -IncludedUserRoles 'all'
    $mfaCapable = @($auth | Where-Object { $_.IsMfaCapable -eq $true }).Count
    $mfaRegistered = @($auth | Where-Object { $_.IsMfaRegistered -eq $true }).Count
    $passwordless = @($auth | Where-Object { $_.IsPasswordlessCapable -eq $true }).Count

    $mfaPopulationUserCount = 'not collected'
    $mfaPopulationSource = 'not collected'
    if ($featureSummary -and $featureSummary.PSObject.Properties.Name -contains 'totalUserCount') {
        $mfaPopulationUserCount = Get-IntValue -Value $featureSummary.totalUserCount -Default 0
        $mfaPopulationSource = 'reports/authenticationMethods/usersRegisteredByFeature totalUserCount'
    }
    elseif (@($auth).Count -gt 0) {
        $mfaPopulationUserCount = @($auth).Count
        $mfaPopulationSource = 'reports/authenticationMethods/userRegistrationDetails row count fallback'
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

    $users = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/users' -Query @{ '$select' = 'id,userType,signInActivity' }))
    $guestUsers = @($users | Where-Object { $_.UserType -eq 'Guest' }).Count
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

    return [ordered]@{
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
    $enabled = @($policies | Where-Object { $_.State -eq 'enabled' }).Count

    $requireMfa = @($policies | Where-Object {
        $_.State -eq 'enabled' -and
        $_.GrantControls.BuiltInControls -contains 'mfa'
    }).Count -gt 0

    $blockLegacyAuth = @($policies | Where-Object {
        $_.State -eq 'enabled' -and
        $_.GrantControls.BuiltInControls -contains 'block' -and
        (
            ($_.Conditions.ClientAppTypes -contains 'exchangeActiveSync' -or
             $_.Conditions.ClientAppTypes -contains 'other') -or
            # Some newer CA policies express this through authenticationFlows instead of clientAppTypes.
            ($_.Conditions.AuthenticationFlows -and
             -not [string]::IsNullOrWhiteSpace($_.Conditions.AuthenticationFlows.TransferMethods))
        )
    }).Count -gt 0

    $deviceCompliance = @($policies | Where-Object {
        $_.State -eq 'enabled' -and
        $_.GrantControls.BuiltInControls -contains 'compliantDevice'
    }).Count -gt 0

    $signInRisk = @($policies | Where-Object {
        $_.State -eq 'enabled' -and
        $_.Conditions.SignInRiskLevels -and $_.Conditions.SignInRiskLevels.Count -gt 0
    }).Count -gt 0

    $hasPoliciesWithExclusions = @($policies | Where-Object {
        $_.State -eq 'enabled' -and
        ($_.Conditions.Users.ExcludeTargets -or $_.Conditions.Applications.ExcludeApplications)
    }).Count -gt 0

    return [ordered]@{
        TotalPolicies      = @($policies).Count
        EnabledPolicies    = $enabled
        CoverageFlags      = [ordered]@{
            RequireMfa         = $requireMfa
            BlockLegacyAuth    = $blockLegacyAuth
            DeviceCompliance   = $deviceCompliance
            SignInRisk         = $signInRisk
            HasPoliciesWithExclusions = $hasPoliciesWithExclusions
        }
    }
}

function Get-DataGovernanceReadiness {
    $labels = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/informationProtection/policy/labels'))
    $published = @($labels | Where-Object { $_.IsPublished -eq $true }).Count
    $autoLabel = @($labels | Where-Object { $_.IsLabelAutoClassified -eq $true -or $_.IsDefault -eq $true }).Count

    $dlp = if (Get-Command Get-DlpCompliancePolicy -ErrorAction SilentlyContinue) { Get-DlpCompliancePolicy -ErrorAction SilentlyContinue | Select-Object Name, Mode, State } else { 'not collected' }
    $retention = if (Get-Command Get-RetentionCompliancePolicy -ErrorAction SilentlyContinue) { Get-RetentionCompliancePolicy -ErrorAction SilentlyContinue | Select-Object Name, Enabled } else { 'not collected' }

    return [ordered]@{
        SensitivityLabelsPublished = $published
        SensitivityLabelsTotal     = @($labels).Count
        AutoLabelingIndicators     = $autoLabel
        DlpPolicies                = $dlp
        RetentionPolicies          = $retention
    }
}

function Get-SharePointOneDriveReadiness {
    param([switch]$Sample, [int]$Max = 200)

    $tenant = $null
    $sites = @()

    try {
        if (Get-Command Get-PnPTenant -ErrorAction SilentlyContinue) { $tenant = Get-PnPTenant -ErrorAction Stop }
    } catch {}
    try {
        if (Get-Command Get-PnPTenantSite -ErrorAction SilentlyContinue) { $sites = @(Get-PnPTenantSite -IncludeOneDriveSites -ErrorAction Stop) }
    } catch {}
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
        ExternalSharingLevel   = if ($tenant) { $tenant.SharingCapability } else { 'not collected' }
        DefaultSharingLinkType = if ($tenant) { $tenant.DefaultSharingLinkType } else { 'not collected' }
        RestrictedSearchEnabled = if ($tenant) { [bool]$tenant.RestrictedSharePointSearch } else { 'not collected' }
        TotalSiteCount         = $totalCount
        Truncated              = [bool]($Sample -and $totalCount -gt $Max)
        Sites                  = @($siteSummaries)
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
            $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue)
            $shared = @($mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' }).Count
            $archive = @($mailboxes | Where-Object { $_.ArchiveStatus -eq 'Active' -or $_.IsArchiveEnabled -eq $true }).Count
            $litigation = @($mailboxes | Where-Object { $_.LitigationHoldEnabled -eq $true }).Count

            return [ordered]@{
                TotalMailboxes       = @($mailboxes).Count
                UserMailboxes        = @($mailboxes | Where-Object { $_.RecipientTypeDetails -in @('UserMailbox','SharedMailbox') }).Count
                SharedMailboxes      = $shared
                ArchiveEnabledCount  = $archive
                LitigationHoldCount  = $litigation
                ExchangeOnlineReady  = @($mailboxes).Count -gt 0
            }
        }

        $users = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/users' -Query @{ '$select' = 'id,userPrincipalName,mail' }))
        $mailboxCount = @($users | Where-Object { -not [string]::IsNullOrWhiteSpace($_.mail) }).Count

        return [ordered]@{
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
            TotalMailboxes       = 0
            UserMailboxes        = 0
            SharedMailboxes      = 'not collected'
            ArchiveEnabledCount  = 'not collected'
            LitigationHoldCount  = 'not collected'
            ExchangeOnlineReady  = $false
        }
    }
}

function Get-TeamsReadiness {
    $teams = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/groups' -Query @{ '$filter' = "resourceProvisioningOptions/any(x:x eq 'Team')"; '$select' = 'id,displayName,visibility,renewedDateTime' }))
    $teamCount = @($teams).Count
    $privateTeams = @($teams | Where-Object { $_.Visibility -eq 'Private' }).Count
    $publicTeams = @($teams | Where-Object { $_.Visibility -eq 'Public' }).Count
    $staleTeams = @($teams | Where-Object {
        $_.PSObject.Properties.Name -contains 'RenewedDateTime' -and
        $_.RenewedDateTime -and
        ((Get-Date).ToUniversalTime() - [DateTime]$_.RenewedDateTime) -gt [TimeSpan]::FromDays(180)
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

    return [ordered]@{
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
    $connections = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/external/connections'))
    $readyConnectors = @($connections | Where-Object { $_.PSObject.Properties.Name -contains 'state' -and $_.state -match 'ready|active' }).Count

    return [ordered]@{
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
    $m365Apps = @()
    try {
        $m365Apps = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/reports/m365AppUserCounts'))
    }
    catch {
        $m365Apps = @()
    }

    return [ordered]@{
        M365AppsUpdateChannel = if (@($m365Apps).Count -gt 0) { 'reported' } else { 'not collected' }
        VersionReadiness      = if (@($m365Apps).Count -gt 0) { 'reported' } else { 'not collected' }
        BrowserSignals        = 'not collected'
        EndpointSignals       = 'not collected'
        ReportCount           = @($m365Apps).Count
    }
}

function Get-AdoptionSignals {
    $rows = @()
    try {
        $rows = @(Get-GraphRestAll -Uri "https://graph.microsoft.com/beta/reports/getOffice365ActiveUserCounts(period='D30')")
    }
    catch {
        $rows = @()
    }

    $copilot = Get-CopilotUsageSnapshot -Period 'D30'

    if ($rows.Count -eq 0) {
        return [ordered]@{
            ActiveUsers30Days      = 0
            TeamsActiveUsers       = 0
            SharePointActiveUsers  = 0
            OneDriveActiveUsers    = 0
            OutlookActiveUsers     = 0
            ReportSource           = 'not collected'
            CopilotActiveUsers30Days = $copilot.ActiveUsers
            CopilotUsageStatus      = $copilot.Status
            CopilotReportSource     = $copilot.Source
            CopilotReportDate       = $copilot.ReportDate
            CopilotCollectionReason = if ([string]::IsNullOrWhiteSpace($copilot.Reason)) { 'reported' } else { $copilot.Reason }
        }
    }

    $latest = $rows |
        Sort-Object { try { [datetime]$_.reportDate } catch { [datetime]::MinValue } } -Descending |
        Select-Object -First 1

    return [ordered]@{
        # 'office365' is the aggregate active-user count across all Microsoft 365 workloads for the period.
        ActiveUsers30Days      = [int]($latest.office365  ?? 0)
        TeamsActiveUsers       = [int]($latest.teams     ?? 0)
        SharePointActiveUsers  = [int]($latest.sharePoint ?? 0)
        OneDriveActiveUsers    = [int]($latest.oneDrive   ?? 0)
        OutlookActiveUsers     = [int]($latest.exchange  ?? 0)
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

    $mfaCapable = if ($identity -and $identity.PSObject.Properties.Name -contains 'MfaCapableCount') { Get-IntValue -Value $identity.MfaCapableCount -Default 0 } else { 0 }
    $mfaRegistered = if ($identity -and $identity.PSObject.Properties.Name -contains 'MfaRegisteredCount') { Get-IntValue -Value $identity.MfaRegisteredCount -Default 0 } else { 0 }
    $passwordless = if ($identity -and $identity.PSObject.Properties.Name -contains 'PasswordlessCount') { Get-IntValue -Value $identity.PasswordlessCount -Default 0 } else { 0 }
    $globalAdmins = if ($identity -and $identity.PSObject.Properties.Name -contains 'GlobalAdminCount') { Get-IntValue -Value $identity.GlobalAdminCount -Default 0 } else { 0 }
    $privilegedRoles = if ($identity -and $identity.PSObject.Properties.Name -contains 'PrivilegedRoleCount') { Get-IntValue -Value $identity.PrivilegedRoleCount -Default 0 } else { 0 }
    $adminRoleCount = $globalAdmins + $privilegedRoles
    $mfaPopulationUserCount = if ($identity -and $identity.PSObject.Properties.Name -contains 'MfaPopulationUserCount') { Get-IntValue -Value $identity.MfaPopulationUserCount -Default 0 } else { 0 }
    $fallbackTotalUsers = if ($metadata -and $metadata.PSObject.Properties.Name -contains 'TotalUsers') { Get-IntValue -Value $metadata.TotalUsers -Default 0 } else { 0 }
    $mfaDenominator = if ($mfaPopulationUserCount -gt 0) { $mfaPopulationUserCount } else { $fallbackTotalUsers }

    return [ordered]@{
        AdminRoleObjectsCount = $adminRoleCount
        MfaCapableUsers       = $mfaCapable
        MfaRegisteredUsers    = $mfaRegistered
        PasswordlessCapableUsers = $passwordless
        MfaPopulationUserCount = if ($mfaDenominator -gt 0) { $mfaDenominator } else { 'not collected' }
        MfaPopulationSource  = if ($identity -and $identity.PSObject.Properties.Name -contains 'MfaPopulationSource') { $identity.MfaPopulationSource } else { 'not collected' }
        MfaRegistrationPercent = if ($mfaDenominator -gt 0) { [math]::Round(($mfaRegistered / $mfaDenominator) * 100, 2) } else { 0 }
        PhishingResistantMfaCoverage = 'not collected'
        PimSignal = 'not collected'
    }
}

function Get-DataProtectionAdvanced {
    $governance = $script:Report.DataGovernance
    $published = if ($governance -and $governance.PSObject.Properties.Name -contains 'SensitivityLabelsPublished') { Get-IntValue -Value $governance.SensitivityLabelsPublished -Default 0 } else { 0 }
    $labelsTotal = if ($governance -and $governance.PSObject.Properties.Name -contains 'SensitivityLabelsTotal') { Get-IntValue -Value $governance.SensitivityLabelsTotal -Default 0 } else { 0 }

    $dlpPolicies = @()
    if ($governance -and $governance.PSObject.Properties.Name -contains 'DlpPolicies' -and $governance.DlpPolicies -isnot [string]) {
        $dlpPolicies = @($governance.DlpPolicies)
    }
    $dlpEnforced = @($dlpPolicies | Where-Object { $_.PSObject.Properties.Name -contains 'Mode' -and $_.Mode -match 'Enforce|Enable' }).Count

    $retentionPolicies = @()
    if ($governance -and $governance.PSObject.Properties.Name -contains 'RetentionPolicies' -and $governance.RetentionPolicies -isnot [string]) {
        $retentionPolicies = @($governance.RetentionPolicies)
    }
    $retentionEnabled = @($retentionPolicies | Where-Object { $_.PSObject.Properties.Name -contains 'Enabled' -and $_.Enabled -eq $true }).Count

    return [ordered]@{
        LabelPoliciesPublished = $published
        LabelPoliciesTotal     = $labelsTotal
        DlpPoliciesEnforced    = $dlpEnforced
        DlpPoliciesTotal       = @($dlpPolicies).Count
        RetentionPoliciesEnabled = $retentionEnabled
        RetentionPoliciesTotal = @($retentionPolicies).Count
        eDiscoverySignal       = 'not collected'
        InsiderRiskSignal      = 'not collected'
    }
}

function Get-SharePointExposureAdvanced {
    $sharePoint = $script:Report.SharePointOneDrive
    $sampleSites = @()
    if ($sharePoint -and $sharePoint.PSObject.Properties.Name -contains 'Sites' -and $sharePoint.Sites -isnot [string]) {
        $sampleSites = @($sharePoint.Sites)
    }

    $externalSharingSites = @($sampleSites | Where-Object { $_.PSObject.Properties.Name -contains 'SharingCapability' -and $_.SharingCapability -in @('ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'AnonymousAccess') }).Count
    $anyoneLinkSites = @($sampleSites | Where-Object { $_.PSObject.Properties.Name -contains 'SharingCapability' -and $_.SharingCapability -eq 'AnonymousAccess' }).Count
    $inactiveSites = @($sampleSites | Where-Object {
        $_.PSObject.Properties.Name -contains 'LastContentModifiedDate' -and
        $_.LastContentModifiedDate -and
        ((Get-Date).ToUniversalTime() - [DateTime]$_.LastContentModifiedDate) -gt [TimeSpan]::FromDays(180)
    }).Count

    $oversharing = $null
    if ($sharePoint -and $sharePoint.PSObject.Properties.Name -contains 'OversharingSignals') {
        $oversharing = $sharePoint.OversharingSignals
    }

    $oversharingStatus = if ($oversharing -and $oversharing.PSObject.Properties.Name -contains 'Status') { [string]$oversharing.Status } else { 'not collected' }
    $sampledAnyoneLinks = if ($oversharing -and $oversharing.PSObject.Properties.Name -contains 'SampledAnyoneLinkCount') { $oversharing.SampledAnyoneLinkCount } else { 'not collected' }
    $sampledOrgWideLinks = if ($oversharing -and $oversharing.PSObject.Properties.Name -contains 'SampledOrgWideLinkCount') { $oversharing.SampledOrgWideLinkCount } else { 'not collected' }

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
        SampledSites = @($sampleSites).Count
        ExternalSharingSites = $externalSharingSites
        AnyoneLinkSites = $anyoneLinkSites
        InactiveSites = $inactiveSites
        OversharedContentSignal = $oversharedContentSignal
        SampledAnyoneLinkCount = $sampledAnyoneLinks
        SampledOrgWideLinkCount = $sampledOrgWideLinks
        UnlabeledContentSignal = 'not collected'
        Truncated = if ($sharePoint -and $sharePoint.PSObject.Properties.Name -contains 'Truncated') { [bool]$sharePoint.Truncated } else { $false }
    }
}

function Get-TeamsAdvanced {
    $teams = $script:Report.Teams
    $totalTeams = if ($teams -and $teams.PSObject.Properties.Name -contains 'TeamCount') { Get-IntValue -Value $teams.TeamCount -Default 0 } else { 0 }
    $publicTeams = if ($teams -and $teams.PSObject.Properties.Name -contains 'PublicTeamCount') { Get-IntValue -Value $teams.PublicTeamCount -Default 0 } else { 'not collected' }
    $staleTeams = if ($teams -and $teams.PSObject.Properties.Name -contains 'StaleTeamCount') { Get-IntValue -Value $teams.StaleTeamCount -Default 0 } else { 'not collected' }

    return [ordered]@{
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
    $connectorCount = if ($searchIndex -and $searchIndex.PSObject.Properties.Name -contains 'ExternalConnectionsCount') { Get-IntValue -Value $searchIndex.ExternalConnectionsCount -Default 0 } else { 0 }
    $readyConnectors = if ($searchIndex -and $searchIndex.PSObject.Properties.Name -contains 'ReadyExternalConnectionsCount') { Get-IntValue -Value $searchIndex.ReadyExternalConnectionsCount -Default 0 } else { 0 }

    return [ordered]@{
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

    $appsReport = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/reports/m365AppUserCounts'))

    return [ordered]@{
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
    $d7 = @((Get-GraphRestAll -Uri "https://graph.microsoft.com/beta/reports/getOffice365ActiveUserCounts(period='D7')"))
    $d30 = @((Get-GraphRestAll -Uri "https://graph.microsoft.com/beta/reports/getOffice365ActiveUserCounts(period='D30')"))
    $d90 = @((Get-GraphRestAll -Uri "https://graph.microsoft.com/beta/reports/getOffice365ActiveUserCounts(period='D90')"))
    $copilotD7 = Get-CopilotUsageSnapshot -Period 'D7'
    $copilotD30 = Get-CopilotUsageSnapshot -Period 'D30'
    $copilotD90 = Get-CopilotUsageSnapshot -Period 'D90'

    $latestD7 = @($d7 | Sort-Object { try { [datetime]$_.reportDate } catch { [datetime]::MinValue } } -Descending | Select-Object -First 1)
    $latestD30 = @($d30 | Sort-Object { try { [datetime]$_.reportDate } catch { [datetime]::MinValue } } -Descending | Select-Object -First 1)
    $latestD90 = @($d90 | Sort-Object { try { [datetime]$_.reportDate } catch { [datetime]::MinValue } } -Descending | Select-Object -First 1)

    return [ordered]@{
        ActiveUsersD7  = if ($latestD7) { [int]($latestD7[0].office365 ?? 0) } else { 0 }
        ActiveUsersD30 = if ($latestD30) { [int]($latestD30[0].office365 ?? 0) } else { 0 }
        ActiveUsersD90 = if ($latestD90) { [int]($latestD90[0].office365 ?? 0) } else { 0 }
        TrendSignal = if ($latestD7 -and $latestD30 -and $latestD90) { 'reported' } else { 'partial' }
        CopilotActiveUsersD7 = $copilotD7.ActiveUsers
        CopilotActiveUsersD30 = $copilotD30.ActiveUsers
        CopilotActiveUsersD90 = $copilotD90.ActiveUsers
        CopilotTrendSignal = if ($copilotD7.Status -eq 'reported' -and $copilotD30.Status -eq 'reported' -and $copilotD90.Status -eq 'reported') { 'reported' } elseif ($copilotD7.Status -eq 'reported' -or $copilotD30.Status -eq 'reported' -or $copilotD90.Status -eq 'reported') { 'partial' } else { 'not collected' }
        CopilotReportSource = if ($copilotD30.Status -eq 'reported') { $copilotD30.Source } elseif ($copilotD7.Status -eq 'reported') { $copilotD7.Source } else { $copilotD90.Source }
        CopilotCollectionReason = if ($copilotD30.Status -eq 'not collected' -and -not [string]::IsNullOrWhiteSpace($copilotD30.Reason)) { $copilotD30.Reason } elseif ($copilotD7.Status -eq 'not collected' -and -not [string]::IsNullOrWhiteSpace($copilotD7.Reason)) { $copilotD7.Reason } elseif ($copilotD90.Status -eq 'not collected' -and -not [string]::IsNullOrWhiteSpace($copilotD90.Reason)) { $copilotD90.Reason } else { 'reported' }
    }
}

function Get-AppGovernanceAdvanced {
    $servicePrincipals = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Query @{ '$select' = 'id,appId,displayName,accountEnabled' }))
    $oauthGrants = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/oauth2PermissionGrants'))

    $highScopeGrants = @($oauthGrants | Where-Object {
        $_.PSObject.Properties.Name -contains 'scope' -and
        [string]$_.scope -match 'Mail\.ReadWrite|Files\.ReadWrite\.All|Sites\.FullControl\.All|Directory\.ReadWrite\.All'
    }).Count

    $graphResourceSp = @((Get-GraphRestAll -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'" -Query @{ '$select' = 'id,appId,displayName,appRoles' }) | Select-Object -First 1)
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

    $licPercent = if (Test-HashtableKey -InputObject $lic -Key 'PrereqReadyPercent') { Get-DoubleValue -Value $lic['PrereqReadyPercent'] -Default 0 } else { 0 }
    $totalUsers = if (Test-HashtableKey -InputObject $script:Report.Metadata -Key 'TotalUsers') { Get-IntValue -Value $script:Report.Metadata['TotalUsers'] -Default 0 } else { 0 }
    $mfaRegisteredCount = if (Test-HashtableKey -InputObject $id -Key 'MfaRegisteredCount') { Get-IntValue -Value $id['MfaRegisteredCount'] -Default 0 } else { 0 }
    $mfaPopulationUserCount = if (Test-HashtableKey -InputObject $id -Key 'MfaPopulationUserCount') { Get-IntValue -Value $id['MfaPopulationUserCount'] -Default 0 } else { 0 }
    $activeUsers30Days = if (Test-HashtableKey -InputObject $ad -Key 'ActiveUsers30Days') { Get-IntValue -Value $ad['ActiveUsers30Days'] -Default 0 } else { 0 }
    $mfaDenominator = if ($mfaPopulationUserCount -gt 0) { $mfaPopulationUserCount } else { $totalUsers }
    $mfaPercent = if ($mfaDenominator -gt 0) { [math]::Round(($mfaRegisteredCount / $mfaDenominator) * 100, 2) } else { 0 }
    # NOTE: ActiveUsers30Days is a 30-day cumulative count from the Microsoft 365 usage report.
    # It is not a unique single-day user population, so the percentage can exceed 100% in tenants
    # where the same user appears on multiple days during the reporting window.
    $activeUserPercent = if ($totalUsers -gt 0 -and $activeUsers30Days -gt 0) { [math]::Round(($activeUsers30Days / $totalUsers) * 100, 2) } else { 0 }

    $caEnabledPolicies = if (Test-HashtableKey -InputObject $ca -Key 'EnabledPolicies') { Get-IntValue -Value $ca['EnabledPolicies'] -Default 0 } else { 0 }
    $govPublished = if (Test-HashtableKey -InputObject $gov -Key 'SensitivityLabelsPublished') { Get-IntValue -Value $gov['SensitivityLabelsPublished'] -Default 0 } else { 0 }
    $spoRestricted = if (Test-HashtableKey -InputObject $spo -Key 'RestrictedSearchEnabled') { $spo['RestrictedSearchEnabled'] } else { 'not collected' }
    $spoOneDrivePercent = if (Test-HashtableKey -InputObject $spo -Key 'OneDriveCoveragePercent') { Get-DoubleValue -Value $spo['OneDriveCoveragePercent'] -Default 0 } else { 0 }
    $spoExternalSharing = if (Test-HashtableKey -InputObject $spo -Key 'ExternalSharingLevel') { [string]$spo['ExternalSharingLevel'] } else { 'not collected' }

    return [ordered]@{
        LicencePrereqPercent      = $licPercent
        MfaRegisteredPercent      = $mfaPercent
        MfaPopulationUserCount    = if ($mfaDenominator -gt 0) { $mfaDenominator } else { 'not collected' }
        MfaPopulationSource       = if (Test-HashtableKey -InputObject $id -Key 'MfaPopulationSource') { $id['MfaPopulationSource'] } else { 'not collected' }
        CaPoliciesEnabled         = $caEnabledPolicies
        SensitivityLabelsPublished = if ($govPublished -gt 0) { $true } else { $false }
        RestrictedSearchEnabled   = if ($spoRestricted -ne 'not collected') { [bool]$spoRestricted } else { $false }
        OneDriveCoveragePercent   = $spoOneDrivePercent
        ExternalSharingLevel      = $spoExternalSharing
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
    $scoring = $script:ScoringModel

    $mfaRegisteredPercent = if ($flags -and $flags.PSObject.Properties.Name -contains 'MfaRegisteredPercent') { Get-DoubleValue -Value $flags.MfaRegisteredPercent -Default 0 } else { 0 }
    $caEnabledPolicies = if ($flags -and $flags.PSObject.Properties.Name -contains 'CaPoliciesEnabled') { Get-IntValue -Value $flags.CaPoliciesEnabled -Default 0 } else { 0 }
    $identityAdvancedMfaPercent = if ($identity -and $identity.PSObject.Properties.Name -contains 'MfaRegistrationPercent') { Get-DoubleValue -Value $identity.MfaRegistrationPercent -Default 0 } else { 0 }

    $identityScore = 0
    if ($mfaRegisteredPercent -ge [double]$scoring.IdentityAccess.MfaRegisteredPercentThreshold) { $identityScore += [int]$scoring.IdentityAccess.MfaRegisteredWeight }
    if ($identityAdvancedMfaPercent -ge [double]$scoring.IdentityAccess.AdvancedMfaPercentThreshold) { $identityScore += [int]$scoring.IdentityAccess.AdvancedMfaWeight }
    if ($caEnabledPolicies -ge [int]$scoring.IdentityAccess.ConditionalAccessPolicyThreshold) { $identityScore += [int]$scoring.IdentityAccess.ConditionalAccessWeight }

    $governanceScore = 0
    $labelsPublished = if ($flags -and $flags.PSObject.Properties.Name -contains 'SensitivityLabelsPublished') { [bool]$flags.SensitivityLabelsPublished } else { $false }
    $dlpEnforced = if ($gov -and $gov.PSObject.Properties.Name -contains 'DlpPoliciesEnforced') { Get-IntValue -Value $gov.DlpPoliciesEnforced -Default 0 } else { 0 }
    $retentionEnabled = if ($gov -and $gov.PSObject.Properties.Name -contains 'RetentionPoliciesEnabled') { Get-IntValue -Value $gov.RetentionPoliciesEnabled -Default 0 } else { 0 }
    if ($labelsPublished) { $governanceScore += [int]$scoring.DataGovernance.LabelsPublishedWeight }
    if ($dlpEnforced -gt 0) { $governanceScore += [int]$scoring.DataGovernance.DlpEnforcedWeight }
    if ($retentionEnabled -gt 0) { $governanceScore += [int]$scoring.DataGovernance.RetentionEnabledWeight }

    $contentScore = 0
    $oneDriveCoveragePercent = if ($flags -and $flags.PSObject.Properties.Name -contains 'OneDriveCoveragePercent') { Get-DoubleValue -Value $flags.OneDriveCoveragePercent -Default 0 } else { 0 }
    $anyoneLinkSites = if ($spo -and $spo.PSObject.Properties.Name -contains 'AnyoneLinkSites') { Get-IntValue -Value $spo.AnyoneLinkSites -Default 0 } else { 0 }
    $restrictedSearchEnabled = if ($flags -and $flags.PSObject.Properties.Name -contains 'RestrictedSearchEnabled') { [bool]$flags.RestrictedSearchEnabled } else { $false }
    if ($oneDriveCoveragePercent -ge [double]$scoring.ContentExposure.OneDriveCoverageThreshold) { $contentScore += [int]$scoring.ContentExposure.OneDriveCoverageWeight }
    if ($anyoneLinkSites -eq 0) { $contentScore += [int]$scoring.ContentExposure.NoAnyoneLinksWeight } else { $contentScore += [int]$scoring.ContentExposure.AnyoneLinksPresentWeight }
    if ($restrictedSearchEnabled) { $contentScore += [int]$scoring.ContentExposure.RestrictedSearchWeight }

    $adoptionScore = 0
    $activeUserPercent = if ($flags -and $flags.PSObject.Properties.Name -contains 'ActiveUserPercentByWorkload') { Get-DoubleValue -Value $flags.ActiveUserPercentByWorkload -Default 0 } else { 0 }
    $adoptionTrendSignal = if ($script:Report.AdoptionAdvanced -and $script:Report.AdoptionAdvanced.PSObject.Properties.Name -contains 'TrendSignal') { [string]$script:Report.AdoptionAdvanced.TrendSignal } else { 'partial' }
    if ($activeUserPercent -ge [double]$scoring.Adoption.ActiveUserHighThreshold) { $adoptionScore += [int]$scoring.Adoption.ActiveUserHighWeight }
    elseif ($activeUserPercent -ge [double]$scoring.Adoption.ActiveUserMediumThreshold) { $adoptionScore += [int]$scoring.Adoption.ActiveUserMediumWeight }
    else { $adoptionScore += [int]$scoring.Adoption.ActiveUserLowWeight }
    if ($adoptionTrendSignal -eq 'reported') { $adoptionScore += [int]$scoring.Adoption.TrendReportedWeight }

    $highRiskGrantCount = if ($apps -and $apps.PSObject.Properties.Name -contains 'HighRiskGrantCount') { Get-IntValue -Value $apps.HighRiskGrantCount -Default 0 } else { 0 }
    $deviceCompliancePercent = if ($endpoint -and $endpoint.PSObject.Properties.Name -contains 'DeviceCompliancePercent') { Get-DoubleValue -Value $endpoint.DeviceCompliancePercent -Default 0 } else { 0 }
    $governanceRisk = if ($highRiskGrantCount -gt 0) { 'elevated' } else { 'moderate' }
    $deviceScore = if ($deviceCompliancePercent -ge [double]$scoring.EndpointReadiness.ComplianceHighThreshold) {
        [int]$scoring.EndpointReadiness.HighScore
    }
    elseif ($deviceCompliancePercent -ge [double]$scoring.EndpointReadiness.ComplianceMediumThreshold) {
        [int]$scoring.EndpointReadiness.MediumScore
    }
    else {
        [int]$scoring.EndpointReadiness.LowScore
    }

    $overall = [math]::Round((($identityScore + $governanceScore + $contentScore + $adoptionScore + $deviceScore) / 5), 2)

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

    $confidenceAdjustedScore = [math]::Round(($overall * $completenessPercent) / 100, 2)

    return [ordered]@{
        ScoringModel = [ordered]@{
            IdentityAccess = $scoring.IdentityAccess
            DataGovernance = $scoring.DataGovernance
            ContentExposure = $scoring.ContentExposure
            Adoption = $scoring.Adoption
            EndpointReadiness = $scoring.EndpointReadiness
        }
        DomainScores = [ordered]@{
            IdentityAccess    = $identityScore
            DataGovernance    = $governanceScore
            ContentExposure   = $contentScore
            Adoption          = $adoptionScore
            EndpointReadiness = $deviceScore
        }
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
            AnyoneLinkSites   = $anyoneLinkSites
        }
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

    $mfaRegisteredPercent = if ($flags -and $flags.PSObject.Properties.Name -contains 'MfaRegisteredPercent') { Get-DoubleValue -Value $flags.MfaRegisteredPercent -Default 0 } else { 0 }
    $labelsPublished = if ($flags -and $flags.PSObject.Properties.Name -contains 'SensitivityLabelsPublished') { [bool]$flags.SensitivityLabelsPublished } else { $false }
    $activeUserPercent = if ($flags -and $flags.PSObject.Properties.Name -contains 'ActiveUserPercentByWorkload') { Get-DoubleValue -Value $flags.ActiveUserPercentByWorkload -Default 0 } else { 0 }
    $dlpEnforced = if ($gov -and $gov.PSObject.Properties.Name -contains 'DlpPoliciesEnforced') { Get-IntValue -Value $gov.DlpPoliciesEnforced -Default 0 } else { 0 }
    $anyoneLinkSites = if ($spo -and $spo.PSObject.Properties.Name -contains 'AnyoneLinkSites') { Get-IntValue -Value $spo.AnyoneLinkSites -Default 0 } else { 0 }
    $highRiskGrantCount = if ($apps -and $apps.PSObject.Properties.Name -contains 'HighRiskGrantCount') { Get-IntValue -Value $apps.HighRiskGrantCount -Default 0 } else { 0 }

    if ($mfaRegisteredPercent -lt 80) {
        $topRisks.Add('MFA registration is below 80% for users; identity posture is insufficient for broad AI rollout.') | Out-Null
        $quickWins.Add('Increase MFA registration coverage with targeted campaigns and registration policy enforcement.') | Out-Null
    }

    if (-not $labelsPublished) {
        $topRisks.Add('Sensitivity labels are not broadly published; data classification coverage is low.') | Out-Null
        $quickWins.Add('Publish baseline sensitivity labels for Exchange, SharePoint, Teams, and Groups.') | Out-Null
    }

    if ($dlpEnforced -eq 0) {
        $topRisks.Add('No enforced DLP policy was detected; risk of unintended data disclosure remains high.') | Out-Null
        $quickWins.Add('Move at least one high-value DLP policy from test mode to enforce mode.') | Out-Null
    }

    if ($anyoneLinkSites -gt 0) {
        $topRisks.Add("$($anyoneLinkSites) sampled SharePoint/OneDrive sites allow anonymous links.") | Out-Null
        $quickWins.Add('Tighten external sharing defaults and reduce anonymous link usage in high-value sites.') | Out-Null
    }

    if ($highRiskGrantCount -gt 0) {
        $topRisks.Add("Detected $($highRiskGrantCount) high-risk OAuth grants with elevated scopes.") | Out-Null
        $quickWins.Add('Review and remove unnecessary high-scope app grants; enforce app governance approvals.') | Out-Null
    }

    if ($activeUserPercent -lt 30) {
        $quickWins.Add('Run user enablement and scenario-led training to improve 30-day workload adoption.') | Out-Null
    }

    return [ordered]@{
        TopRisks = @($topRisks)
        QuickWins = @($quickWins)
        OverallReadiness = if ($evidence.OverallScore -ge 75) { 'strong' } elseif ($evidence.OverallScore -ge 50) { 'moderate' } else { 'needs improvement' }
    }
}

# ============================================================================
# MAIN
# ============================================================================
try {
    $resolvedOutputPath = Resolve-OutputFilePath -Path $OutputPath
    Write-Phase 'Start'
    Write-Step "Starting AI readiness export to '$resolvedOutputPath'..." -Color Yellow
    Write-Reassurance
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
        $null = Ensure-GraphConsent -RequiredScopes $requiredScopes
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
    $script:Report.ReadinessFlags = Get-ReadinessFlags
    $script:Report.ReadinessEvidence = Get-ReadinessEvidence
    $script:Report.Recommendations = Get-Recommendations

    Write-Phase 'Output'
    Write-Step "Writing JSON export to '$resolvedOutputPath'..." -Color Cyan
    $script:Report | ConvertTo-Json -Depth 15 | Out-File -FilePath $resolvedOutputPath -Encoding utf8
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
        Write-Host ("Overall AI readiness score: {0}" -f $script:Report.ReadinessEvidence.OverallScore) -ForegroundColor Cyan
    }
    if ($script:Report.Recommendations -and @($script:Report.Recommendations.TopRisks).Count -gt 0) {
        Write-Host 'Top risk:' -ForegroundColor Yellow
        Write-Host (" - {0}" -f $script:Report.Recommendations.TopRisks[0]) -ForegroundColor Yellow
    }
    if ($script:Report.Recommendations -and @($script:Report.Recommendations.QuickWins).Count -gt 0) {
        Write-Host 'Top quick win:' -ForegroundColor Green
        Write-Host (" - {0}" -f $script:Report.Recommendations.QuickWins[0]) -ForegroundColor Green
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