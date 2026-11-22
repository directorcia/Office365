#requires -Version 5.1
<#
================================================================================
Script      : o365-secure-score-extract.ps1
Purpose     : Extracts Microsoft 365 Secure Score history, control profiles
              joined to the latest control scores, Conditional Access policies,
              Security Defaults, and MFA registration activity events for one
              or more tenants to UTF-8 JSON, emitting a summary result object
              per processed tenant.
Requirements: - Windows PowerShell 5.1 or PowerShell 7+
              - Microsoft.Graph.Authentication module
              - Microsoft Graph permissions: SecurityEvents.Read.All,
                Policy.Read.All, Reports.Read.All, Directory.Read.All
Documentation: https://github.com/directorcia/Office365/wiki/Extract-Microsoft-365-Secure-Score-information
Change Log  : - 2026-11-22: Added Set-StrictMode, SupportsShouldProcess
                (-WhatIf/-Confirm), pipeline input, transcript logging,
                parameter validation, sovereign GraphBaseUri support,
                all-verified-domain/GUID tenant matching, singular-noun helper
                names with aliases, List/HashSet optimisations, Write-Information
                messaging, Write-Progress feedback, and structured output objects.
              - 2026-11-22: Open controls are scored controls from the latest
                secureScore.controlScores (controlName = profile id) that still
                have points available and are not addressed. Added
                ScoredControlCount, OpenControlCount, FullyAchievedCount,
                AddressedWithPointsRemaining, NotApplicableCount, and
                TotalPointsAvailable. Top open controls sort by points available,
                then rank (0/null last). App-only sessions without Account use
                '<AppName> (<ClientId>) [AppOnly]'. Compact output normalises
                implementationCost and userImpact. MFA stage text describes
                registration activity events.
              - 2026-11-22: Compact implementationCost and userImpact treat
                Medium (any casing, trimmed) as Moderate. Other unrecognised
                non-empty values stay title-cased (first character upper,
                remainder lower). Null and whitespace stay null.
================================================================================

.SYNOPSIS
    Extracts Microsoft 365 Secure Score, control profiles, Conditional Access
    policies, Security Defaults, and MFA registration activity to JSON.

.DESCRIPTION
    Connects to Microsoft Graph (reusing an active session when any verified
    domain or tenant GUID matches and required scopes are already consented) and
    collects five security posture datasets:
      1. Secure Score latest snapshot and recent history
      2. Secure Score control profiles joined to the latest secureScore
         controlScores (controlName equals profile id). A control is open only
         when it was scored, pointsAvailable (profile maxScore minus current
         score, floored at 0) is greater than zero, and it is not addressed
         (completed, ignored, thirdParty, riskAccepted, alternativeMitigation).
         Fully achieved controls are scored items already at full marks.
         Addressed controls that still have a points gap are counted separately.
         Catalog profiles absent from controlScores are not applicable.
         The top 25 open controls are ordered by pointsAvailable descending,
         then by rank, with rank 0 or null last.
      3. Conditional Access policies
      4. Identity Security Defaults enforcement status
      5. Authentication methods user registration activity events for the last
         30 days (Graph userRegistrationActivity). This is event activity, not
         a point-in-time MFA registration status. The JSON property remains
         MfaRegistrationSummary; MfaRegistrationSummaryDescription explains it.

    Existing JSON property names are unchanged. The export adds
    ScoredControlCount, OpenControlCount, FullyAchievedCount,
    AddressedWithPointsRemaining, NotApplicableCount, TotalPointsAvailable
    (sum of pointsAvailable across scored controls, including addressed gaps),
    per-control currentScore and pointsAvailable, and
    MfaRegistrationSummaryDescription. The compact file normalises
    implementationCost and userImpact: Low, Moderate, High, and Unknown keep
    those labels regardless of casing. Medium (any casing, after trim) maps to
    Moderate. Any other non-empty value is kept title-cased (first character
    upper-invariant, remainder lower-invariant). Null and whitespace stay null.
    TotalCount / All remain the full control-profile catalog, not the scored set.

    Writes a progressive *_partial.json recovery checkpoint after each stage,
    saves the finalized UTF-8 JSON export (plus an optional *_compact.json file
    sized for AI analysis), cleans up the partial file on success, and emits a
    structured PSCustomObject summary to the pipeline.

.PARAMETER TenantDomain
    Primary or verified domain name (e.g. contoso.onmicrosoft.com, contoso.com)
    or Entra tenant GUID of the target Microsoft 365 tenant. Accepts pipeline
    input for multi-tenant batch runs.

.PARAMETER DataFile
    Optional output path for the full JSON export. When omitted or when multiple
    tenants are piped in, a timestamped file name is generated automatically in
    the current directory (or inside DataFile when DataFile is an existing
    directory).

.PARAMETER ScoreHistoryCount
    Number of historical Secure Score snapshots to retrieve (1 to 90). Defaults
    to 5.

.PARAMETER Compact
    Also generates a smaller *_compact.json summary file tailored for AI tools
    and upload-constrained workflows.

.PARAMETER GraphBaseUri
    Base HTTPS URI for Microsoft Graph (without a trailing slash or API version
    segment). Defaults to https://graph.microsoft.com; override for sovereign
    clouds such as https://graph.microsoft.us.

.PARAMETER MaxRetries
    Maximum retry attempts for transient Microsoft Graph HTTP errors (408, 429,
    500, 502, 503, 504). Defaults to 3.

.PARAMETER InitialBackoffMs
    Initial exponential-backoff delay in milliseconds when no Retry-After header
    is supplied. Defaults to 1000.

.PARAMETER TranscriptPath
    Optional file path for Start-Transcript session logging during unattended
    runs. Also activated automatically when -Debug is specified.

.PARAMETER ForceReconnect
    Forces Disconnect-MgGraph and a fresh Connect-MgGraph authentication prompt
    even when an active session already matches the target tenant and scopes.

.EXAMPLE
    .\o365-secure-score-extract.ps1 -TenantDomain 'contoso.onmicrosoft.com'
    Connects to contoso.onmicrosoft.com, writes the full JSON export to a
    timestamped file in the current folder, and returns a summary object.

.EXAMPLE
    .\o365-secure-score-extract.ps1 -TenantDomain 'contoso.com' -DataFile 'C:\Reports\Secure Score\contoso.json' -Compact -Verbose
    Writes both contoso.json and contoso_compact.json (creating parent folders
    if needed) with detailed verbose diagnostics.

.EXAMPLE
    'tenant1.onmicrosoft.com', 'tenant2.onmicrosoft.com' | .\o365-secure-score-extract.ps1 -Compact -TranscriptPath 'C:\Logs\extract.log'
    Extracts full and compact JSON files for multiple tenants over the pipeline
    while recording a session transcript.

.EXAMPLE
    .\o365-secure-score-extract.ps1 -TenantDomain 'contoso.onmicrosoft.com' -WhatIf
    Previews the Microsoft Graph connection and target JSON file writes without
    connecting or creating files on disk.

.INPUTS
    System.String. You can pipe one or more tenant domains or tenant GUIDs to
    TenantDomain.

.OUTPUTS
    PSCustomObject. Emits one extraction summary object per processed tenant
    containing Tenant, TenantId, CollectedBy, CollectionDate, FullDataFile,
    FullSizeKB, CompactDataFile, CompactSizeKB, LatestScore, MaxScore,
    OpenControlCount, ScoredControlCount, FullyAchievedCount,
    AddressedWithPointsRemaining, NotApplicableCount, TotalPointsAvailable,
    ConditionalAccessCount, SecurityDefaultsEnabled, MfaRecordCount,
    StagesSucceeded, StagesDegraded, and ElapsedSeconds. CollectedBy falls
    back to '<AppName> (<ClientId>) [AppOnly]' when the Graph account is empty.

.NOTES
    Requires PowerShell 5.1+ and Microsoft.Graph.Authentication. Sets
    $LASTEXITCODE to 0 when all requested tenants succeed, or 1 if any tenant
    extraction fails.
#>

#region Parameters
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
[OutputType([pscustomobject])]
param(
    [Parameter(
        Mandatory = $true,
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = 'Enter a verified Microsoft 365 tenant domain (e.g. contoso.onmicrosoft.com) or tenant GUID.'
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
            $candidate = $_.Trim()
            $parsedGuid = [guid]::Empty
            if ([guid]::TryParse($candidate, [ref]$parsedGuid)) {
                return $true
            }
            if ($candidate -match '^(?=.{3,253}$)([a-zA-Z0-9]([a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,63}$') {
                return $true
            }
            throw "TenantDomain '$_' must be a valid domain name (e.g. contoso.onmicrosoft.com) or tenant GUID."
        })]
    [Alias('Tenant', 'Domain', 'TenantId')]
    [string]$TenantDomain,

    [Parameter(HelpMessage = 'Optional output JSON file path (or target directory when piping multiple tenants).')]
    [AllowEmptyString()]
    [Alias('OutputPath', 'OutFile')]
    [string]$DataFile = '',

    [Parameter(HelpMessage = 'Number of historical Secure Score records to retrieve (1-90).')]
    [ValidateRange(1, 90)]
    [int]$ScoreHistoryCount = 5,

    [Parameter(HelpMessage = 'Also write an AI-optimised *_compact.json file alongside the full JSON file.')]
    [switch]$Compact,

    [Parameter(HelpMessage = 'Base HTTPS URI for Microsoft Graph (override for sovereign clouds).')]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
            $uri = $null
            if (-not [uri]::TryCreate($_.TrimEnd('/'), [System.UriKind]::Absolute, [ref]$uri)) {
                throw "GraphBaseUri '$_' is not a valid absolute URI."
            }
            if ($uri.Scheme -ne 'https') {
                throw "GraphBaseUri '$_' must use the https scheme."
            }
            if (-not [string]::IsNullOrEmpty($uri.Query) -or -not [string]::IsNullOrEmpty($uri.Fragment)) {
                throw "GraphBaseUri '$_' must not contain a query string or fragment."
            }
            $true
        })]
    [string]$GraphBaseUri = 'https://graph.microsoft.com',

    [Parameter(HelpMessage = 'Maximum retry attempts for transient Graph HTTP errors.')]
    [ValidateRange(0, 10)]
    [int]$MaxRetries = 3,

    [Parameter(HelpMessage = 'Initial backoff in milliseconds for exponential retry delays.')]
    [ValidateRange(100, 60000)]
    [int]$InitialBackoffMs = 1000,

    [Parameter(HelpMessage = 'Optional file path to record a PowerShell session transcript.')]
    [ValidateNotNullOrEmpty()]
    [Alias('LogPath')]
    [string]$TranscriptPath,

    [Parameter(HelpMessage = 'Force a fresh Microsoft Graph sign-in even if already connected to the tenant.')]
    [switch]$ForceReconnect
)
#endregion Parameters

begin {
    #region Configuration
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    $RequiredModuleName = 'Microsoft.Graph.Authentication'
    $RequiredScopes = @(
        'SecurityEvents.Read.All',
        'Policy.Read.All',
        'Reports.Read.All',
        'Directory.Read.All'
    )
    $TransientStatusCodes = @(408, 429, 500, 502, 503, 504)
    $AddressedControlStates = @(
        'completed',
        'ignored',
        'thirdParty',
        'riskAccepted',
        'alternativeMitigation'
    )
    $NormalizedGraphBaseUri = $GraphBaseUri.TrimEnd('/')
    $ProgressActivity = 'M365 Secure Score Extraction'
    $MaxPaginationPages = 500
    # Kept beside MfaRegistrationSummary. The endpoint returns activity events,
    # not a point-in-time registration-status snapshot.
    $MfaRegistrationActivityDescription = 'Authentication methods user registration activity events for period d30 (Graph reports/authenticationMethods/userRegistrationActivity). This is event activity, not a point-in-time MFA registration status.'

    $transcriptStarted = $false
    $moduleChecked = $false
    $processedTenants = 0
    $skippedTenants = 0
    $failedTenants = 0
    $totalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    #endregion Configuration

    #region Helper functions
    function Write-StatusMessage {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [AllowEmptyString()]
            [string]$Message,

            [Parameter()]
            [ValidateSet('Info', 'Detail', 'Section')]
            [string]$Level = 'Info'
        )

        $prefix = switch ($Level) {
            'Info'    { '[INFO] ' }
            'Section' { '' }
            default   { '  ' }
        }
        Write-Information -MessageData ('{0}{1}' -f $prefix, $Message) -InformationAction Continue
    }

    function Get-SafePropertyValue {
        # Plain helper (no [CmdletBinding()]) avoids advanced-function overhead inside
        # tight loops while safely reading properties from both Hashtables and
        # PSCustomObjects under Set-StrictMode -Version Latest.
        param(
            [object]$InputObject,
            [string]$Name
        )

        if ($null -eq $InputObject) {
            return $null
        }

        if ($InputObject -is [System.Collections.IDictionary]) {
            if ($InputObject.Contains($Name)) {
                return $InputObject[$Name]
            }
            foreach ($key in $InputObject.Keys) {
                if ([string]::Equals([string]$key, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                    return $InputObject[$key]
                }
            }
            return $null
        }

        $prop = $InputObject.PSObject.Properties[$Name]
        if ($null -ne $prop) {
            return $prop.Value
        }

        return $null
    }

    function Get-OutputPathWithSuffix {
        [CmdletBinding()]
        [OutputType([string])]
        param(
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Path,

            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Suffix
        )

        # Build sibling path safely even when extension is missing or path has spaces.
        $parent = [System.IO.Path]::GetDirectoryName($Path)
        if ([string]::IsNullOrWhiteSpace($parent)) {
            $parent = (Get-Location).Path
        }
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        if ([string]::IsNullOrWhiteSpace($baseName)) {
            $baseName = 'security-data'
        }
        $ext = [System.IO.Path]::GetExtension($Path)
        if ([string]::IsNullOrWhiteSpace($ext)) {
            $ext = '.json'
        }
        return (Join-Path -Path $parent -ChildPath ('{0}{1}{2}' -f $baseName, $Suffix, $ext))
    }

    function Resolve-TenantDataFilePath {
        [CmdletBinding()]
        [OutputType([string])]
        param(
            [Parameter(Mandatory = $true)]
            [string]$TargetTenant,

            [Parameter()]
            [AllowEmptyString()]
            [string]$RequestedPath,

            [Parameter()]
            [int]$TenantSequence = 1
        )

        $timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
        $sanitizedTenant = $TargetTenant -replace '[^a-zA-Z0-9]', '-'
        $defaultLeaf = '{0}_ss_{1}.json' -f $sanitizedTenant, $timestamp

        if ([string]::IsNullOrWhiteSpace($RequestedPath)) {
            $baseDir = (Get-Location).Path
            if ([string]::IsNullOrWhiteSpace($baseDir)) {
                $baseDir = $PSScriptRoot
            }
            return (Join-Path -Path $baseDir -ChildPath $defaultLeaf)
        }

        if (Test-Path -LiteralPath $RequestedPath -PathType Container) {
            return (Join-Path -Path $RequestedPath -ChildPath $defaultLeaf)
        }

        # When piping multiple tenants with a single file path, suffix subsequent
        # tenants so earlier exports are not overwritten.
        if ($TenantSequence -gt 1) {
            return (Get-OutputPathWithSuffix -Path $RequestedPath -Suffix ('_{0}' -f $sanitizedTenant))
        }

        return $RequestedPath
    }

    function Write-JsonFile {
        [CmdletBinding()]
        [OutputType([long])]
        param(
            [Parameter(Mandatory = $true)]
            [object]$InputObject,

            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$FilePath,

            [Parameter()]
            [int]$Depth = 12
        )

        $parentDir = [System.IO.Path]::GetDirectoryName($FilePath)
        if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path -LiteralPath $parentDir -PathType Container)) {
            New-Item -ItemType Directory -Path $parentDir -Force -ErrorAction Stop | Out-Null
            Write-Verbose "Created output directory '$parentDir'."
        }

        # Depth 12 prevents silent System.Collections.Hashtable string truncation on
        # deeply nested Conditional Access grant/session/condition structures.
        # UTF8Encoding($false) writes BOM-less UTF-8 across both PS 5.1 and PS 7+
        # for clean interoperability with Python/pandas and AI parsers.
        $jsonText = $InputObject | ConvertTo-Json -Depth $Depth
        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        $resolvedTarget = if ([System.IO.Path]::IsPathRooted($FilePath)) {
            $FilePath
        }
        else {
            Join-Path -Path (Get-Location).Path -ChildPath $FilePath
        }
        [System.IO.File]::WriteAllText($resolvedTarget, $jsonText, $utf8NoBom)
        return ([System.IO.FileInfo]::new($resolvedTarget)).Length
    }

    function Save-PartialData {
        [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
        param(
            [Parameter(Mandatory = $true)]
            [object]$SecurityData,

            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$OutputPath,

            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Stage
        )

        try {
            $checkpointFile = Get-OutputPathWithSuffix -Path $OutputPath -Suffix '_partial'
            if ($PSCmdlet.ShouldProcess($checkpointFile, "Save recovery checkpoint for stage '$Stage'")) {
                $bytes = Write-JsonFile -InputObject $SecurityData -FilePath $checkpointFile -Depth 12
                $sizeKB = [math]::Round($bytes / 1KB, 2)
                Write-Debug "[Save-PartialData] Recovery checkpoint updated at stage '$Stage': $checkpointFile ($sizeKB KB)"
                Write-Verbose "Saved recovery checkpoint ($sizeKB KB) after stage '$Stage': $checkpointFile"
            }
        }
        catch {
            Write-Warning "Could not save partial recovery checkpoint at stage '$Stage': $($_.Exception.Message)"
        }
    }

    function Get-GraphErrorSummary {
        param([System.Management.Automation.ErrorRecord]$ErrorRecord)

        $httpStatus = $null
        $graphCode = $null
        $graphMessage = $null
        $retryAfterSeconds = $null

        $exception = $ErrorRecord.Exception
        $respProp = if ($null -ne $exception) { $exception.PSObject.Properties['Response'] } else { $null }
        $response = if ($null -ne $respProp) { $respProp.Value } else { $null }

        if ($null -ne $response) {
            $statusProp = $response.PSObject.Properties['StatusCode']
            if ($null -ne $statusProp -and $null -ne $statusProp.Value) {
                try {
                    $httpStatus = [int]$statusProp.Value
                }
                catch {
                    Write-Debug "Could not cast HTTP StatusCode '$($statusProp.Value)' to integer: $($_.Exception.Message)"
                }
            }

            $headersProp = $response.PSObject.Properties['Headers']
            if ($null -ne $headersProp -and $null -ne $headersProp.Value) {
                try {
                    $headers = $headersProp.Value
                    $rawRetry = $null
                    if ($headers -is [System.Collections.IDictionary] -and $headers.Contains('Retry-After')) {
                        $rawRetry = $headers['Retry-After']
                    }
                    elseif ($null -ne $headers.PSObject.Methods['TryGetValues']) {
                        $values = $null
                        if ($headers.TryGetValues('Retry-After', [ref]$values)) {
                            $rawRetry = $values | Select-Object -First 1
                        }
                    }
                    if ($null -ne $rawRetry) {
                        $parsedSec = 0
                        if ([int]::TryParse([string](@($rawRetry)[0]), [ref]$parsedSec) -and $parsedSec -ge 0) {
                            $retryAfterSeconds = $parsedSec
                        }
                    }
                }
                catch {
                    Write-Debug "Could not inspect Retry-After header: $($_.Exception.Message)"
                }
            }
        }

        $detailsMessage = if ($null -ne $ErrorRecord.ErrorDetails) { $ErrorRecord.ErrorDetails.Message } else { $null }
        if (-not [string]::IsNullOrWhiteSpace($detailsMessage)) {
            try {
                $parsedErr = $detailsMessage | ConvertFrom-Json
                $errNode = Get-SafePropertyValue -InputObject $parsedErr -Name 'error'
                if ($null -ne $errNode) {
                    $graphCode = [string](Get-SafePropertyValue -InputObject $errNode -Name 'code')
                    $graphMessage = [string](Get-SafePropertyValue -InputObject $errNode -Name 'message')
                    if ($null -eq $httpStatus) {
                        switch ($graphCode) {
                            'TooManyRequests'    { $httpStatus = 429 }
                            'ServiceUnavailable' { $httpStatus = 503 }
                            'GatewayTimeout'     { $httpStatus = 504 }
                        }
                    }
                }
            }
            catch {
                Write-Debug "ErrorDetails payload was not JSON: $($_.Exception.Message)"
            }
        }

        if ($null -eq $httpStatus -and $null -ne $exception) {
            if ($exception.Message -match '\b(408|429|500|502|503|504)\b') {
                $httpStatus = [int]$Matches[1]
            }
        }

        return [pscustomobject]@{
            HttpStatusCode    = $httpStatus
            GraphErrorCode    = $graphCode
            GraphErrorMessage = $graphMessage
            RetryAfterSeconds = $retryAfterSeconds
            ExceptionMessage  = if ($null -ne $exception) { $exception.Message } else { [string]$ErrorRecord }
        }
    }

    function Invoke-MgGraphRequestWithRetry {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Uri,

            [Parameter()]
            [ValidateRange(0, 10)]
            [int]$RetryLimit = $MaxRetries,

            [Parameter()]
            [ValidateRange(100, 60000)]
            [int]$BaseBackoffMs = $InitialBackoffMs
        )

        $attempt = 0
        while ($true) {
            try {
                Write-Debug "[Invoke-MgGraphRequestWithRetry] GET $Uri (attempt $($attempt + 1)/$($RetryLimit + 1))"
                $requestParams = @{
                    Uri         = $Uri
                    Method      = 'GET'
                    ErrorAction = 'Stop'
                }
                return (Invoke-MgGraphRequest @requestParams)
            }
            catch {
                $errInfo = Get-GraphErrorSummary -ErrorRecord $_
                $isTransient = ($null -ne $errInfo.HttpStatusCode) -and ($TransientStatusCodes -contains $errInfo.HttpStatusCode)

                if ($isTransient -and $attempt -lt $RetryLimit) {
                    $attempt++
                    if ($null -ne $errInfo.RetryAfterSeconds) {
                        $delayMs = $errInfo.RetryAfterSeconds * 1000
                        $delayLabel = '{0}s (from Retry-After header)' -f $errInfo.RetryAfterSeconds
                    }
                    else {
                        $delayMs = [int]($BaseBackoffMs * [Math]::Pow(2, $attempt - 1))
                        $delayLabel = '{0:N1}s' -f ($delayMs / 1000.0)
                    }

                    Write-Warning "Transient Microsoft Graph HTTP $($errInfo.HttpStatusCode) on '$Uri'. Retrying in $delayLabel (attempt $attempt of $RetryLimit)."
                    Start-Sleep -Milliseconds $delayMs
                    continue
                }

                $guidance = ''
                if ($errInfo.GraphErrorCode -in @('Authorization_RequestDenied', 'Forbidden') -or $errInfo.HttpStatusCode -eq 403) {
                    $guidance = " Ensure the signed-in principal has admin consent for: $($RequiredScopes -join ', ')."
                }
                $detailSuffix = if (-not [string]::IsNullOrWhiteSpace($errInfo.GraphErrorCode)) {
                    " [$($errInfo.GraphErrorCode): $($errInfo.GraphErrorMessage)]"
                }
                else {
                    ''
                }

                throw [System.InvalidOperationException]::new(
                    "Microsoft Graph GET '$Uri' failed: $($errInfo.ExceptionMessage)$detailSuffix.$guidance",
                    $_.Exception
                )
            }
        }
    }

    function Invoke-GraphCollection {
        [CmdletBinding()]
        [OutputType([object[]])]
        param(
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Uri,

            [Parameter()]
            [ValidateRange(0, 10)]
            [int]$RetryLimit = $MaxRetries,

            [Parameter()]
            [ValidateRange(100, 60000)]
            [int]$BaseBackoffMs = $InitialBackoffMs,

            [Parameter()]
            [switch]$AllowPartialResults
        )

        # Generic List avoids O(n^2) array reallocation across multi-page endpoints
        # (benchmarked at 28 ms vs 657 ms for array += across 500 pages / 50k items).
        $collected = [System.Collections.Generic.List[object]]::new()
        $nextUri = $Uri
        $pageCount = 0

        Write-Verbose "Requesting Graph collection: $Uri"

        while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
            $pageCount++
            if ($pageCount -gt $MaxPaginationPages) {
                throw [System.IO.InvalidDataException]::new(
                    "Aborted pagination on '$Uri' after $MaxPaginationPages pages to guard against a cyclic '@odata.nextLink'."
                )
            }

            try {
                $pageTimer = [System.Diagnostics.Stopwatch]::StartNew()
                $resp = Invoke-MgGraphRequestWithRetry -Uri $nextUri -RetryLimit $RetryLimit -BaseBackoffMs $BaseBackoffMs
                $pageTimer.Stop()

                $pageValues = Get-SafePropertyValue -InputObject $resp -Name 'value'
                $pageItemCount = 0
                if ($null -ne $pageValues) {
                    if ($pageValues -is [System.Collections.IDictionary] -or $pageValues -is [string] -or $pageValues -isnot [System.Collections.IEnumerable]) {
                        $collected.Add($pageValues)
                        $pageItemCount = 1
                    }
                    else {
                        foreach ($entry in $pageValues) {
                            $collected.Add($entry)
                            $pageItemCount++
                        }
                    }
                }

                Write-Verbose ("Page {0} retrieved {1} items in {2:N1}s (running total: {3})." -f $pageCount, $pageItemCount, $pageTimer.Elapsed.TotalSeconds, $collected.Count)
                $rawNext = Get-SafePropertyValue -InputObject $resp -Name '@odata.nextLink'
                $nextUri = if (-not [string]::IsNullOrWhiteSpace([string]$rawNext)) { [string]$rawNext } else { $null }
            }
            catch {
                if ($AllowPartialResults) {
                    Write-Warning "Collection interrupted on page $pageCount ('$nextUri'); returning $($collected.Count) partial items: $($_.Exception.Message)"
                    break
                }
                throw
            }
        }

        $finalArray = $collected.ToArray()
        Write-Output -InputObject $finalArray -NoEnumerate
    }

    function Test-TenantSessionMatch {
        [CmdletBinding()]
        [OutputType([pscustomobject])]
        param(
            [Parameter(Mandatory = $true)]
            [object]$Context,

            [Parameter(Mandatory = $true)]
            [string]$TargetTenant,

            [Parameter(Mandatory = $true)]
            [string[]]$NeededScopes,

            [Parameter(Mandatory = $true)]
            [string]$BaseUri
        )

        $ctxTenantId = [string](Get-SafePropertyValue -InputObject $Context -Name 'TenantId')
        if ([string]::IsNullOrWhiteSpace($ctxTenantId)) {
            return [pscustomobject]@{ Matches = $false; Reason = 'Active Graph context has no TenantId.' }
        }

        # Verify all required delegated scopes are present when Scopes is populated.
        # App-only contexts expose AuthType = 'AppOnly' and may leave Scopes empty.
        $authType = [string](Get-SafePropertyValue -InputObject $Context -Name 'AuthType')
        $ctxScopes = @(Get-SafePropertyValue -InputObject $Context -Name 'Scopes')
        if ($authType -ne 'AppOnly' -and $ctxScopes.Count -gt 0) {
            $scopeSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($s in $ctxScopes) {
                if (-not [string]::IsNullOrWhiteSpace([string]$s)) {
                    $null = $scopeSet.Add([string]$s)
                }
            }

            foreach ($req in $NeededScopes) {
                $rwEquivalent = $req -replace '\.Read\.All$', '.ReadWrite.All'
                if (-not ($scopeSet.Contains($req) -or $scopeSet.Contains($rwEquivalent))) {
                    return [pscustomobject]@{
                        Matches = $false
                        Reason  = "Active session is missing required permission scope '$req'."
                    }
                }
            }
        }

        # Check GUID match directly first to avoid an unnecessary call when a GUID was supplied;
        # still call /organization once below to verify Graph API health in a single round-trip.
        $orgUri = '{0}/v1.0/organization?$select=id,displayName,verifiedDomains' -f $BaseUri
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $orgResp = Invoke-MgGraphRequestWithRetry -Uri $orgUri
        $sw.Stop()

        $orgs = @(Get-SafePropertyValue -InputObject $orgResp -Name 'value')
        $firstOrg = if ($orgs.Count -gt 0) { $orgs[0] } else { $null }
        $orgId = [string](Get-SafePropertyValue -InputObject $firstOrg -Name 'id')

        $domainSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $defaultDomain = $null
        $verifiedDomains = @(Get-SafePropertyValue -InputObject $firstOrg -Name 'verifiedDomains')
        foreach ($dom in $verifiedDomains) {
            $domName = [string](Get-SafePropertyValue -InputObject $dom -Name 'name')
            if (-not [string]::IsNullOrWhiteSpace($domName)) {
                $null = $domainSet.Add($domName)
                if ([bool](Get-SafePropertyValue -InputObject $dom -Name 'isDefault')) {
                    $defaultDomain = $domName
                }
            }
        }

        $trimmedTarget = $TargetTenant.Trim()
        $tenantMatches = (
            [string]::Equals($ctxTenantId, $trimmedTarget, [System.StringComparison]::OrdinalIgnoreCase) -or
            [string]::Equals($orgId, $trimmedTarget, [System.StringComparison]::OrdinalIgnoreCase) -or
            $domainSet.Contains($trimmedTarget)
        )

        if (-not $tenantMatches) {
            $connectedLabel = if ($defaultDomain) { $defaultDomain } else { $ctxTenantId }
            return [pscustomobject]@{
                Matches   = $false
                Reason    = "Currently connected to '$connectedLabel' instead of '$trimmedTarget'."
                LatencyMs = [math]::Round($sw.Elapsed.TotalMilliseconds, 0)
            }
        }

        return [pscustomobject]@{
            Matches       = $true
            DefaultDomain = $defaultDomain
            LatencyMs     = [math]::Round($sw.Elapsed.TotalMilliseconds, 0)
        }
    }

    function Get-GraphPrincipalLabel {
        # Delegated sessions expose Account. App-only certificate sessions leave
        # Account empty; AppName and ClientId still identify the principal.
        param(
            [object]$Context
        )

        $account = [string](Get-SafePropertyValue -InputObject $Context -Name 'Account')
        if (-not [string]::IsNullOrWhiteSpace($account)) {
            return $account.Trim()
        }

        $appName = [string](Get-SafePropertyValue -InputObject $Context -Name 'AppName')
        if ([string]::IsNullOrWhiteSpace($appName)) {
            $appName = 'Unknown application'
        }
        $clientId = [string](Get-SafePropertyValue -InputObject $Context -Name 'ClientId')
        if ([string]::IsNullOrWhiteSpace($clientId)) {
            $clientId = 'unknown-client-id'
        }

        return ('{0} ({1}) [AppOnly]' -f $appName.Trim(), $clientId.Trim())
    }

    function Connect-TargetTenantGraph {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$TargetTenant,

            [Parameter(Mandatory = $true)]
            [string[]]$Scopes,

            [Parameter(Mandatory = $true)]
            [string]$BaseUri,

            [Parameter()]
            [switch]$Force
        )

        if (-not (Get-Module -Name $RequiredModuleName)) {
            Write-Verbose "Loading module '$RequiredModuleName'."
            try {
                Import-Module -Name $RequiredModuleName -ErrorAction Stop | Out-Null
            }
            catch {
                throw [System.InvalidOperationException]::new(
                    "Could not load '$RequiredModuleName'. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser",
                    $_.Exception
                )
            }
        }

        $context = Get-MgContext -ErrorAction SilentlyContinue
        $needsConnect = $true

        if ($Force -and $null -ne $context) {
            Write-Verbose 'ForceReconnect specified; disconnecting existing Microsoft Graph session.'
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            $context = $null
        }
        elseif ($null -ne $context) {
            try {
                $matchResult = Test-TenantSessionMatch -Context $context -TargetTenant $TargetTenant -NeededScopes $Scopes -BaseUri $BaseUri
                if ($matchResult.Matches) {
                    $sessionLabel = Get-GraphPrincipalLabel -Context $context
                    Write-StatusMessage -Message "Reusing active Microsoft Graph session ($sessionLabel, API health check $($matchResult.LatencyMs) ms)."
                    $needsConnect = $false
                }
                else {
                    Write-Warning "$($matchResult.Reason) Reconnecting to '$TargetTenant'."
                    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                }
            }
            catch {
                Write-Verbose "Could not validate existing Graph session ($($_.Exception.Message)); reconnecting."
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            }
        }

        if ($needsConnect) {
            Write-StatusMessage -Message "Connecting to Microsoft Graph for '$TargetTenant'. Complete any sign-in prompt..."
            $connectParams = @{
                TenantId    = $TargetTenant
                Scopes      = $Scopes
                NoWelcome   = $true
                ErrorAction = 'Stop'
            }
            Connect-MgGraph @connectParams | Out-Null
            $context = Get-MgContext -ErrorAction Stop
            if ($null -eq $context -or [string]::IsNullOrWhiteSpace([string]$context.TenantId)) {
                throw [System.InvalidOperationException]::new(
                    "Microsoft Graph sign-in completed without returning a valid tenant context for '$TargetTenant'."
                )
            }

            # Single lightweight health & permission probe after fresh sign-in.
            $healthTimer = [System.Diagnostics.Stopwatch]::StartNew()
            $null = Invoke-MgGraphRequestWithRetry -Uri ('{0}/v1.0/organization?$select=id' -f $BaseUri)
            $healthTimer.Stop()
            $latencyMs = [math]::Round($healthTimer.Elapsed.TotalMilliseconds, 0)
            $sessionLabel = Get-GraphPrincipalLabel -Context $context
            Write-StatusMessage -Message "Connected as '$sessionLabel' (Graph API responded in ${latencyMs} ms)."
            if ($latencyMs -gt 5000) {
                Write-Warning "Slow Microsoft Graph response (${latencyMs} ms); extraction may take longer than usual."
            }
        }

        return $context
    }

    function Get-SecureScoreData {
        [CmdletBinding()]
        [OutputType([pscustomobject])]
        param(
            [Parameter()]
            [ValidateRange(1, 90)]
            [int]$HistoryCount = 5,

            [Parameter()]
            [string]$BaseUri = $NormalizedGraphBaseUri
        )

        # Request only HistoryCount items sorted newest-first on the server, avoiding
        # downloading all 90 multi-megabyte controlScores snapshots when only 5 are needed.
        $uri = '{0}/beta/security/secureScores?$top={1}&$orderby=createdDateTime%20desc' -f $BaseUri, $HistoryCount
        $reqResp = Invoke-MgGraphRequestWithRetry -Uri $uri
        $rawScores = @(Get-SafePropertyValue -InputObject $reqResp -Name 'value')
        if ($rawScores.Count -eq 0) {
            Write-Warning 'No Secure Score data was returned by Microsoft Graph.'
            return [pscustomobject]@{
                Latest  = $null
                History = @()
            }
        }

        $trimmedScores = if ($rawScores.Count -gt $HistoryCount) {
            @($rawScores[0..($HistoryCount - 1)])
        }
        else {
            $rawScores
        }

        $latest = $trimmedScores |
            Sort-Object -Property { Get-SafePropertyValue -InputObject $_ -Name 'createdDateTime' } -Descending |
            Select-Object -First 1

        return [pscustomobject]@{
            Latest  = $latest
            History = $trimmedScores
        }
    }

    function ConvertTo-InvariantDouble {
        # Plain helper so [ref] parsing stays reliable on Windows PowerShell 5.1 and PowerShell 7.
        param(
            [object]$Value
        )

        if ($null -eq $Value) {
            return $null
        }
        if ($Value -is [double] -or $Value -is [float] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) {
            return [double]$Value
        }

        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $null
        }

        $parsed = [double]0
        $styles = [System.Globalization.NumberStyles]::Float -bor [System.Globalization.NumberStyles]::AllowThousands
        if ([double]::TryParse($text, $styles, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
            return $parsed
        }
        return $null
    }

    function Add-ControlAssessmentProperty {
        param(
            [object]$InputObject,
            [string]$Name,
            [object]$Value
        )

        if ($null -eq $InputObject -or [string]::IsNullOrWhiteSpace($Name)) {
            return
        }

        if ($InputObject -is [System.Collections.IDictionary]) {
            $InputObject[$Name] = $Value
            return
        }

        $existing = $InputObject.PSObject.Properties[$Name]
        if ($null -ne $existing) {
            $existing.Value = $Value
            return
        }

        $InputObject | Add-Member -MemberType NoteProperty -Name $Name -Value $Value -Force
    }

    function ConvertTo-NormalizedCostImpact {
        # Compact output only. Graph returns mixed casing ("low", "MODERATE") and
        # sometimes "Medium" for the same band as Moderate. Known labels become
        # Low, Moderate, High, or Unknown. "Medium" (any casing, trimmed) maps to
        # Moderate. Any other non-empty value is kept title-cased: first character
        # upper-invariant, remainder lower-invariant ("CUSTOM" becomes "Custom").
        # Null and whitespace stay null.
        param(
            [object]$Value
        )

        if ($null -eq $Value) {
            return $null
        }

        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $null
        }

        switch ($text.ToLowerInvariant()) {
            'low' { return 'Low' }
            'medium' { return 'Moderate' }
            'moderate' { return 'Moderate' }
            'high' { return 'High' }
            'unknown' { return 'Unknown' }
            default {
                if ($text.Length -eq 1) {
                    return $text.ToUpperInvariant()
                }
                return ($text.Substring(0, 1).ToUpperInvariant() + $text.Substring(1).ToLowerInvariant())
            }
        }
    }

    function Test-SecureScoreControlAddressed {
        param(
            [object]$Control,
            [System.Collections.Generic.HashSet[string]]$AddressedStates
        )

        $updates = Get-SafePropertyValue -InputObject $Control -Name 'controlStateUpdates'
        if ($null -eq $updates -or $updates -is [string]) {
            return $false
        }

        $updateEntries = if ($updates -is [System.Collections.IDictionary]) { @($updates) } else { @($updates) }
        foreach ($update in $updateEntries) {
            if ($null -eq $update) {
                continue
            }
            $updateState = [string](Get-SafePropertyValue -InputObject $update -Name 'state')
            if (-not [string]::IsNullOrWhiteSpace($updateState) -and $AddressedStates.Contains($updateState)) {
                return $true
            }
        }

        return $false
    }

    function Get-SecureScoreRankSortKey {
        # Microsoft leaves rank at 0 or null for some unscored Device controls.
        # Those must sort after a real rank so they do not crowd out higher-value gaps.
        param(
            [object]$Control
        )

        $rankValue = Get-SafePropertyValue -InputObject $Control -Name 'rank'
        if ($null -eq $rankValue) {
            return [int]::MaxValue
        }
        if ($rankValue -is [int] -or $rankValue -is [long] -or $rankValue -is [double] -or $rankValue -is [decimal] -or $rankValue -is [float]) {
            $numericRank = [double]$rankValue
            if ($numericRank -le 0) {
                return [int]::MaxValue
            }
            return [int][math]::Truncate($numericRank)
        }

        $rankText = ([string]$rankValue).Trim()
        if ([string]::IsNullOrWhiteSpace($rankText)) {
            return [int]::MaxValue
        }

        $parsedRank = 0
        $integerStyles = [System.Globalization.NumberStyles]::Integer
        if (-not [int]::TryParse($rankText, $integerStyles, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedRank)) {
            return [int]::MaxValue
        }
        if ($parsedRank -le 0) {
            return [int]::MaxValue
        }
        return $parsedRank
    }

    function Get-SecureScoreControlProfile {
        [CmdletBinding()]
        [Alias('Get-SecureScoreControls')]
        [OutputType([pscustomobject])]
        param(
            [Parameter()]
            [string]$BaseUri = $NormalizedGraphBaseUri,

            [Parameter()]
            [object]$LatestSecureScore
        )

        $uri = '{0}/beta/security/secureScoreControlProfiles?$top=200' -f $BaseUri
        $controls = Invoke-GraphCollection -Uri $uri

        # Addressed states are compared case-insensitively. A profile is addressed
        # when any controlStateUpdates entry uses one of these states.
        $addressedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($stateName in $AddressedControlStates) {
            $null = $addressedSet.Add($stateName)
        }

        # Join key is controlScore.controlName == secureScoreControlProfile.id.
        # Only the newest secureScore record is authoritative for what applies today.
        $scoreByControlName = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $rawControlScores = $null
        if ($null -ne $LatestSecureScore) {
            $rawControlScores = Get-SafePropertyValue -InputObject $LatestSecureScore -Name 'controlScores'
        }
        foreach ($scoreEntry in @($rawControlScores)) {
            if ($null -eq $scoreEntry) {
                continue
            }
            $scoredName = ([string](Get-SafePropertyValue -InputObject $scoreEntry -Name 'controlName')).Trim()
            if ([string]::IsNullOrWhiteSpace($scoredName)) {
                continue
            }
            if (-not $scoreByControlName.ContainsKey($scoredName)) {
                $scoreByControlName[$scoredName] = $scoreEntry
            }
        }

        $profileById = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $notApplicableCount = 0
        foreach ($control in $controls) {
            $profileId = ([string](Get-SafePropertyValue -InputObject $control -Name 'id')).Trim()
            if ([string]::IsNullOrWhiteSpace($profileId) -or -not $scoreByControlName.ContainsKey($profileId)) {
                $notApplicableCount++
                continue
            }
            if (-not $profileById.ContainsKey($profileId)) {
                $profileById[$profileId] = $control
            }
        }

        $openList = [System.Collections.Generic.List[object]]::new()
        $fullyAchievedCount = 0
        $addressedWithPointsRemaining = 0
        $totalPoints = [double]0

        foreach ($controlName in @($scoreByControlName.Keys)) {
            $scoreEntry = $scoreByControlName[$controlName]
            $matchedProfile = $null
            if ($profileById.ContainsKey($controlName)) {
                $matchedProfile = $profileById[$controlName]
            }

            $currentScore = ConvertTo-InvariantDouble -Value (Get-SafePropertyValue -InputObject $scoreEntry -Name 'score')
            if ($null -eq $currentScore) {
                $currentScore = ConvertTo-InvariantDouble -Value (Get-SafePropertyValue -InputObject $scoreEntry -Name 'currentScore')
            }
            if ($null -eq $currentScore) {
                $currentScore = [double]0
            }

            $maxScore = $null
            if ($null -ne $matchedProfile) {
                $maxScore = ConvertTo-InvariantDouble -Value (Get-SafePropertyValue -InputObject $matchedProfile -Name 'maxScore')
            }
            if ($null -eq $maxScore) {
                $maxScore = ConvertTo-InvariantDouble -Value (Get-SafePropertyValue -InputObject $scoreEntry -Name 'maxScore')
            }
            if ($null -eq $maxScore) {
                $maxScore = [double]0
            }

            # Floor at 0. Secure Score values are reported to 2 decimal places.
            $pointsAvailable = [math]::Round(($maxScore - $currentScore), 2)
            if ($pointsAvailable -lt 0) {
                $pointsAvailable = [double]0
            }
            $totalPoints += $pointsAvailable

            $isAddressed = $false
            if ($null -ne $matchedProfile) {
                $isAddressed = Test-SecureScoreControlAddressed -Control $matchedProfile -AddressedStates $addressedSet
                Add-ControlAssessmentProperty -InputObject $matchedProfile -Name 'score' -Value $currentScore
                Add-ControlAssessmentProperty -InputObject $matchedProfile -Name 'currentScore' -Value $currentScore
                Add-ControlAssessmentProperty -InputObject $matchedProfile -Name 'pointsAvailable' -Value $pointsAvailable
            }

            $hasPointsRemaining = $pointsAvailable -gt 0
            if ($hasPointsRemaining -and -not $isAddressed) {
                if ($null -ne $matchedProfile) {
                    $openList.Add($matchedProfile)
                }
                else {
                    $openList.Add([pscustomobject]@{
                            id                 = $controlName
                            title              = Get-SafePropertyValue -InputObject $scoreEntry -Name 'description'
                            rank               = $null
                            tier               = $null
                            maxScore           = $maxScore
                            implementationCost = $null
                            userImpact         = $null
                            controlCategory    = Get-SafePropertyValue -InputObject $scoreEntry -Name 'controlCategory'
                            score              = $currentScore
                            currentScore       = $currentScore
                            pointsAvailable    = $pointsAvailable
                        })
                }
            }
            elseif ($hasPointsRemaining -and $isAddressed) {
                $addressedWithPointsRemaining++
            }
            else {
                $fullyAchievedCount++
            }
        }

        $openArray = $openList.ToArray()
        $topOpen = @(
            $openArray |
                Sort-Object -Property @{
                    Expression = {
                        $parsedPoints = ConvertTo-InvariantDouble -Value (Get-SafePropertyValue -InputObject $_ -Name 'pointsAvailable')
                        if ($null -eq $parsedPoints) { [double]0 } else { $parsedPoints }
                    }
                    Descending = $true
                }, @{
                    Expression = { Get-SecureScoreRankSortKey -Control $_ }
                    Descending = $false
                } |
                Select-Object -First 25
        )

        return [pscustomobject]@{
            All                          = $controls
            Open                         = $openArray
            TopOpen                      = $topOpen
            ScoredControlCount           = $scoreByControlName.Count
            OpenControlCount             = $openArray.Length
            FullyAchievedCount           = $fullyAchievedCount
            AddressedWithPointsRemaining = $addressedWithPointsRemaining
            NotApplicableCount           = $notApplicableCount
            TotalPointsAvailable         = [math]::Round($totalPoints, 2)
        }
    }

    function Get-ConditionalAccessPolicy {
        [CmdletBinding()]
        [Alias('Get-ConditionalAccessPolicies')]
        [OutputType([object[]])]
        param(
            [Parameter()]
            [string]$BaseUri = $NormalizedGraphBaseUri
        )

        $uri = '{0}/v1.0/identity/conditionalAccess/policies?$top=100&$select=id,displayName,state,createdDateTime,modifiedDateTime,conditions,grantControls,sessionControls' -f $BaseUri
        $policies = Invoke-GraphCollection -Uri $uri

        $sanitized = [System.Collections.Generic.List[object]]::new($policies.Count)
        foreach ($policy in $policies) {
            $sanitized.Add([pscustomobject]@{
                    DisplayName     = Get-SafePropertyValue -InputObject $policy -Name 'displayName'
                    State           = Get-SafePropertyValue -InputObject $policy -Name 'state'
                    CreatedDate     = Get-SafePropertyValue -InputObject $policy -Name 'createdDateTime'
                    ModifiedDate    = Get-SafePropertyValue -InputObject $policy -Name 'modifiedDateTime'
                    Conditions      = Get-SafePropertyValue -InputObject $policy -Name 'conditions'
                    GrantControls   = Get-SafePropertyValue -InputObject $policy -Name 'grantControls'
                    SessionControls = Get-SafePropertyValue -InputObject $policy -Name 'sessionControls'
                    Id              = Get-SafePropertyValue -InputObject $policy -Name 'id'
                })
        }

        $finalPolicies = $sanitized.ToArray()
        Write-Output -InputObject $finalPolicies -NoEnumerate
    }

    function Get-SecurityDefaultsStatus {
        [CmdletBinding()]
        [OutputType([pscustomobject])]
        param(
            [Parameter()]
            [string]$BaseUri = $NormalizedGraphBaseUri
        )

        try {
            $uri = '{0}/v1.0/policies/identitySecurityDefaultsEnforcementPolicy' -f $BaseUri
            $policy = Invoke-MgGraphRequestWithRetry -Uri $uri
            return [pscustomobject]@{
                Policy    = $policy
                Succeeded = $true
            }
        }
        catch {
            Write-Warning "Could not retrieve Security Defaults status (continuing with null): $($_.Exception.Message)"
            return [pscustomobject]@{
                Policy    = $null
                Succeeded = $false
            }
        }
    }

    function Get-MfaRegistrationSummary {
        [CmdletBinding()]
        [OutputType([pscustomobject])]
        param(
            [Parameter()]
            [string]$BaseUri = $NormalizedGraphBaseUri
        )

        # Graph beta userRegistrationActivity(period='d30') returns registration activity
        # events for the period. It is not a point-in-time MFA registration status.
        try {
            $uri = "{0}/beta/reports/authenticationMethods/userRegistrationActivity(period='d30')" -f $BaseUri
            $records = Invoke-GraphCollection -Uri $uri -AllowPartialResults
            return [pscustomobject]@{
                Records   = $records
                Succeeded = $true
            }
        }
        catch {
            Write-Warning "Could not retrieve MFA registration activity events (continuing with empty collection): $($_.Exception.Message)"
            return [pscustomobject]@{
                Records   = @()
                Succeeded = $false
            }
        }
    }

    function ConvertTo-SecureScoreSummary {
        [CmdletBinding()]
        [Alias('Convert-SecureScoreDataSummary')]
        [OutputType([pscustomobject])]
        param(
            [Parameter(Mandatory = $true)]
            [object]$SecurityData
        )

        $historyItems = @(Get-SafePropertyValue -InputObject $SecurityData.SecureScore -Name 'History')
        # Wrap after Sort-Object. A single history date would otherwise unwrap to a scalar
        # and .Count would throw under Set-StrictMode when -Compact is requested.
        $sortedDates = @(
            @(
                foreach ($entry in $historyItems) {
                    $dt = Get-SafePropertyValue -InputObject $entry -Name 'createdDateTime'
                    if ($null -ne $dt) { $dt }
                }
            ) | Sort-Object
        )

        $oldestDate = if ($sortedDates.Count -gt 0) { $sortedDates[0] } else { $null }
        $newestDate = if ($sortedDates.Count -gt 0) { $sortedDates[$sortedDates.Count - 1] } else { $null }

        $topOpenProjected = [System.Collections.Generic.List[object]]::new()
        foreach ($ctrl in @(Get-SafePropertyValue -InputObject $SecurityData.SecureScoreControls -Name 'TopOpen')) {
            $topOpenProjected.Add([pscustomobject]@{
                    id                 = Get-SafePropertyValue -InputObject $ctrl -Name 'id'
                    title              = Get-SafePropertyValue -InputObject $ctrl -Name 'title'
                    rank               = Get-SafePropertyValue -InputObject $ctrl -Name 'rank'
                    implementationCost = ConvertTo-NormalizedCostImpact -Value (Get-SafePropertyValue -InputObject $ctrl -Name 'implementationCost')
                    userImpact         = ConvertTo-NormalizedCostImpact -Value (Get-SafePropertyValue -InputObject $ctrl -Name 'userImpact')
                    threats            = Get-SafePropertyValue -InputObject $ctrl -Name 'threats'
                    tier               = Get-SafePropertyValue -InputObject $ctrl -Name 'tier'
                    remediation        = Get-SafePropertyValue -InputObject $ctrl -Name 'remediation'
                    controlCategory    = Get-SafePropertyValue -InputObject $ctrl -Name 'controlCategory'
                    actionUrl          = Get-SafePropertyValue -InputObject $ctrl -Name 'actionUrl'
                    maxScore           = Get-SafePropertyValue -InputObject $ctrl -Name 'maxScore'
                    score              = Get-SafePropertyValue -InputObject $ctrl -Name 'score'
                    currentScore       = Get-SafePropertyValue -InputObject $ctrl -Name 'currentScore'
                    pointsAvailable    = Get-SafePropertyValue -InputObject $ctrl -Name 'pointsAvailable'
                })
        }

        $caCompact = @(
            $SecurityData.ConditionalAccess |
                Select-Object -Property DisplayName, State, CreatedDate, ModifiedDate, Conditions, GrantControls, SessionControls -First 50
        )

        return [pscustomobject]@{
            Tenant                 = $SecurityData.Tenant
            TenantId               = $SecurityData.TenantId
            CollectionDate         = $SecurityData.CollectionDate
            CollectedBy            = $SecurityData.CollectedBy
            SecureScore            = @{
                Latest         = $SecurityData.SecureScore.Latest
                HistorySummary = @{
                    Count      = $historyItems.Count
                    OldestDate = $oldestDate
                    NewestDate = $newestDate
                }
            }
            SecureScoreControls    = @{
                TotalCount                   = @($SecurityData.SecureScoreControls.All).Count
                OpenCount                    = @($SecurityData.SecureScoreControls.Open).Count
                ScoredControlCount           = $SecurityData.SecureScoreControls.ScoredControlCount
                OpenControlCount             = $SecurityData.SecureScoreControls.OpenControlCount
                FullyAchievedCount           = $SecurityData.SecureScoreControls.FullyAchievedCount
                AddressedWithPointsRemaining = $SecurityData.SecureScoreControls.AddressedWithPointsRemaining
                NotApplicableCount           = $SecurityData.SecureScoreControls.NotApplicableCount
                TotalPointsAvailable         = $SecurityData.SecureScoreControls.TotalPointsAvailable
                TopOpen                      = $topOpenProjected.ToArray()
            }
            ConditionalAccess      = $caCompact
            SecurityDefaults       = $SecurityData.SecurityDefaults
            MfaRegistrationSummary = $SecurityData.MfaRegistrationSummary
            MfaRegistrationSummaryDescription = Get-SafePropertyValue -InputObject $SecurityData -Name 'MfaRegistrationSummaryDescription'
        }
    }
    #endregion Helper functions

    # Start optional session transcript when -TranscriptPath or -Debug is provided.
    $effectiveTranscriptPath = if (-not [string]::IsNullOrWhiteSpace($TranscriptPath)) {
        $TranscriptPath
    }
    elseif ($PSBoundParameters.ContainsKey('Debug')) {
        Join-Path -Path (Get-Location).Path -ChildPath ('o365-secure-score-extract_{0}.log' -f (Get-Date).ToString('yyyyMMdd_HHmmss'))
    }
    else {
        $null
    }

    if (-not [string]::IsNullOrWhiteSpace($effectiveTranscriptPath)) {
        if ($PSCmdlet.ShouldProcess($effectiveTranscriptPath, 'Start session transcript')) {
            $transcriptParent = [System.IO.Path]::GetDirectoryName($effectiveTranscriptPath)
            if (-not [string]::IsNullOrWhiteSpace($transcriptParent) -and -not (Test-Path -LiteralPath $transcriptParent -PathType Container)) {
                New-Item -ItemType Directory -Path $transcriptParent -Force -ErrorAction Stop | Out-Null
            }
            Start-Transcript -Path $effectiveTranscriptPath -Append -ErrorAction Stop | Out-Null
            $transcriptStarted = $true
        }
    }
}

process {
    #region Main logic
    $tenantIndex = $processedTenants + $skippedTenants + $failedTenants + 1
    $targetTenant = $TenantDomain.Trim()
    $resolvedDataFile = Resolve-TenantDataFilePath -TargetTenant $targetTenant -RequestedPath $DataFile -TenantSequence $tenantIndex
    $resolvedCompactFile = if ($Compact) { Get-OutputPathWithSuffix -Path $resolvedDataFile -Suffix '_compact' } else { $null }

    $actionDescription = if ($Compact) {
        "Extract M365 Secure Score data to '$resolvedDataFile' and '$resolvedCompactFile'"
    }
    else {
        "Extract M365 Secure Score data to '$resolvedDataFile'"
    }

    if (-not $PSCmdlet.ShouldProcess($targetTenant, $actionDescription)) {
        $skippedTenants++
        return
    }

    $tenantStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $stagesSucceeded = 0
    $stagesDegraded = 0

    try {
        Write-StatusMessage -Level Section -Message '==============================================================='
        Write-StatusMessage -Level Section -Message "  M365 Secure Score Data Extraction - $targetTenant"
        Write-StatusMessage -Level Section -Message '==============================================================='
        Write-StatusMessage -Message "Output File: $resolvedDataFile"
        Write-StatusMessage -Message "Score History Records: $ScoreHistoryCount"

        if (-not $moduleChecked) {
            $importCommand = Get-Command -Name Import-Module -ErrorAction SilentlyContinue
            $isRealImportCmdlet = $null -ne $importCommand -and $importCommand.CommandType -eq [System.Management.Automation.CommandTypes]::Cmdlet
            if ($isRealImportCmdlet -and -not (Get-Module -Name $RequiredModuleName) -and -not (Get-Module -ListAvailable -Name $RequiredModuleName)) {
                throw [System.InvalidOperationException]::new(
                    "Required module '$RequiredModuleName' is not installed. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
                )
            }
            $moduleChecked = $true
        }

        Write-Progress -Id 1 -Activity $ProgressActivity -Status "Authenticating to '$targetTenant'..." -PercentComplete 5
        $context = Connect-TargetTenantGraph -TargetTenant $targetTenant -Scopes $RequiredScopes -BaseUri $NormalizedGraphBaseUri -Force:$ForceReconnect

        $securityData = [pscustomobject]@{
            Tenant                            = $targetTenant
            CollectionDate                    = $null
            TenantId                          = $context.TenantId
            CollectedBy                       = Get-GraphPrincipalLabel -Context $context
            SecureScore                       = $null
            SecureScoreControls               = $null
            ConditionalAccess                 = $null
            SecurityDefaults                  = $null
            MfaRegistrationSummary            = $null
            MfaRegistrationSummaryDescription = $MfaRegistrationActivityDescription
        }

        # Stage 1: Secure Score
        Write-Progress -Id 1 -Activity $ProgressActivity -Status '[1/5] Collecting Secure Score history (typically 5-30s)...' -PercentComplete 15
        Write-StatusMessage -Message '[1/5] Collecting Secure Score data (typically 5-30s depending on tenant size)...'
        $stageTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $securityData.SecureScore = Get-SecureScoreData -HistoryCount $ScoreHistoryCount -BaseUri $NormalizedGraphBaseUri
        $stageTimer.Stop()
        $stagesSucceeded++
        Write-StatusMessage -Level Detail -Message ("Completed Secure Score history ({0} records) in {1:N1}s." -f @($securityData.SecureScore.History).Count, $stageTimer.Elapsed.TotalSeconds)
        Save-PartialData -SecurityData $securityData -OutputPath $resolvedDataFile -Stage 'secure-score'

        # Stage 2: Secure Score Control Profiles
        Write-Progress -Id 1 -Activity $ProgressActivity -Status '[2/5] Collecting Secure Score control profiles...' -PercentComplete 35
        Write-StatusMessage -Message '[2/5] Collecting Secure Score control profiles...'
        $stageTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $securityData.SecureScoreControls = Get-SecureScoreControlProfile -BaseUri $NormalizedGraphBaseUri -LatestSecureScore $securityData.SecureScore.Latest
        $stageTimer.Stop()
        $stagesSucceeded++
        $controlSummary = $securityData.SecureScoreControls
        Write-StatusMessage -Level Detail -Message (
            'Retrieved {0} catalog controls ({1} scored, {2} open, {3} fully achieved, {4} addressed with points remaining, {5} not applicable, {6} points available) in {7:N1}s.' -f @(
                @($controlSummary.All).Count,
                $controlSummary.ScoredControlCount,
                $controlSummary.OpenControlCount,
                $controlSummary.FullyAchievedCount,
                $controlSummary.AddressedWithPointsRemaining,
                $controlSummary.NotApplicableCount,
                $controlSummary.TotalPointsAvailable,
                $stageTimer.Elapsed.TotalSeconds
            )
        )
        Save-PartialData -SecurityData $securityData -OutputPath $resolvedDataFile -Stage 'controls'

        # Stage 3: Conditional Access Policies
        Write-Progress -Id 1 -Activity $ProgressActivity -Status '[3/5] Collecting Conditional Access policies...' -PercentComplete 55
        Write-StatusMessage -Message '[3/5] Collecting Conditional Access policies...'
        $stageTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $securityData.ConditionalAccess = Get-ConditionalAccessPolicy -BaseUri $NormalizedGraphBaseUri
        $stageTimer.Stop()
        $stagesSucceeded++
        Write-StatusMessage -Level Detail -Message ("Retrieved {0} Conditional Access policies in {1:N1}s." -f @($securityData.ConditionalAccess).Count, $stageTimer.Elapsed.TotalSeconds)
        Save-PartialData -SecurityData $securityData -OutputPath $resolvedDataFile -Stage 'conditional-access'

        # Stage 4: Security Defaults
        Write-Progress -Id 1 -Activity $ProgressActivity -Status '[4/5] Checking Identity Security Defaults status...' -PercentComplete 75
        Write-StatusMessage -Message '[4/5] Checking Identity Security Defaults status...'
        $stageTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $secDefaultsOutcome = Get-SecurityDefaultsStatus -BaseUri $NormalizedGraphBaseUri
        $securityData.SecurityDefaults = $secDefaultsOutcome.Policy
        $stageTimer.Stop()
        if ($secDefaultsOutcome.Succeeded) { $stagesSucceeded++ } else { $stagesDegraded++ }
        Write-StatusMessage -Level Detail -Message ("Completed Security Defaults check in {0:N1}s." -f $stageTimer.Elapsed.TotalSeconds)
        Save-PartialData -SecurityData $securityData -OutputPath $resolvedDataFile -Stage 'security-defaults'

        # Stage 5: MFA registration activity events (not a registration-status snapshot)
        Write-Progress -Id 1 -Activity $ProgressActivity -Status '[5/5] Collecting MFA registration activity (events)...' -PercentComplete 88
        Write-StatusMessage -Message '[5/5] Collecting MFA registration activity (events)...'
        $stageTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $mfaOutcome = Get-MfaRegistrationSummary -BaseUri $NormalizedGraphBaseUri
        $securityData.MfaRegistrationSummary = $mfaOutcome.Records
        $stageTimer.Stop()
        if ($mfaOutcome.Succeeded) { $stagesSucceeded++ } else { $stagesDegraded++ }
        Write-StatusMessage -Level Detail -Message ("Retrieved {0} MFA registration activity events in {1:N1}s." -f @($securityData.MfaRegistrationSummary).Count, $stageTimer.Elapsed.TotalSeconds)

        # Finalize and write JSON output(s) - skips a redundant 5th _partial disk write
        # because the final file is written immediately and _partial is removed next.
        Write-Progress -Id 1 -Activity $ProgressActivity -Status 'Writing JSON export files...' -PercentComplete 96
        $securityData.CollectionDate = (Get-Date).ToString('o')

        $fullBytes = Write-JsonFile -InputObject $securityData -FilePath $resolvedDataFile -Depth 12
        $fullSizeKB = [math]::Round($fullBytes / 1KB, 2)
        Write-StatusMessage -Message "Saved full JSON export: $resolvedDataFile ($fullSizeKB KB)"

        $compactPathOut = $null
        $compactSizeKB = $null
        if ($Compact) {
            try {
                $summarizedData = ConvertTo-SecureScoreSummary -SecurityData $securityData
                $compactBytes = Write-JsonFile -InputObject $summarizedData -FilePath $resolvedCompactFile -Depth 12
                $compactSizeKB = [math]::Round($compactBytes / 1KB, 2)
                $reductionPct = if ($fullBytes -gt 0) { [math]::Round((1 - ($compactBytes / $fullBytes)) * 100, 1) } else { 0 }
                $compactPathOut = $resolvedCompactFile
                Write-StatusMessage -Message "Saved compact JSON export: $resolvedCompactFile ($compactSizeKB KB, $reductionPct% smaller)"
            }
            catch {
                $stagesDegraded++
                Write-Warning "Failed to create compact summary file '$resolvedCompactFile' (full export is intact): $($_.Exception.Message)"
            }
        }

        # Remove temporary recovery checkpoint now that full JSON export succeeded.
        $partialFile = Get-OutputPathWithSuffix -Path $resolvedDataFile -Suffix '_partial'
        if (Test-Path -LiteralPath $partialFile -PathType Leaf) {
            Remove-Item -LiteralPath $partialFile -Force -ErrorAction SilentlyContinue
            Write-Debug "[Main] Removed temporary recovery file: $partialFile"
        }

        $tenantStopwatch.Stop()
        $processedTenants++

        $latestObj = $securityData.SecureScore.Latest
        $latestCurrentScore = Get-SafePropertyValue -InputObject $latestObj -Name 'currentScore'
        $latestMaxScore = Get-SafePropertyValue -InputObject $latestObj -Name 'maxScore'
        $secDefaultsEnabled = Get-SafePropertyValue -InputObject $securityData.SecurityDefaults -Name 'isEnabled'

        [pscustomobject]@{
            Tenant                       = $securityData.Tenant
            TenantId                     = $securityData.TenantId
            CollectedBy                  = $securityData.CollectedBy
            CollectionDate               = $securityData.CollectionDate
            FullDataFile                 = $resolvedDataFile
            FullSizeKB                   = $fullSizeKB
            CompactDataFile              = $compactPathOut
            CompactSizeKB                = $compactSizeKB
            LatestScore                  = $latestCurrentScore
            MaxScore                     = $latestMaxScore
            OpenControlCount             = @($securityData.SecureScoreControls.Open).Count
            ScoredControlCount           = $securityData.SecureScoreControls.ScoredControlCount
            FullyAchievedCount           = $securityData.SecureScoreControls.FullyAchievedCount
            AddressedWithPointsRemaining = $securityData.SecureScoreControls.AddressedWithPointsRemaining
            NotApplicableCount           = $securityData.SecureScoreControls.NotApplicableCount
            TotalPointsAvailable         = $securityData.SecureScoreControls.TotalPointsAvailable
            ConditionalAccessCount       = @($securityData.ConditionalAccess).Count
            SecurityDefaultsEnabled      = $secDefaultsEnabled
            MfaRecordCount               = @($securityData.MfaRegistrationSummary).Count
            StagesSucceeded              = $stagesSucceeded
            StagesDegraded               = $stagesDegraded
            ElapsedSeconds               = [math]::Round($tenantStopwatch.Elapsed.TotalSeconds, 2)
        }
    }
    catch {
        $tenantStopwatch.Stop()
        $failedTenants++
        $partialHint = Get-OutputPathWithSuffix -Path $resolvedDataFile -Suffix '_partial'
        $recoveryNote = if (Test-Path -LiteralPath $partialHint -PathType Leaf) {
            " Partial recovery checkpoint preserved at '$partialHint'."
        }
        else {
            ''
        }
        Write-Warning "Extraction failed for tenant '$targetTenant' after $([math]::Round($tenantStopwatch.Elapsed.TotalSeconds, 1))s.$recoveryNote Error: $($_.Exception.Message)"
        $global:LASTEXITCODE = 1
        throw
    }
    finally {
        Write-Progress -Id 1 -Activity $ProgressActivity -Completed
    }
    #endregion Main logic
}

end {
    #region Cleanup
    $totalStopwatch.Stop()
    Write-StatusMessage -Message ("Extraction summary - Processed: {0}; Skipped: {1}; Failed: {2}; Total elapsed: {3:N1}s." -f $processedTenants, $skippedTenants, $failedTenants, $totalStopwatch.Elapsed.TotalSeconds)
    if ($failedTenants -eq 0) {
        $global:LASTEXITCODE = 0
    }
    else {
        $global:LASTEXITCODE = 1
    }

    if ($transcriptStarted) {
        try {
            Stop-Transcript -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Warning "Could not stop session transcript: $($_.Exception.Message)"
        }
    }
    #endregion Cleanup
}
