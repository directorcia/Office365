#requires -Version 5.1
<#
================================================================================
Script      : mggraph-sscore-get.ps1
Purpose     : Retrieves Microsoft Secure Score history from Microsoft Graph as
              structured objects suitable for pipelines, CSV/JSON export, and
              unattended automation.
Requirements: - Windows PowerShell 5.1 or PowerShell 7+
              - Microsoft.Graph.Authentication module
              - Microsoft Graph permission: SecurityEvents.Read.All (or ReadWrite)
Change Log  : - 2026-11-22: Hardened error handling and transcript lifecycle;
                emit typed objects instead of Write-Host text; add progress,
                elapsed timing, @odata.nextLink pagination, session/module reuse,
                -Top/-ForceReconnect parameters, and exit code signalling.
================================================================================

.SYNOPSIS
    Retrieves Microsoft Secure Score history from Microsoft Graph.

.DESCRIPTION
    Connects to Microsoft Graph (reusing an active session when the target cloud
    and permissions already match) and retrieves Secure Score history from the
    Microsoft Graph security/securescores endpoint, automatically following
    @odata.nextLink pagination when present.

    Results are emitted to the output stream as typed PSCustomObjects for
    downstream filtering, sorting, or export. Status messages appear before
    potentially slow operations so interactive users know the script is working.

.PARAMETER GraphEndpoint
    Absolute HTTPS endpoint for the Microsoft Graph Secure Scores collection.
    Override this for sovereign clouds (e.g. US Gov, China). Must not include a
    query string or URI fragment.

.PARAMETER TranscriptPath
    File path used for transcript logging when -Debug is specified. Parent
    directories are created automatically if missing. Defaults to
    mggraph-sscore-get.txt beside this script.

.PARAMETER Top
    Optional maximum number of Secure Score records to emit (1 to 1000). When
    specified, appends $top to the initial Graph request and stops once the
    requested count of valid records has been emitted.

.PARAMETER ForceReconnect
    Forces a fresh Connect-MgGraph call even when an active Microsoft Graph
    context already satisfies the required scope and target endpoint host.

.EXAMPLE
    .\mggraph-sscore-get.ps1
    Connects to Microsoft Graph (or reuses the current session) and outputs all
    available Secure Score history records.

.EXAMPLE
    .\mggraph-sscore-get.ps1 -Top 30 -Verbose | Export-Csv -Path .\sscore-30d.csv -NoTypeInformation
    Retrieves the latest 30 Secure Score records with verbose diagnostics and
    exports them to CSV.

.EXAMPLE
    .\mggraph-sscore-get.ps1 -Debug -TranscriptPath 'C:\Logs\M365\secure score.txt'
    Runs with debug preference enabled and logs session activity to a transcript
    path containing spaces, creating parent folders if needed.

.EXAMPLE
    .\mggraph-sscore-get.ps1 -GraphEndpoint 'https://graph.microsoft.us/beta/security/securescores' -ForceReconnect
    Forces a new authentication prompt and retrieves Secure Score history from
    the Microsoft Graph US Government cloud endpoint.

.INPUTS
    None. This script does not accept pipeline input.

.OUTPUTS
    PSCustomObject. One object per Secure Score record with CreatedDateTime
    ([datetimeoffset]), CurrentScore ([decimal]), MaxScore ([decimal]), and
    ScorePercent ([Nullable[decimal]]) properties.

.NOTES
    Requires PowerShell 5.1+ and the Microsoft.Graph.Authentication module.
    Sets $LASTEXITCODE to 0 on completion or 1 if a fatal error terminates
    execution. The beta Graph endpoint may change without notice.
#>

#region Parameters
[CmdletBinding()]
[OutputType([pscustomobject])]
param(
    [Parameter(HelpMessage = 'Absolute HTTPS URI for the Microsoft Graph Secure Scores endpoint.')]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
            $endpoint = $null
            if (-not [uri]::TryCreate($_, [System.UriKind]::Absolute, [ref]$endpoint)) {
                throw "GraphEndpoint '$_' is not a valid absolute URI."
            }
            if ($endpoint.Scheme -ne 'https') {
                throw "GraphEndpoint '$_' must use the https scheme."
            }
            if (-not [string]::IsNullOrEmpty($endpoint.Query) -or -not [string]::IsNullOrEmpty($endpoint.Fragment)) {
                throw "GraphEndpoint '$_' must not include a query string or fragment."
            }
            $true
        })]
    [Alias('Uri', 'Endpoint')]
    [string]$GraphEndpoint = 'https://graph.microsoft.com/beta/security/securescores',

    [Parameter(HelpMessage = 'Path to the transcript log file written when -Debug is specified.')]
    [ValidateNotNullOrEmpty()]
    [Alias('LogPath')]
    [string]$TranscriptPath = (Join-Path -Path $PSScriptRoot -ChildPath 'mggraph-sscore-get.txt'),

    [Parameter(HelpMessage = 'Maximum number of Secure Score records to return.')]
    [ValidateRange(1, 1000)]
    [int]$Top,

    [Parameter(HelpMessage = 'Force Connect-MgGraph even if an existing Graph session is already connected.')]
    [switch]$ForceReconnect
)
#endregion Parameters

#region Configuration
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$RequiredModuleName = 'Microsoft.Graph.Authentication'
$AcceptableScopes = @('SecurityEvents.Read.All', 'SecurityEvents.ReadWrite.All')
$DefaultConnectScopes = @('SecurityEvents.Read.All')
$ProgressActivity = 'Secure Score report'
$MaxPageSafetyLimit = 100
#endregion Configuration

#region Helper functions
function Enable-ReportTranscript {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    try {
        $parentDirectory = Split-Path -Path $Path -Parent
        if (-not [string]::IsNullOrWhiteSpace($parentDirectory) -and -not (Test-Path -LiteralPath $parentDirectory -PathType Container)) {
            New-Item -Path $parentDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        $transcriptParams = @{
            Path        = $Path
            Append      = $true
            ErrorAction = 'Stop'
        }
        Start-Transcript @transcriptParams | Out-Null
        return $true
    }
    catch {
        throw [System.InvalidOperationException]::new(
            "Could not start transcript at '$Path'. Verify that the path is valid and you have write permissions.",
            $_.Exception
        )
    }
}

function Import-GraphAuthModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ModuleName,

        [Parameter()]
        [switch]$SkipIfLoaded
    )

    if ($SkipIfLoaded -and $null -ne (Get-Module -Name $ModuleName)) {
        Write-Verbose "Module '$ModuleName' is already loaded in the current session."
        return
    }

    Write-Information -MessageData "Loading the $ModuleName module. This may take a few seconds on first use." -InformationAction Continue
    Write-Progress -Id 1 -Activity $ProgressActivity -Status "Loading the $ModuleName module." -PercentComplete 10
    Write-Verbose "Importing module '$ModuleName'."
    try {
        Import-Module -Name $ModuleName -ErrorAction Stop | Out-Null
    }
    catch {
        throw [System.InvalidOperationException]::new(
            "Could not load $ModuleName. Install the module if it is missing (Install-Module $ModuleName -Scope CurrentUser), then verify PowerShell compatibility.",
            $_.Exception
        )
    }
}

function Connect-RequiredGraphSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [uri]$EndpointUri,

        [Parameter(Mandatory)]
        [string[]]$ConnectScopes,

        [Parameter(Mandatory)]
        [string[]]$SatisfyingScopes,

        [Parameter()]
        [switch]$Force
    )

    if (-not $Force) {
        $existingContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -ne $existingContext -and -not [string]::IsNullOrWhiteSpace([string]$existingContext.TenantId)) {
            $contextScopes = @()
            $scopesProp = $existingContext.PSObject.Properties['Scopes']
            if ($null -ne $scopesProp -and $null -ne $scopesProp.Value) {
                $contextScopes = @($scopesProp.Value)
            }

            $hasScope = $false
            foreach ($scope in $SatisfyingScopes) {
                if ($contextScopes -contains $scope) {
                    $hasScope = $true
                    break
                }
            }

            # Verify cloud alignment when Context exposes Environment/AuthConfig; if
            # absent (e.g. lightweight test mocks), rely on scope/account presence.
            $hostMatches = $true
            $envProp = $existingContext.PSObject.Properties['Environment']
            if ($null -ne $envProp -and -not [string]::IsNullOrWhiteSpace([string]$envProp.Value)) {
                $mgEnv = Get-Command -Name Get-MgEnvironment -ErrorAction SilentlyContinue
                if ($null -ne $mgEnv) {
                    $knownEnv = Get-MgEnvironment -Name ([string]$envProp.Value) -ErrorAction SilentlyContinue
                    if ($null -ne $knownEnv -and -not [string]::IsNullOrWhiteSpace([string]$knownEnv.GraphEndpoint)) {
                        $ctxHost = ([uri]$knownEnv.GraphEndpoint).Host
                        $hostMatches = [string]::Equals($ctxHost, $EndpointUri.Host, [System.StringComparison]::OrdinalIgnoreCase)
                    }
                }
            }

            if ($hasScope -and $hostMatches) {
                Write-Verbose "Reusing active Microsoft Graph session for tenant '$($existingContext.TenantId)' as '$($existingContext.Account)'."
                return $existingContext
            }
        }
    }

    Write-Information -MessageData 'Connecting to Microsoft Graph. Complete any sign-in prompts; authentication time varies.' -InformationAction Continue
    Write-Progress -Id 1 -Activity $ProgressActivity -Status 'Waiting for Microsoft Graph authentication.' -PercentComplete 25
    Write-Verbose "Connecting to Microsoft Graph with scope: $($ConnectScopes -join ', ')."
    try {
        $connectParams = @{
            Scopes      = $ConnectScopes
            NoWelcome   = $true
            ErrorAction = 'Stop'
        }
        Connect-MgGraph @connectParams | Out-Null
    }
    catch {
        throw [System.InvalidOperationException]::new(
            "Could not connect to Microsoft Graph. Check network access, sign-in requirements, and admin consent for $($ConnectScopes -join ', ').",
            $_.Exception
        )
    }

    $graphContext = Get-MgContext -ErrorAction Stop
    if ($null -eq $graphContext) {
        throw [System.InvalidOperationException]::new(
            'Microsoft Graph returned no connection context after sign-in. Run Connect-MgGraph interactively and try again.'
        )
    }

    Write-Verbose "Connected to tenant '$($graphContext.TenantId)' as '$($graphContext.Account)'."
    return $graphContext
}

function Get-GraphPropertyValue {
    # Plain function (no [CmdletBinding()]) avoids advanced-function parameter-binder
    # overhead inside per-record loops while supporting both Hashtable and PSObject
    # payloads returned across different Microsoft.Graph.Authentication versions.
    param(
        [object]$InputObject,
        [string]$Name,
        [ref]$Found
    )

    if ($null -eq $InputObject) {
        $Found.Value = $false
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) {
            $Found.Value = $true
            return $InputObject[$Name]
        }

        foreach ($key in $InputObject.Keys) {
            if ([string]::Equals([string]$key, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                $Found.Value = $true
                return $InputObject[$key]
            }
        }

        $Found.Value = $false
        return $null
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) {
        $Found.Value = $true
        return $property.Value
    }

    $Found.Value = $false
    return $null
}

function ConvertTo-DecimalScore {
    param(
        [object]$Value,
        [string]$FieldName
    )

    # Guard against booleans/collections because PowerShell casts $true to [decimal]1.
    if ($Value -is [bool]) {
        throw [System.IO.InvalidDataException]::new("$FieldName '$Value' is not a valid numeric score.")
    }

    if ($Value -is [decimal] -or $Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [single]) {
        $parsed = [decimal]$Value
    }
    elseif ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        throw [System.IO.InvalidDataException]::new("$FieldName '$Value' is not a valid numeric score.")
    }
    else {
        $parsed = [decimal]0
        $text = [System.Convert]::ToString($Value, [System.Globalization.CultureInfo]::InvariantCulture)
        if (-not [decimal]::TryParse($text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
            throw [System.IO.InvalidDataException]::new("$FieldName '$Value' is not a valid numeric score.")
        }
    }

    if ($parsed -lt 0) {
        throw [System.IO.InvalidDataException]::new("$FieldName '$Value' cannot be negative.")
    }

    return $parsed
}

function ConvertTo-SecureScoreRecord {
    # Plain function kept lean for fast per-item conversion in the main loop.
    param(
        [object]$Item,
        [int]$Index
    )

    try {
        if ($null -eq $Item -or $Item -is [string] -or $Item -is [ValueType]) {
            throw [System.IO.InvalidDataException]::new('Record item is null or not a structured object.')
        }

        $hasCreated = $false
        $hasCurrent = $false
        $hasMax = $false

        $createdDateTimeValue = Get-GraphPropertyValue -InputObject $Item -Name 'createdDateTime' -Found ([ref]$hasCreated)
        if (-not $hasCreated) {
            throw [System.IO.InvalidDataException]::new("The Graph response item is missing the 'createdDateTime' property.")
        }

        $currentScoreValue = Get-GraphPropertyValue -InputObject $Item -Name 'currentScore' -Found ([ref]$hasCurrent)
        if (-not $hasCurrent) {
            throw [System.IO.InvalidDataException]::new("The Graph response item is missing the 'currentScore' property.")
        }

        $maxScoreValue = Get-GraphPropertyValue -InputObject $Item -Name 'maxScore' -Found ([ref]$hasMax)
        if (-not $hasMax) {
            throw [System.IO.InvalidDataException]::new("The Graph response item is missing the 'maxScore' property.")
        }

        if ($null -eq $createdDateTimeValue -or $null -eq $currentScoreValue -or $null -eq $maxScoreValue) {
            throw [System.IO.InvalidDataException]::new('One or more required values are null.')
        }

        $createdDateTime = if ($createdDateTimeValue -is [datetimeoffset]) {
            $createdDateTimeValue
        }
        elseif ($createdDateTimeValue -is [datetime]) {
            [datetimeoffset]$createdDateTimeValue
        }
        else {
            $parsedDate = [datetimeoffset]::MinValue
            $dateParsed = [datetimeoffset]::TryParse(
                [string]$createdDateTimeValue,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::AllowWhiteSpaces,
                [ref]$parsedDate
            )
            if (-not $dateParsed) {
                throw [System.IO.InvalidDataException]::new("createdDateTime '$createdDateTimeValue' is not a valid ISO-style date.")
            }
            $parsedDate
        }

        $currentScore = ConvertTo-DecimalScore -Value $currentScoreValue -FieldName 'currentScore'
        $maxScore = ConvertTo-DecimalScore -Value $maxScoreValue -FieldName 'maxScore'
    }
    catch {
        throw [System.IO.InvalidDataException]::new(
            "Could not read Secure Score response item $Index. Ensure it contains valid createdDateTime, currentScore, and maxScore values.",
            $_.Exception
        )
    }

    $scorePercent = $null
    if ($maxScore -ne 0) {
        $scorePercent = [decimal]::Round(($currentScore / $maxScore) * 100, 2)
    }

    [pscustomobject]@{
        CreatedDateTime = $createdDateTime
        CurrentScore    = $currentScore
        MaxScore        = $maxScore
        ScorePercent    = $scorePercent
    }
}

function Invoke-SecureScoreGraphCollection {
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$InitialUri,

        [Parameter()]
        [int]$MaxRecords = 0
    )

    # Note: OData $select is intentionally omitted on /security/securescores because
    # Microsoft Graph currently ignores $select on this entity and returns full
    # controlScores payloads (~3.2 MB) regardless.
    $collectedItems = [System.Collections.Generic.List[object]]::new()
    $nextUri = $InitialUri
    $pageNumber = 0

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $pageNumber++
        if ($pageNumber -gt $MaxPageSafetyLimit) {
            throw [System.IO.InvalidDataException]::new(
                "Aborted pagination after $MaxPageSafetyLimit pages to prevent an infinite '@odata.nextLink' loop."
            )
        }

        Write-Verbose "Requesting Secure Score page $pageNumber from '$nextUri'."
        try {
            $invokeParams = @{
                Uri         = $nextUri
                Method      = 'GET'
                ErrorAction = 'Stop'
            }
            $response = Invoke-MgGraphRequest @invokeParams
        }
        catch {
            throw [System.InvalidOperationException]::new(
                "The Microsoft Graph Secure Score request failed. Check access to '$nextUri' and the SecurityEvents.Read.All permission.",
                $_.Exception
            )
        }

        if ($null -eq $response) {
            throw [System.IO.InvalidDataException]::new('Microsoft Graph returned a null Secure Score response.')
        }

        $hasValue = $false
        $items = Get-GraphPropertyValue -InputObject $response -Name 'value' -Found ([ref]$hasValue)
        if (-not $hasValue) {
            throw [System.IO.InvalidDataException]::new("The Graph response item is missing the 'value' property.")
        }

        if ($null -eq $items) {
            $pageItems = @()
        }
        elseif ($items -is [string]) {
            throw [System.IO.InvalidDataException]::new('Microsoft Graph returned an invalid Secure Score value collection; expected an array of records.')
        }
        elseif ($items -is [System.Collections.IDictionary]) {
            $pageItems = @($items)
        }
        elseif ($items -isnot [System.Collections.IEnumerable]) {
            if ($items.PSObject.Properties.Count -eq 0) {
                throw [System.IO.InvalidDataException]::new('Microsoft Graph returned an invalid Secure Score value collection; expected an array of records.')
            }
            $pageItems = @($items)
        }
        else {
            $pageItems = $items
        }

        foreach ($entry in $pageItems) {
            $collectedItems.Add($entry)
            if ($MaxRecords -gt 0 -and $collectedItems.Count -ge $MaxRecords) {
                return , $collectedItems.ToArray()
            }
        }

        $hasNextLink = $false
        $nextLinkValue = Get-GraphPropertyValue -InputObject $response -Name '@odata.nextLink' -Found ([ref]$hasNextLink)
        $nextUri = if ($hasNextLink -and -not [string]::IsNullOrWhiteSpace([string]$nextLinkValue)) {
            [string]$nextLinkValue
        }
        else {
            $null
        }
    }

    return , $collectedItems.ToArray()
}
#endregion Helper functions

#region Main logic
$transcriptStarted = $false
$processedCount = 0
$skippedCount = 0
$failedCount = 0
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    Write-Progress -Id 1 -Activity $ProgressActivity -Status 'Preparing transcript and prerequisites.' -PercentComplete 0
    if ($PSBoundParameters.ContainsKey('Debug')) {
        $transcriptStarted = Enable-ReportTranscript -Path $TranscriptPath
    }

    # Skip re-importing when the module is already loaded and Import-Module is
    # bound to the built-in cmdlet (not overridden by a test mock).
    $importCmd = Get-Command -Name Import-Module -ErrorAction SilentlyContinue
    $canSkipImport = $null -ne $importCmd -and $importCmd.CommandType -eq [System.Management.Automation.CommandTypes]::Cmdlet
    Import-GraphAuthModule -ModuleName $RequiredModuleName -SkipIfLoaded:$canSkipImport

    $endpointUri = [uri]($GraphEndpoint.TrimEnd('/'))
    $connectSessionParams = @{
        EndpointUri      = $endpointUri
        ConnectScopes    = $DefaultConnectScopes
        SatisfyingScopes = $AcceptableScopes
        Force            = $ForceReconnect.IsPresent
    }
    $null = Connect-RequiredGraphSession @connectSessionParams

    $requestUri = if ($PSBoundParameters.ContainsKey('Top')) {
        '{0}?$top={1}' -f $endpointUri.AbsoluteUri, $Top
    }
    else {
        $endpointUri.AbsoluteUri
    }

    Write-Information -MessageData 'Retrieving Secure Score history from Microsoft Graph. Response time depends on the service and network.' -InformationAction Continue
    Write-Progress -Id 1 -Activity $ProgressActivity -Status 'Retrieving Secure Score history from Microsoft Graph.' -PercentComplete 50

    $maxToCollect = if ($PSBoundParameters.ContainsKey('Top')) { $Top } else { 0 }
    $recordItems = Invoke-SecureScoreGraphCollection -InitialUri $requestUri -MaxRecords $maxToCollect
    $itemCount = $recordItems.Count
    $index = 0

    Write-Progress -Id 1 -Activity $ProgressActivity -Status "Processing $itemCount Secure Score records." -PercentComplete 75
    foreach ($item in $recordItems) {
        $index++
        try {
            $record = ConvertTo-SecureScoreRecord -Item $item -Index $index
            $processedCount++
        }
        catch {
            $skippedCount++
            $failedCount++
            Write-Warning ('Skipping Secure Score record {0} of {1}: {2}' -f $index, $itemCount, $_.Exception.Message)
            continue
        }

        if ($index -eq 1 -or $index -eq $itemCount -or $index % 10 -eq 0) {
            $percentComplete = 75 + [int](($index / [Math]::Max(1, $itemCount)) * 24)
            Write-Progress -Id 1 -Activity $ProgressActivity -Status "Processed $index of $itemCount Secure Score records." -PercentComplete $percentComplete
        }
        Write-Output $record
    }

    $elapsedSeconds = $stopwatch.Elapsed.TotalSeconds
    Write-Information -MessageData "Secure Score report completed. Processed: $processedCount; skipped: $skippedCount; failed: $failedCount. Elapsed: $($elapsedSeconds.ToString('N1')) seconds." -InformationAction Continue
    $global:LASTEXITCODE = 0
}
catch {
    $failedCount++
    $elapsedSeconds = $stopwatch.Elapsed.TotalSeconds
    Write-Information -MessageData "Secure Score report failed. Processed: $processedCount; skipped: $skippedCount; failed: $failedCount. Elapsed: $($elapsedSeconds.ToString('N1')) seconds. $($_.Exception.Message)" -InformationAction Continue
    $global:LASTEXITCODE = 1
    throw
}
#endregion Main logic

#region Cleanup
finally {
    $stopwatch.Stop()
    Write-Progress -Id 1 -Activity $ProgressActivity -Completed
    if ($transcriptStarted) {
        try {
            Stop-Transcript -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Warning "Could not stop the transcript at '$TranscriptPath': $($_.Exception.Message)"
        }
    }
}
#endregion Cleanup
