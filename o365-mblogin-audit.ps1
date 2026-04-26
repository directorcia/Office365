<#
.SYNOPSIS
Retrieves mailbox login audit events from Microsoft 365 unified audit logs.

.DESCRIPTION
Queries unified audit logs for mailbox activity and owner sign-in visibility in Exchange Online,
converts UTC timestamps to local time, and returns strongly shaped objects for review or export.

The script performs prerequisite validation, checks access to Search-UnifiedAuditLog, applies
strict mode, and uses robust error handling with terminating errors for unrecoverable scenarios.

.PARAMETER Hours
Number of hours to look back from EndDate when StartDate is not explicitly provided.

.PARAMETER StartDate
Optional start date for audit search. If omitted, StartDate is computed as EndDate minus Hours.

.PARAMETER EndDate
Optional end date for audit search. Defaults to the current date/time.

.PARAMETER RecordType
Record type to query in unified audit logs. Default is ExchangeItem.

.PARAMETER Operations
Audit operation values to query. Defaults to MailboxLogin and MailItemsAccessed so non-owner and owner mailbox activity can both be shown.

.PARAMETER IncludeOwnerSignInEvents
When enabled (default), also queries UserLoggedIn events and includes Exchange/Outlook-related owner sign-ins.

.PARAMETER BatchResultSize
Number of rows requested per Search-UnifiedAuditLog batch query. Default is 5000.

.PARAMETER MaxBatchesPerQuery
Maximum number of batches to retrieve per query path to prevent runaway loops.

.PARAMETER OwnerSignInMaxBatches
Maximum number of owner sign-in (UserLoggedIn) batches. Defaults lower than mailbox activity to avoid long tenant-wide scans.

.PARAMETER OwnerSignInUserId
Optional UPN to filter owner sign-in events server-side. Strongly recommended for faster owner sign-in queries.

.PARAMETER LightweightTriageMode
Runs a companion fast triage workflow that requires exactly one user and one operation.

.PARAMETER TriageUserId
User UPN for lightweight triage mode. Required when LightweightTriageMode is enabled.

.PARAMETER TriageOperation
Single operation for lightweight triage mode. Required when LightweightTriageMode is enabled.

.PARAMETER ShowGridView
Displays results in Out-GridView when available.

.PARAMETER ExportCsvPath
Optional CSV output path. Uses PSScriptRoot if a relative path is provided.

.PARAMETER PassThru
Returns result objects to pipeline.
If no explicit output option is selected, the script returns objects by default.

.PARAMETER DisableAutoConnectExchangeOnline
Disables automatic Exchange Online connection attempts when Search-UnifiedAuditLog is unavailable.

.PARAMETER UserPrincipalName
Optional user principal name used for automatic Exchange Online connection.

.PARAMETER UseDeviceAuth
Uses device code authentication for automatic Exchange Online connection.

.EXAMPLE
.\o365-mblogin-audit.ps1 -Hours 24 -ShowGridView -Verbose

Retrieves mailbox login and mailbox access activity from the last 24 hours and opens a grid view.

.EXAMPLE
.\o365-mblogin-audit.ps1 -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) -ExportCsvPath .\mailbox-logins.csv -PassThru

Retrieves seven days of events, exports to CSV, and writes objects to pipeline.

.EXAMPLE
.\o365-mblogin-audit.ps1 -UseDeviceAuth -Hours 12 -ShowGridView

Automatically connects to Exchange Online with device authentication, then retrieves 12 hours of mailbox login data.

.EXAMPLE
.\o365-mblogin-audit.ps1 -Hours 24 -Operations MailItemsAccessed -OwnerSignInUserId admin@contoso.com -OwnerSignInMaxBatches 20

Retrieves owner-focused mailbox activity plus targeted owner sign-ins for a specific user.

.EXAMPLE
.\o365-mblogin-audit.ps1 -DisableAutoConnectExchangeOnline

Runs with no automatic connection attempt and requires an existing Exchange Online session.

.EXAMPLE
.\o365-mblogin-audit.ps1 -LightweightTriageMode -TriageUserId admin@contoso.com -TriageOperation UserLoggedIn

Runs a fast single-user/single-operation incident triage query.

.NOTES
Author: Refactored by GitHub Copilot
Version: 2.2.0
Date: 2026-04-26
Prerequisite: Connected to Exchange Online with permissions to run Search-UnifiedAuditLog.
#>

[CmdletBinding()]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 720)]
    [int]$Hours = 48,

    [Parameter(Mandatory = $false)]
    [datetime]$StartDate,

    [Parameter(Mandatory = $false)]
    [datetime]$EndDate = (Get-Date),

    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[A-Za-z0-9]+$')]
    [string]$RecordType = 'ExchangeItem',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet('MailboxLogin', 'MailItemsAccessed')]
    [string[]]$Operations = @('MailboxLogin', 'MailItemsAccessed'),

    [Parameter(Mandatory = $false)]
    [bool]$IncludeOwnerSignInEvents = $true,

    [Parameter(Mandatory = $false)]
    [ValidateRange(100, 5000)]
    [int]$BatchResultSize = 5000,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 500)]
    [int]$MaxBatchesPerQuery = 100,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 200)]
    [int]$OwnerSignInMaxBatches = 10,

    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$OwnerSignInUserId,

    [Parameter(Mandatory = $false)]
    [switch]$LightweightTriageMode,

    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$TriageUserId,

    [Parameter(Mandatory = $false)]
    [ValidateSet('MailboxLogin', 'MailItemsAccessed', 'UserLoggedIn')]
    [string]$TriageOperation,

    [Parameter(Mandatory = $false)]
    [switch]$ShowGridView,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ExportCsvPath,

    [Parameter(Mandatory = $false)]
    [switch]$PassThru,

    [Parameter(Mandatory = $false)]
    [switch]$DisableAutoConnectExchangeOnline,

    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [switch]$UseDeviceAuth
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Stage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [string]$Color = 'Cyan'
    )

    Write-Host -ForegroundColor $Color ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message)
}

function New-TerminatingErrorRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [string]$ErrorId,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorCategory]$Category,

        [Parameter(Mandatory = $false)]
        [System.Exception]$InnerException
    )

    $exception = if ($null -ne $InnerException) {
        [System.InvalidOperationException]::new($Message, $InnerException)
    }
    else {
        [System.InvalidOperationException]::new($Message)
    }

    return [System.Management.Automation.ErrorRecord]::new($exception, $ErrorId, $Category, $null)
}

function Resolve-OutputPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return [System.IO.Path]::Combine($PSScriptRoot, $Path)
}

function Get-SafeAuditPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName,

        [Parameter(Mandatory = $false)]
        [string]$DefaultValue = ''
    )

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property -or $null -eq $property.Value) {
        return $DefaultValue
    }

    return [string]$property.Value
}

function Test-UnifiedAuditLogAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$ProbeStartDate,

        [Parameter(Mandatory = $true)]
        [datetime]$ProbeEndDate
    )

    Write-Stage -Message 'Validating Search-UnifiedAuditLog availability and permissions...' -Color 'DarkCyan'

    try {
        $null = Get-Command -Name Search-UnifiedAuditLog -ErrorAction Stop
    }
    catch {
        $message = 'Search-UnifiedAuditLog is unavailable. Connect to Exchange Online first (Connect-ExchangeOnline), or run .\o365-connect-exo.ps1 before this script.'
        $errorRecord = New-TerminatingErrorRecord -Message $message -ErrorId 'CmdletNotFound' -Category ([System.Management.Automation.ErrorCategory]::ObjectNotFound) -InnerException $_.Exception
        $PSCmdlet.ThrowTerminatingError($errorRecord)
    }

    try {
        $null = Search-UnifiedAuditLog -StartDate $ProbeStartDate -EndDate $ProbeEndDate -RecordType ExchangeItem -Operations MailboxLogin -ResultSize 1 -ErrorAction Stop
        Write-Stage -Message 'Unified audit query validation passed.' -Color 'Green'
    }
    catch {
        $errorRecord = New-TerminatingErrorRecord -Message 'Unified audit log query failed. Verify Exchange Online connection and RBAC permissions for Search-UnifiedAuditLog.' -ErrorId 'UnifiedAuditAccessDenied' -Category ([System.Management.Automation.ErrorCategory]::SecurityError) -InnerException $_.Exception
        $PSCmdlet.ThrowTerminatingError($errorRecord)
    }
}

function Connect-ExchangeOnlineIfNeeded {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Upn,

        [Parameter(Mandatory = $false)]
        [switch]$DeviceAuth
    )

    $existingCommand = Get-Command -Name Search-UnifiedAuditLog -ErrorAction SilentlyContinue
    if ($null -ne $existingCommand) {
        Write-Verbose -Message 'Search-UnifiedAuditLog is already available in this session.'
        Write-Host -ForegroundColor Green 'Exchange Online audit cmdlets are already available. Using existing Exchange connection.'
        return
    }

    Write-Host -ForegroundColor Yellow 'No active Exchange Online audit cmdlet/session detected.'
    Write-Host -ForegroundColor Cyan 'Creating a new Exchange Online connection for this script run...'
    Write-Stage -Message 'Checking ExchangeOnlineManagement module availability...' -Color 'DarkCyan'

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        $errorRecord = New-TerminatingErrorRecord -Message 'ExchangeOnlineManagement module is not installed. Install it first with: Install-Module ExchangeOnlineManagement -Scope CurrentUser' -ErrorId 'ExchangeOnlineModuleMissing' -Category ([System.Management.Automation.ErrorCategory]::ObjectNotFound)
        $PSCmdlet.ThrowTerminatingError($errorRecord)
    }

    Import-Module -Name ExchangeOnlineManagement -ErrorAction Stop
    Write-Stage -Message 'ExchangeOnlineManagement module imported.' -Color 'Green'

    try {
        $connectSplat = @{
            ShowBanner   = $false
            ShowProgress = $false
            ErrorAction  = 'Stop'
        }

        if ($Upn) {
            $connectSplat.UserPrincipalName = $Upn
        }

        if ($DeviceAuth) {
            $connectSplat.Device = $true
        }

        Connect-ExchangeOnline @connectSplat | Out-Null
        Write-Host -ForegroundColor Green 'Exchange Online connection established successfully.'
    }
    catch {
        $errorRecord = New-TerminatingErrorRecord -Message 'Automatic Exchange Online connection failed. Run .\o365-connect-exo.ps1 or Connect-ExchangeOnline manually, then rerun this script.' -ErrorId 'ExchangeOnlineConnectFailed' -Category ([System.Management.Automation.ErrorCategory]::OpenError) -InnerException $_.Exception
        $PSCmdlet.ThrowTerminatingError($errorRecord)
    }
}

function Get-MailboxLoginAuditRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$QueryStartDate,

        [Parameter(Mandatory = $true)]
        [datetime]$QueryEndDate,

        [Parameter(Mandatory = $true)]
        [string]$AuditRecordType,

        [Parameter(Mandatory = $true)]
        [string[]]$AuditOperation,

        [Parameter(Mandatory = $true)]
        [System.TimeZoneInfo]$LocalTimeZone,

        [Parameter(Mandatory = $true)]
        [ValidateRange(100, 5000)]
        [int]$ResultSize,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 500)]
        [int]$MaxBatches
    )

    $sessionId = [System.Guid]::NewGuid().ToString('N')
    $records = [System.Collections.Generic.List[psobject]]::new()
    $auditBatch = @([pscustomobject]@{ Placeholder = $true })
    $batchNumber = 0

    while ($auditBatch.Count -gt 0) {
        if ($batchNumber -ge $MaxBatches) {
            Write-Warning ("Stopping mailbox activity query after {0} batches (MaxBatches limit reached). Narrow the time range or increase -MaxBatchesPerQuery if needed." -f $MaxBatches)
            break
        }

        $batchNumber++
        $sessionCommand = if ($batchNumber -eq 1) { 'ReturnLargeSet' } else { 'ReturnNextPreviewPage' }
        Write-Stage -Message ("Querying mailbox activity batch {0} (operations: {1}; result size: {2})..." -f $batchNumber, ($AuditOperation -join ','), $ResultSize) -Color 'DarkCyan'
        $auditErrors = $null

        $querySplat = @{
            StartDate      = $QueryStartDate
            EndDate        = $QueryEndDate
            RecordType     = $AuditRecordType
            Operations     = $AuditOperation
            SessionId      = $sessionId
            SessionCommand = $sessionCommand
            ResultSize     = $ResultSize
            ErrorAction    = 'Stop'
            ErrorVariable  = '+auditErrors'
        }

        $auditBatch = @(Search-UnifiedAuditLog @querySplat)

        if ($auditErrors) {
            Write-Verbose -Message ("Search returned non-terminating errors: {0}" -f ($auditErrors.Count))
        }

        if ($auditBatch.Count -eq 0) {
            Write-Stage -Message ("Mailbox activity batch {0} returned no additional rows." -f $batchNumber) -Color 'DarkGray'
            break
        }

        Write-Stage -Message ("Mailbox activity batch {0} returned {1} row(s)." -f $batchNumber, $auditBatch.Count) -Color 'Green'

        # Expand JSON payload once per batch to reduce repeated object projection overhead.
        $convertedOutput = $auditBatch | Select-Object -ExpandProperty AuditData | ConvertFrom-Json

        foreach ($entry in $convertedOutput) {
            $accessType = 'Unknown'
            $rawLogonType = Get-SafeAuditPropertyValue -InputObject $entry -PropertyName 'LogonType'
            $operationValue = Get-SafeAuditPropertyValue -InputObject $entry -PropertyName 'Operation'
            $entryUserId = Get-SafeAuditPropertyValue -InputObject $entry -PropertyName 'UserId'

            if ($operationValue -eq 'MailboxLogin') {
                $accessType = 'NonOwner'
            }
            else {
                switch -Regex ($rawLogonType) {
                    '^(0|Owner)$' { $accessType = 'Owner'; break }
                    '^(1|Admin)$' { $accessType = 'NonOwner-Admin'; break }
                    '^(2|Delegate)$' { $accessType = 'NonOwner-Delegate'; break }
                    default { $accessType = 'Unknown' }
                }
            }

            $targetMailbox = Get-SafeAuditPropertyValue -InputObject $entry -PropertyName 'MailboxOwnerUPN'
            if ([string]::IsNullOrWhiteSpace($targetMailbox)) {
                $targetMailbox = $entryUserId
            }

            $records.Add([pscustomobject]@{
                    CreationTime     = $entry.CreationTime
                    LocalTime        = [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]$entry.CreationTime, $LocalTimeZone)
                    UserId           = $entryUserId
                    TargetMailbox    = $targetMailbox
                    Operation        = $operationValue
                    AccessType       = $accessType
                    LogonType        = $rawLogonType
                    ClientIpAddress  = Get-SafeAuditPropertyValue -InputObject $entry -PropertyName 'ClientIpAddress'
                })
        }
    }

    Write-Stage -Message ("Mailbox activity query complete. Collected {0} row(s)." -f $records.Count) -Color 'Green'

    return $records
}

function Get-OwnerMailboxSignInRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$QueryStartDate,

        [Parameter(Mandatory = $true)]
        [datetime]$QueryEndDate,

        [Parameter(Mandatory = $true)]
        [System.TimeZoneInfo]$LocalTimeZone,

        [Parameter(Mandatory = $true)]
        [ValidateRange(100, 5000)]
        [int]$ResultSize,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 500)]
        [int]$MaxBatches,

        [Parameter(Mandatory = $false)]
        [string]$TargetUserId
    )

    $exchangeAppId = '00000002-0000-0ff1-ce00-000000000000'
    $sessionId = [System.Guid]::NewGuid().ToString('N')
    $records = [System.Collections.Generic.List[psobject]]::new()
    $auditBatch = @([pscustomobject]@{ Placeholder = $true })
    $batchNumber = 0

    while ($auditBatch.Count -gt 0) {
        if ($batchNumber -ge $MaxBatches) {
            Write-Warning ("Stopping owner sign-in query after {0} batches (MaxBatches limit reached). Narrow the time range or increase -MaxBatchesPerQuery if needed." -f $MaxBatches)
            break
        }

        $batchNumber++
        $sessionCommand = if ($batchNumber -eq 1) { 'ReturnLargeSet' } else { 'ReturnNextPreviewPage' }
        $targetUserMessage = if ([string]::IsNullOrWhiteSpace($TargetUserId)) { 'all users' } else { $TargetUserId }
        Write-Stage -Message ("Querying owner sign-in batch {0} (operation: UserLoggedIn; user: {1}; result size: {2})..." -f $batchNumber, $targetUserMessage, $ResultSize) -Color 'DarkCyan'

        $querySplat = @{
            StartDate      = $QueryStartDate
            EndDate        = $QueryEndDate
            Operations     = 'UserLoggedIn'
            SessionId      = $sessionId
            SessionCommand = $sessionCommand
            ResultSize     = $ResultSize
            ErrorAction    = 'Stop'
        }

        if (-not [string]::IsNullOrWhiteSpace($TargetUserId)) {
            $querySplat.UserIds = $TargetUserId
        }

        $auditBatch = @(Search-UnifiedAuditLog @querySplat)

        if ($auditBatch.Count -eq 0) {
            Write-Stage -Message ("Owner sign-in batch {0} returned no additional rows." -f $batchNumber) -Color 'DarkGray'
            break
        }

        Write-Stage -Message ("Owner sign-in batch {0} returned {1} row(s)." -f $batchNumber, $auditBatch.Count) -Color 'Green'

        $convertedOutput = $auditBatch | Select-Object -ExpandProperty AuditData | ConvertFrom-Json
        foreach ($entry in $convertedOutput) {
            $appId = Get-SafeAuditPropertyValue -InputObject $entry -PropertyName 'AppId'
            $applicationDisplayName = Get-SafeAuditPropertyValue -InputObject $entry -PropertyName 'ApplicationDisplayName'
            $resourceDisplayName = Get-SafeAuditPropertyValue -InputObject $entry -PropertyName 'ResourceDisplayName'
            $workload = Get-SafeAuditPropertyValue -InputObject $entry -PropertyName 'Workload'

            $isExchangeOrOutlookSignIn = (
                $appId -eq $exchangeAppId -or
                $applicationDisplayName -match 'Exchange|Outlook' -or
                $resourceDisplayName -match 'Exchange|Outlook' -or
                $workload -match 'Exchange'
            )

            if (-not $isExchangeOrOutlookSignIn) {
                continue
            }

            $signInUser = Get-SafeAuditPropertyValue -InputObject $entry -PropertyName 'UserId'
            if ([string]::IsNullOrWhiteSpace($signInUser)) {
                $signInUser = Get-SafeAuditPropertyValue -InputObject $entry -PropertyName 'UserKey'
            }

            $records.Add([pscustomobject]@{
                    CreationTime     = $entry.CreationTime
                    LocalTime        = [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]$entry.CreationTime, $LocalTimeZone)
                    UserId           = $signInUser
                    TargetMailbox    = $signInUser
                    Operation        = 'UserLoggedIn'
                    AccessType       = 'OwnerSignIn'
                    LogonType        = $null
                    ClientIpAddress  = Get-SafeAuditPropertyValue -InputObject $entry -PropertyName 'ClientIP'
                })
        }
    }

    Write-Stage -Message ("Owner sign-in query complete. Collected {0} row(s)." -f $records.Count) -Color 'Green'

    return $records
}

function Get-ExchangeItemOperationSample {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$QueryStartDate,

        [Parameter(Mandatory = $true)]
        [datetime]$QueryEndDate,

        [Parameter(Mandatory = $false)]
        [int]$MaxRows = 500
    )

    $sampleRows = @(Search-UnifiedAuditLog -StartDate $QueryStartDate -EndDate $QueryEndDate -RecordType ExchangeItem -ResultSize $MaxRows -ErrorAction Stop)
    if ($sampleRows.Count -eq 0) {
        return @()
    }

    $converted = $sampleRows | Select-Object -ExpandProperty AuditData | ConvertFrom-Json
    return @($converted | Group-Object -Property Operation | Sort-Object -Property Count -Descending | Select-Object -First 10 -Property Name, Count)
}

function Get-LightweightTriageAuditRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$QueryStartDate,

        [Parameter(Mandatory = $true)]
        [datetime]$QueryEndDate,

        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [ValidateSet('MailboxLogin', 'MailItemsAccessed', 'UserLoggedIn')]
        [string]$Operation,

        [Parameter(Mandatory = $true)]
        [System.TimeZoneInfo]$LocalTimeZone,

        [Parameter(Mandatory = $true)]
        [ValidateRange(100, 5000)]
        [int]$ResultSize,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 200)]
        [int]$MaxBatches
    )

    Write-Stage -Message 'Starting lightweight triage query path...' -Color 'DarkCyan'

    if ($Operation -eq 'UserLoggedIn') {
        return @(Get-OwnerMailboxSignInRecord -QueryStartDate $QueryStartDate -QueryEndDate $QueryEndDate -LocalTimeZone $LocalTimeZone -ResultSize $ResultSize -MaxBatches $MaxBatches -TargetUserId $UserId)
    }

    $records = @(Get-MailboxLoginAuditRecord -QueryStartDate $QueryStartDate -QueryEndDate $QueryEndDate -AuditRecordType 'ExchangeItem' -AuditOperation @($Operation) -LocalTimeZone $LocalTimeZone -ResultSize $ResultSize -MaxBatches $MaxBatches)
    return @($records | Where-Object { $_.UserId -eq $UserId -or $_.TargetMailbox -eq $UserId })
}

try {
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        $errorRecord = New-TerminatingErrorRecord -Message 'PowerShell 5.1 or later is required.' -ErrorId 'UnsupportedPowerShellVersion' -Category ([System.Management.Automation.ErrorCategory]::NotImplemented)
        $PSCmdlet.ThrowTerminatingError($errorRecord)
    }

    if (-not $PSBoundParameters.ContainsKey('StartDate')) {
        $StartDate = $EndDate.AddHours(-$Hours)
    }

    if ($StartDate -ge $EndDate) {
        $errorRecord = New-TerminatingErrorRecord -Message 'StartDate must be earlier than EndDate.' -ErrorId 'InvalidDateRange' -Category ([System.Management.Automation.ErrorCategory]::InvalidArgument)
        $PSCmdlet.ThrowTerminatingError($errorRecord)
    }

    if ($LightweightTriageMode) {
        if ([string]::IsNullOrWhiteSpace($TriageUserId)) {
            $errorRecord = New-TerminatingErrorRecord -Message 'TriageUserId is required when LightweightTriageMode is enabled.' -ErrorId 'MissingTriageUserId' -Category ([System.Management.Automation.ErrorCategory]::InvalidArgument)
            $PSCmdlet.ThrowTerminatingError($errorRecord)
        }

        if ([string]::IsNullOrWhiteSpace($TriageOperation)) {
            $errorRecord = New-TerminatingErrorRecord -Message 'TriageOperation is required when LightweightTriageMode is enabled.' -ErrorId 'MissingTriageOperation' -Category ([System.Management.Automation.ErrorCategory]::InvalidArgument)
            $PSCmdlet.ThrowTerminatingError($errorRecord)
        }

        if (-not $PSBoundParameters.ContainsKey('BatchResultSize')) {
            $BatchResultSize = 1000
        }

        if (-not $PSBoundParameters.ContainsKey('MaxBatchesPerQuery')) {
            $MaxBatchesPerQuery = 5
        }

        if (-not $PSBoundParameters.ContainsKey('IncludeOwnerSignInEvents')) {
            $IncludeOwnerSignInEvents = $false
        }
    }

    # Log only non-sensitive bound parameters for operational diagnostics.
    $safeParameters = @{}
    foreach ($key in $PSBoundParameters.Keys) {
        if ($key -notin @('Credential', 'Password', 'Token', 'ClientSecret')) {
            $safeParameters[$key] = $PSBoundParameters[$key]
        }
    }

    if ($safeParameters.Count -gt 0) {
        $parameterSummary = $safeParameters.GetEnumerator() | Sort-Object -Property Name | ForEach-Object { "{0}={1}" -f $_.Name, $_.Value }
        Write-Verbose -Message ("Running with parameters: {0}" -f ($parameterSummary -join '; '))
    }

    Write-Host -ForegroundColor Cyan 'Mailbox login audit script started.'
    Write-Stage -Message ("Time window: {0} to {1}" -f $StartDate, $EndDate) -Color 'DarkCyan'
    Write-Stage -Message ("Requested operations: {0}" -f ($Operations -join ', ')) -Color 'DarkCyan'
    Write-Stage -Message ("Lightweight triage mode: {0}" -f $LightweightTriageMode) -Color 'DarkCyan'
    if ($LightweightTriageMode) {
        Write-Stage -Message ("Triage user: {0}" -f $TriageUserId) -Color 'DarkCyan'
        Write-Stage -Message ("Triage operation: {0}" -f $TriageOperation) -Color 'DarkCyan'
    }

    if ($Operations.Count -eq 1 -and $Operations[0] -eq 'MailboxLogin') {
        Write-Warning 'Operation MailboxLogin is primarily non-owner mailbox access. Include MailItemsAccessed to improve owner activity visibility.'
    }

    Write-Stage -Message ("Include owner sign-in events: {0}" -f $IncludeOwnerSignInEvents) -Color 'DarkCyan'
    Write-Stage -Message ("Owner sign-in user filter: {0}" -f $(if ([string]::IsNullOrWhiteSpace($OwnerSignInUserId)) { 'none (tenant-wide)' } else { $OwnerSignInUserId })) -Color 'DarkCyan'
    Write-Stage -Message ("Batch result size: {0}" -f $BatchResultSize) -Color 'DarkCyan'
    Write-Stage -Message ("Max batches per query: {0}" -f $MaxBatchesPerQuery) -Color 'DarkCyan'
    Write-Stage -Message ("Max owner sign-in batches: {0}" -f $OwnerSignInMaxBatches) -Color 'DarkCyan'

    Write-Stage -Message 'Resolving local timezone...' -Color 'DarkCyan'
    $localTimeZone = Get-TimeZone -ErrorAction Stop
    Write-Stage -Message ("Local timezone resolved: {0}" -f $localTimeZone.Id) -Color 'Green'

    if (-not $DisableAutoConnectExchangeOnline) {
        Write-Stage -Message 'Ensuring Exchange Online connectivity...' -Color 'DarkCyan'
        Connect-ExchangeOnlineIfNeeded -Upn $UserPrincipalName -DeviceAuth:$UseDeviceAuth
    }

    Test-UnifiedAuditLogAccess -ProbeStartDate $StartDate -ProbeEndDate $EndDate

    if ($LightweightTriageMode) {
        $results = @(Get-LightweightTriageAuditRecord -QueryStartDate $StartDate -QueryEndDate $EndDate -UserId $TriageUserId -Operation $TriageOperation -LocalTimeZone $localTimeZone -ResultSize $BatchResultSize -MaxBatches $MaxBatchesPerQuery)
    }
    else {
        Write-Stage -Message 'Starting mailbox activity query...' -Color 'DarkCyan'
        $mailboxAccessResults = @(Get-MailboxLoginAuditRecord -QueryStartDate $StartDate -QueryEndDate $EndDate -AuditRecordType $RecordType -AuditOperation $Operations -LocalTimeZone $localTimeZone -ResultSize $BatchResultSize -MaxBatches $MaxBatchesPerQuery)
        $ownerSignInResults = @()
        if ($IncludeOwnerSignInEvents) {
            if ([string]::IsNullOrWhiteSpace($OwnerSignInUserId)) {
                Write-Warning 'Owner sign-in query is running tenant-wide because -OwnerSignInUserId was not provided. This can be slow in larger tenants.'
            }
            Write-Stage -Message 'Starting owner sign-in query...' -Color 'DarkCyan'
            $ownerSignInResults = @(Get-OwnerMailboxSignInRecord -QueryStartDate $StartDate -QueryEndDate $EndDate -LocalTimeZone $localTimeZone -ResultSize $BatchResultSize -MaxBatches $OwnerSignInMaxBatches -TargetUserId $OwnerSignInUserId)
        }

        $results = @($mailboxAccessResults + $ownerSignInResults | Sort-Object -Property CreationTime, UserId, Operation, ClientIpAddress -Unique)
    }

    Write-Host -ForegroundColor Green ("Retrieved {0} mailbox login audit record(s)." -f $results.Count)

    if ($results.Count -gt 0) {
        $summary = $results | Group-Object -Property AccessType | Sort-Object -Property Name | ForEach-Object { "{0}={1}" -f $_.Name, $_.Count }
        Write-Host -ForegroundColor Green ("Access summary: {0}" -f ($summary -join '; '))
    }

    if ($results.Count -eq 0) {
        Write-Warning 'No mailbox access records were found for the selected time window for the requested operations.'
        Write-Host -ForegroundColor Yellow 'Collecting quick diagnostic sample of Exchange mailbox operations for the same time window...'

        $operationSample = @(Get-ExchangeItemOperationSample -QueryStartDate $StartDate -QueryEndDate $EndDate)
        if ($operationSample.Count -gt 0) {
            Write-Host -ForegroundColor Yellow 'Top ExchangeItem operations found (sample):'
            $operationSample | Format-Table -AutoSize | Out-Host
            Write-Host -ForegroundColor Yellow 'If you want broader mailbox activity, rerun with -Operations MailboxLogin,MailItemsAccessed and/or increase -Hours.'
        }
        else {
            Write-Warning 'No ExchangeItem operations were returned in the sample query. Try increasing -Hours (for example, 168) or validate unified audit log ingestion in your tenant.'
        }
    }

    if ($ExportCsvPath) {
        Write-Stage -Message 'Exporting results to CSV...' -Color 'DarkCyan'
        $resolvedExportPath = Resolve-OutputPath -Path $ExportCsvPath
        $results | Export-Csv -Path $resolvedExportPath -NoTypeInformation -Encoding UTF8
        Write-Host -ForegroundColor Green ("Exported {0} records to {1}" -f $results.Count, $resolvedExportPath)
    }

    if ($ShowGridView) {
        Write-Stage -Message 'Opening Out-GridView for interactive review...' -Color 'DarkCyan'
        $outGridViewCommand = Get-Command -Name Out-GridView -ErrorAction SilentlyContinue
        if ($null -ne $outGridViewCommand) {
            $results | Out-GridView -Title 'Mailbox Login Audit Results'
        }
        else {
            Write-Warning 'Out-GridView is not available on this host. Install Microsoft.PowerShell.GraphicalTools or run in Windows PowerShell.'
        }
    }

    $emitToPipeline = $PassThru -or (-not $ShowGridView -and -not $ExportCsvPath)
    if ($emitToPipeline) {
        $results
    }
}
catch {
    $errorRecord = New-TerminatingErrorRecord -Message $_.Exception.Message -ErrorId 'MailboxLoginAuditScriptFailed' -Category ([System.Management.Automation.ErrorCategory]::NotSpecified) -InnerException $_.Exception
    $PSCmdlet.ThrowTerminatingError($errorRecord)
}
finally {
    Write-Host -ForegroundColor Cyan 'Mailbox login audit script completed.'
}

<#
SUMMARY
- Issues fixed:
  - Removed legacy WMI time zone lookup and brittle loop initialization patterns.
  - Removed inefficient array += usage for result accumulation.
  - Corrected case inconsistencies and legacy style usage.

- Improvements made:
  - Added full comment-based help with parameters, examples, and notes.
  - Added CmdletBinding, strict mode, validated parameters, and safer path handling.
  - Added explicit prerequisite and RBAC/access probe before full query.
    - Added structured helper functions and terminating error records for critical failures.
    - Added predictable paging controls, user-targeted owner sign-in filtering, and de-duplication.
    - Added strict-mode-safe property extraction helpers for variable audit payload schemas.

- Performance gains expected:
  - Lower memory churn by using a generic list instead of repeated array concatenation.
  - Reduced projection overhead with batch JSON expansion and direct PSCustomObject creation.
    - Faster owner-sign-in retrieval when UserIds filter is used.
    - Reduced duplicate output rows from paged query overlap.

- Security enhancements applied:
  - Added input validation and date-range validation.
  - Added safe parameter logging that excludes secret-like keys.
  - Eliminated insecure execution policy guidance and unsupported dynamic execution patterns.
    - Enforced terminating error propagation for deterministic failure behavior.
#>
