<#
.SYNOPSIS
    Collects read-only Microsoft 365 AI readiness signals into one JSON report.

.DESCRIPTION
    Connects to Microsoft Graph and any available Microsoft 365 services,
    collects tenant configuration and posture signals, records collector
    failures as partial or unavailable data, and writes one ordered JSON report.
    The default sign-in is delegated device code. Optional certificate
    parameters acquire an app-only Graph token, or the script reuses an existing
    Connect-MgGraph session. Tenant configuration is never changed. The report
    file is written only after ShouldProcess approval. -NonInteractive and
    -Headless skip the browser, clipboard, prompts, module install, and device code.

.PARAMETER TenantId
    Tenant ID or domain to use for authentication. When omitted, the current
    account's organization context or the identity provider's organizations
    endpoint is used.

.PARAMETER OutputPath
    Destination for the consolidated JSON report. Relative paths are resolved
    from the current location. Parent directories are created when approved.

.PARAMETER GraphClientId
    Entra application client ID used for Microsoft Graph device-code login.

.PARAMETER PnPClientId
    Optional Entra application client ID for PnP.PowerShell authentication.
    Environment variables ENTRAID_APP_ID, ENTRAID_CLIENT_ID, and PNP_CLIENT_ID
    are checked when this parameter is omitted.

.PARAMETER GraphScopes
    Optional delegated Graph scopes to request during authentication.

.PARAMETER RequestRequiredGraphScopesUpFront
    Requests the exporter's full delegated Graph scope set before collection.

.PARAMETER ForceInteractiveTeamsPnPLogin
    Forces a new Teams sign-in instead of reusing an existing Teams session.

.PARAMETER SharePointCollectionChild
    Internal mode used by the script's isolated SharePoint collection process.

.PARAMETER SharePointAdminUrl
    SharePoint admin center URL required by the internal child collection mode.

.PARAMETER IncludeSampling
    Enables sampling for supported large SharePoint collections.

.PARAMETER IncludeDiagnostics
    Includes collector timing details in the JSON report.

.PARAMETER IncludeDetailedData
    Includes detailed Secure Score control-profile data and a longer score history.

.PARAMETER FailOnCollectorError
    Opt-in for schedulers. When at least one section was processed and
    FailedCollectors is greater than zero, the process exits 3. Without this
    switch that run still exits 0, which is the result existing schedules
    already receive. Exit 2 is unchanged when no section was processed.

.PARAMETER SampleSize
    Maximum number of SharePoint sites to include when sampling is enabled.
    Valid values are 1 through 5000; the default is 200.

.PARAMETER NonInteractive
    Headless switch. Disables browser launch, clipboard copy, prompts, and
    device-code or interactive sign-in fallback.

.PARAMETER Headless
    Same behavior as -NonInteractive.

.PARAMETER GraphCertificateThumbprint
    Thumbprint of the app-only certificate in CurrentUser\My or LocalMachine\My.

.PARAMETER GraphCertificatePath
    PFX path for the app-only Graph client-credentials assertion. The password
    is read from the environment variable named by
    -GraphCertificatePasswordEnvironmentVariable.

.PARAMETER GraphCertificatePasswordEnvironmentVariable
    Name of the environment variable that holds the PFX password. This is not
    the password itself.

.PARAMETER ExchangeAppId
    Application id for Connect-ExchangeOnline certificate authentication.
    In headless mode, an omitted value falls back to -GraphClientId.

.PARAMETER ExchangeOrganization
    Organization passed to Connect-ExchangeOnline. Defaults to -TenantId.

.PARAMETER ExchangeCertificateThumbprint
    Certificate thumbprint for Connect-ExchangeOnline.

.PARAMETER ExchangeCertificatePath
    PFX path for Connect-ExchangeOnline.

.PARAMETER ExchangeCertificatePasswordEnvironmentVariable
    Environment variable name for the Exchange PFX password.

.PARAMETER IppsAppId
    Application id for Connect-IPPSSession certificate authentication.

.PARAMETER IppsOrganization
    Organization passed to Connect-IPPSSession. Defaults to -TenantId.

.PARAMETER IppsCertificateThumbprint
    Certificate thumbprint for Connect-IPPSSession.

.PARAMETER IppsCertificatePath
    PFX path for Connect-IPPSSession.

.PARAMETER IppsCertificatePasswordEnvironmentVariable
    Environment variable name for the IPPS PFX password.

.PARAMETER TeamsAppId
    Application id for Connect-MicrosoftTeams certificate authentication.

.PARAMETER TeamsOrganization
    Tenant id passed to Connect-MicrosoftTeams. Defaults to -TenantId.

.PARAMETER TeamsCertificateThumbprint
    Certificate thumbprint for Connect-MicrosoftTeams.

.PARAMETER TeamsCertificatePath
    PFX path for Connect-MicrosoftTeams. On Linux the certificate object is
    passed to -Certificate.

.PARAMETER TeamsCertificatePasswordEnvironmentVariable
    Environment variable name for the Teams PFX password.

.PARAMETER PnPCertificateThumbprint
    Certificate thumbprint for Connect-PnPOnline.

.PARAMETER PnPCertificatePath
    PFX path for Connect-PnPOnline.

.PARAMETER PnPCertificatePasswordEnvironmentVariable
    Environment variable name for the PnP PFX password.

.EXAMPLE
    .\m365-ai-ready-export.ps1
    Collects available signals and writes a timestamped report in the current folder.

.EXAMPLE
    .\m365-ai-ready-export.ps1 -TenantId contoso.onmicrosoft.com -OutputPath 'C:\Reports\Contoso readiness.json'
    Collects for the specified tenant and writes the report at the supplied path.

.EXAMPLE
    .\m365-ai-ready-export.ps1 -IncludeSampling -SampleSize 500 -IncludeDiagnostics -WhatIf
    Collects using a 500-site sample and previews, but does not write, the report.

.EXAMPLE
    $env:GRAPH_CERT_PASSWORD = '<pfx-password>'
    .\m365-ai-ready-export.ps1 -Headless -TenantId contoso.onmicrosoft.com -GraphClientId 11111111-2222-3333-4444-555555555555 -GraphCertificatePath .\app.pfx -GraphCertificatePasswordEnvironmentVariable GRAPH_CERT_PASSWORD -PnPClientId 11111111-2222-3333-4444-555555555555 -PnPCertificatePath .\app.pfx -PnPCertificatePasswordEnvironmentVariable GRAPH_CERT_PASSWORD
    Headless app-only Graph and PnP certificate sign-in. The password stays in the environment.

.EXAMPLE
    Connect-MgGraph -TenantId contoso.onmicrosoft.com -ClientId 11111111-2222-3333-4444-555555555555 -CertificateThumbprint 0123456789ABCDEF0123456789ABCDEF01234567 -NoWelcome
    .\m365-ai-ready-export.ps1 -NonInteractive -TenantId contoso.onmicrosoft.com
    Reuses an existing Connect-MgGraph app-only session. AuthenticationMode is ExistingSession.

.INPUTS
    None. The script does not accept pipeline input.

.OUTPUTS
    System.Management.Automation.PSCustomObject
    Writes a consolidated JSON file and emits a summary object with OutputPath,
    ReportWritten, OverallScore, OverallReadiness, ProcessedSections,
    PartialSections, SkippedSections, FailedCollectors, and ExitCode.

.NOTES
    Requires PowerShell 7.0 or later. The default Microsoft Graph sign-in is
    interactive device code. -NonInteractive/-Headless with a certificate, or an
    existing Connect-MgGraph session, runs without a prompt. Optional
    MicrosoftTeams, PnP.PowerShell, and ExchangeOnlineManagement modules enable
    additional signals. Graph permissions depend on the collectors; unavailable
    permissions are recorded as data gaps.

    Exit codes:
    0  The run finished. Collector errors stay in FailedCollectors and do not
       change this code unless -FailOnCollectorError is set.
    1  A fatal error stopped the run before the normal summary.
    2  FailedCollectors is greater than zero and no section was processed.
    3  -FailOnCollectorError was set, at least one section was processed, and
       FailedCollectors is greater than zero.

    Change log:
    - 2026-11-23: PhishingResistantMfaCoverage is the percent of enabled
      members with a phishing-resistant method. PimSignal reports eligible
      privileged assignments. ScriptVersion is 1.3.
    - 2026-11-22: A failed DLP or retention read marks Data governance partial
      and names the failed cmdlet. Tenant settings that block anyone links
      earn the no-anyone-links score, with the same reduced completeness as
      other tenant-only sharing evidence.
    - 2026-11-22: DLP and retention cmdlet failures are recorded on Errors.
      A failed Get-LabelPolicy, including a failed retry, marks Data governance
      partial when labels were returned.
    - 2026-11-22: Purview label cmdlets use -ErrorAction Stop. A failure is
      not collected and is recorded on Errors. Get-LabelPolicy is retried once
      when labels exist and the first policy read is empty; a second empty
      read stays not collected instead of a scored zero.
    - 2026-11-22: A Data governance section that collected neither labels nor
      compliance data keeps the Purview connection failure in its reason.
      -FailOnCollectorError exits 3 when the run finished with collector errors.
    - 2026-11-22: Exchange Online and Purview connect before MicrosoftTeams
      is imported, so Exchange Online Management 3.10.1 can load
      Microsoft.Identity.Client 4.83.1. A failed Exchange or Purview
      connection is recorded on Errors and the connection status. DLP and
      retention policy lists stay flat. Identity scoring uses the
      enabled-member MFA percentage. Adoption divides by enabled members.
      PnP.PowerShell stays out of the parent process.
    - 2026-11-22: Parent process no longer loads PnP.PowerShell (SharePoint
      collection stays in the child). Content exposure, Office 365 active
      users, Copilot report fallbacks, Teams recording, retention/DLP,
      label policy counts, Secure Score addressed gaps, licence denominator,
      and metric-level completeness are corrected.
    - 2026-11-22: Adds write confirmation support and distinguishes failed
      SharePoint/Exchange collection from successful empty results.
    - 2026-11-22: App-only certificate authentication for Graph, Exchange,
      Purview, Teams, and PnP; existing-session reuse; AuthenticationMode
      DeviceCode, AppOnlyCertificate, or ExistingSession. Directory.Read.All
      and *.ReadWrite.All satisfy narrower reads, and collectors treat only
      HTTP 403 as missing. CSV reports are UTF-8 decoded. Copilot counts users
      active in the period. Secure Score gaps join the latest controlScores.
      Gap walks stop on dates. Labels fall back to Purview and stay null when
      not collected. SharePoint settings can come from Graph. MFA is split by
      members, guests, and admins.
    Documentation:
    https://github.com/directorcia/Office365/wiki/Microsoft-365-AI-Readiness-Export-%E2%80%90-Execution-and-Operations-Guide
#>

#requires -Version 7.0

# ============================================================================
# PARAMETERS
# ============================================================================
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Standard')]
[OutputType([PSCustomObject])]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    'PSAvoidUsingPlainTextForPassword',
    '',
    Justification='Certificate parameters that contain Password name an environment variable. They never accept the password value.')]

param(
    [Parameter(ParameterSetName = 'Standard')]
    [Parameter(ParameterSetName = 'Child')]
    [string]$TenantId,

    [Parameter(ParameterSetName = 'Standard')]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = ".\AIReadiness_$(Get-Date -Format 'yyyyMMdd_HHmmss').json",

    [Parameter(ParameterSetName = 'Standard')]
    [Parameter(ParameterSetName = 'Child')]
    [ValidateNotNullOrEmpty()]
    [string]$GraphClientId = '04b07795-8ddb-461a-bbee-02f9e1bf7b46',

    [Parameter(ParameterSetName = 'Standard')]
    [Parameter(ParameterSetName = 'Child')]
    [string]$PnPClientId,

    [Parameter(ParameterSetName = 'Standard')]
    [string[]]$GraphScopes,

    [Parameter(ParameterSetName = 'Standard')]
    [switch]$RequestRequiredGraphScopesUpFront,

    [Parameter(ParameterSetName = 'Standard')]
    [switch]$ForceInteractiveTeamsPnPLogin,

    [Parameter(ParameterSetName = 'Child', Mandatory = $true, HelpMessage = 'Internal switch used by the isolated SharePoint child process.')]
    [switch]$SharePointCollectionChild,

    [Parameter(ParameterSetName = 'Child', Mandatory = $true, HelpMessage = 'SharePoint admin URL (e.g. https://contoso-admin.sharepoint.com) for child collection mode.')]
    [ValidateNotNullOrEmpty()]
    [string]$SharePointAdminUrl,

    [Parameter(ParameterSetName = 'Standard')]
    [Parameter(ParameterSetName = 'Child')]
    [switch]$IncludeSampling,

    [Parameter(ParameterSetName = 'Standard')]
    [switch]$IncludeDiagnostics,

    [Parameter(ParameterSetName = 'Standard')]
    [switch]$IncludeDetailedData,

    [Parameter(ParameterSetName = 'Standard')]
    [switch]$FailOnCollectorError,

    [Parameter(ParameterSetName = 'Standard')]
    [Parameter(ParameterSetName = 'Child')]
    [ValidateRange(1, 5000)]
    [int]$SampleSize = 200,
    [switch]$NonInteractive,
    [switch]$Headless,
    [string]$GraphCertificateThumbprint,
    [string]$GraphCertificatePath,
    [string]$GraphCertificatePasswordEnvironmentVariable,
    [string]$ExchangeAppId,
    [string]$ExchangeOrganization,
    [string]$ExchangeCertificateThumbprint,
    [string]$ExchangeCertificatePath,
    [string]$ExchangeCertificatePasswordEnvironmentVariable,
    [string]$IppsAppId,
    [string]$IppsOrganization,
    [string]$IppsCertificateThumbprint,
    [string]$IppsCertificatePath,
    [string]$IppsCertificatePasswordEnvironmentVariable,
    [string]$TeamsAppId,
    [string]$TeamsOrganization,
    [string]$TeamsCertificateThumbprint,
    [string]$TeamsCertificatePath,
    [string]$TeamsCertificatePasswordEnvironmentVariable,
    [string]$PnPCertificateThumbprint,
    [string]$PnPCertificatePath,
    [string]$PnPCertificatePasswordEnvironmentVariable
)

# ============================================================================
# CONFIGURATION & RUNTIME STATE
# ============================================================================
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$script:ExitCode = 0
$script:ExportOptions = [ordered]@{
    TenantId                         = $TenantId
    PnPClientId                      = $PnPClientId
    GraphScopes                      = $GraphScopes
    RequestRequiredGraphScopesUpFront = [bool]$RequestRequiredGraphScopesUpFront
    ForceInteractiveTeamsPnPLogin    = [bool]$ForceInteractiveTeamsPnPLogin
    IncludeDetailedData              = [bool]$IncludeDetailedData
    NonInteractive                   = [bool]$NonInteractive
    Headless                         = [bool]$Headless
}
$script:NonInteractive = [bool]$NonInteractive
$script:Headless = [bool]$Headless
$script:GraphCertificateThumbprint = $GraphCertificateThumbprint
$script:GraphCertificatePath = $GraphCertificatePath
$script:GraphCertificatePasswordEnvironmentVariable = $GraphCertificatePasswordEnvironmentVariable
$script:ExchangeAppId = $ExchangeAppId
$script:ExchangeOrganization = $ExchangeOrganization
$script:ExchangeCertificateThumbprint = $ExchangeCertificateThumbprint
$script:ExchangeCertificatePath = $ExchangeCertificatePath
$script:ExchangeCertificatePasswordEnvironmentVariable = $ExchangeCertificatePasswordEnvironmentVariable
$script:IppsAppId = $IppsAppId
$script:IppsOrganization = $IppsOrganization
$script:IppsCertificateThumbprint = $IppsCertificateThumbprint
$script:IppsCertificatePath = $IppsCertificatePath
$script:IppsCertificatePasswordEnvironmentVariable = $IppsCertificatePasswordEnvironmentVariable
$script:TeamsAppId = $TeamsAppId
$script:TeamsOrganization = $TeamsOrganization
$script:TeamsCertificateThumbprint = $TeamsCertificateThumbprint
$script:TeamsCertificatePath = $TeamsCertificatePath
$script:TeamsCertificatePasswordEnvironmentVariable = $TeamsCertificatePasswordEnvironmentVariable
$script:PnPCertificateThumbprint = $PnPCertificateThumbprint
$script:PnPCertificatePath = $PnPCertificatePath
$script:PnPCertificatePasswordEnvironmentVariable = $PnPCertificatePasswordEnvironmentVariable

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================
function Write-Step {
    param(
        [Parameter(Mandatory)][string]$Message,
        [string]$Color = 'Cyan'
    )
    Write-Information -MessageData $Message -Tags @('AIReadiness', $Color) -InformationAction Continue
    Write-Verbose $Message
}

function Write-Phase {
    param([Parameter(Mandatory)][string]$Title)

    Write-Information -MessageData '' -Tags @('AIReadiness') -InformationAction Continue
    Write-Information -MessageData "===== $Title =====" -Tags @('AIReadiness', 'Magenta') -InformationAction Continue
    Write-Verbose "Phase: $Title"
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
    if (-not [string]::IsNullOrWhiteSpace($script:PnPClientId)) {
        return $script:PnPClientId
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
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AdminUrl
    )

    $effectivePnPClientId = Get-PnPClientIdForConnection
    $effectiveTenant = if (-not [string]::IsNullOrWhiteSpace($script:TenantId)) { $script:TenantId } else { $null }
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

    $pnpThumbprint = Get-ScriptStringParameter -Name 'PnPCertificateThumbprint'
    $pnpCertificatePath = Get-ScriptStringParameter -Name 'PnPCertificatePath'
    $pnpPasswordVariable = Get-ScriptStringParameter -Name 'PnPCertificatePasswordEnvironmentVariable'
    if ((-not [string]::IsNullOrWhiteSpace($pnpThumbprint) -or -not [string]::IsNullOrWhiteSpace($pnpCertificatePath)) -and -not [string]::IsNullOrWhiteSpace($effectivePnPClientId)) {
        $certificateParams = @{}
        foreach ($key in $baseParams.Keys) { $certificateParams[$key] = $baseParams[$key] }
        if (-not [string]::IsNullOrWhiteSpace($pnpCertificatePath)) {
            $certificateParams['CertificatePath'] = $pnpCertificatePath
            $certificatePassword = Get-EnvironmentSecureString -VariableName $pnpPasswordVariable
            if ($null -ne $certificatePassword) { $certificateParams['CertificatePassword'] = $certificatePassword }
        }
        else {
            $certificateParams['Thumbprint'] = (Get-NormalizedCertificateThumbprint -Thumbprint $pnpThumbprint)
        }

        try {
            Connect-PnPOnline @certificateParams
            return [ordered]@{
                Succeeded = $true
                Method = 'Certificate'
                PnPClientIdUsed = $effectivePnPClientId
            }
        }
        catch {
            $attemptErrors.Add("Certificate: $($_.Exception.Message)") | Out-Null
            if (Test-NonInteractiveMode) {
                return [ordered]@{
                    Succeeded = $false
                    Reason = "Headless PnP certificate authentication failed: $($attemptErrors -join ' | ')"
                    PnPClientIdUsed = $effectivePnPClientId
                }
            }
        }
    }
    elseif (Test-NonInteractiveMode) {
        return [ordered]@{
            Succeeded = $false
            Reason = 'Headless mode does not fall back to interactive or device-code PnP sign-in. Pass -PnPClientId with -PnPCertificateThumbprint or -PnPCertificatePath.'
            PnPClientIdUsed = if (-not [string]::IsNullOrWhiteSpace($effectivePnPClientId)) { $effectivePnPClientId } else { 'not provided' }
        }
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
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
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
        throw 'PnP.PowerShell for PowerShell 7+ was not found (only a Windows PowerShell module path was detected). Install it manually in pwsh with: Install-Module -Name PnP.PowerShell -Scope CurrentUser'
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
    try {
        $location = [string]$tenantAssemblies[0].Location
    }
    catch {
        # Dynamic or in-memory assemblies throw NotSupportedException on .Location.
        Write-Debug "Assembly location unavailable: $($_.Exception.Message)"
    }

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
$script:ExchangeConnected = $false
$script:ComplianceConnected = $false
$script:ExchangeConnectionStatus = 'not attempted'
$script:ComplianceConnectionStatus = 'not attempted'
$script:ExchangeConnectionError = ''
$script:ComplianceConnectionError = ''
$script:SharePointChildCollectionSucceeded = $false
$script:GraphMissingScopes = @()
$script:GraphProvidedScopes = @()
$script:GraphRestAllCache = @{}
$script:GraphUsersCache = @{}
$script:GraphReportRowsCache = @{}
$script:CopilotUsageCache = @{}
$script:GraphRestMaxPageCount = 500
$script:GraphAuthenticationMode = $null
$script:GraphAppClientId = $null
$script:GraphAppDisplayName = $null
$script:GraphAuthenticationCertificate = $null
$script:GraphLastHttpStatus = $null
$script:GraphLastErrorMessage = $null
$script:GraphLastErrorCode = $null
$script:GraphReportLastError = $null
$script:GraphReportDiagnostics = [System.Collections.Generic.List[object]]::new()

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

    $graphReportDiagnosticsVar = Get-Variable -Name GraphReportDiagnostics -Scope Script -ErrorAction SilentlyContinue
    if ($null -eq $graphReportDiagnosticsVar -or $null -eq $graphReportDiagnosticsVar.Value) {
        $script:GraphReportDiagnostics = [System.Collections.Generic.List[object]]::new()
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

function ConvertTo-SharingCapabilityName {
    # SharePoint SharingCapability is Microsoft.SharePoint.Client.SharingCapabilities
    # (the tenant property uses the same values). A collector may pass the live enum,
    # its underlying integer, a numeric string from JSON, or the enum name:
    #   0 Disabled
    #   1 ExternalUserSharingOnly
    #   2 ExternalUserAndGuestSharing  (Anyone / guest links)
    #   3 ExistingExternalUserSharingOnly
    # AnonymousAccess is not a member of this enum. Unrecognised values are returned
    # unchanged so a future name is still visible to the caller.
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return $null }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    $parsedInt = 0
    $isInteger = [int]::TryParse($text, [System.Globalization.NumberStyles]::Integer, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedInt)
    if (-not $isInteger -and ($Value -is [System.Enum])) {
        try {
            $parsedInt = [int]$Value
            $isInteger = $true
        }
        catch {
            $isInteger = $false
        }
    }

    if ($isInteger) {
        switch ($parsedInt) {
            0 { return 'Disabled' }
            1 { return 'ExternalUserSharingOnly' }
            2 { return 'ExternalUserAndGuestSharing' }
            3 { return 'ExistingExternalUserSharingOnly' }
            default { return $text }
        }
    }

    $knownNames = @(
        'Disabled',
        'ExternalUserSharingOnly',
        'ExternalUserAndGuestSharing',
        'ExistingExternalUserSharingOnly'
    )
    foreach ($name in $knownNames) {
        if ($text -eq $name) { return $name }
    }

    return $text
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

function Test-NonInteractiveMode {
    foreach ($name in @('NonInteractive', 'Headless')) {
        $variable = Get-Variable -Name $name -Scope Script -ErrorAction SilentlyContinue
        if ($null -ne $variable -and [bool]$variable.Value) {
            return $true
        }
    }

    return $false
}

function Get-ScriptStringParameter {
    param([Parameter(Mandatory)][string]$Name)

    $variable = Get-Variable -Name $Name -Scope Script -ErrorAction SilentlyContinue
    if ($null -eq $variable -or $null -eq $variable.Value) { return '' }
    return [string]$variable.Value
}

function ConvertTo-CollectedInt {
    param([AllowNull()][object]$Value)

    if (-not (Test-ValueCollected -Value $Value)) { return $null }
    return (Get-IntValue -Value $Value -Default 0)
}

function Test-GraphPermissionSatisfied {
    param(
        [string]$GrantedPermission,
        [string]$RequiredPermission
    )

    $granted = ConvertTo-NormalizedGraphPermissionName -Value $GrantedPermission
    $required = ConvertTo-NormalizedGraphPermissionName -Value $RequiredPermission
    if ([string]::IsNullOrWhiteSpace($granted) -or [string]::IsNullOrWhiteSpace($required)) { return $false }
    if ($granted -eq $required) { return $true }

    # *.ReadWrite.All satisfies the same resource's *.Read.All.
    if ($required -match '^(?<prefix>.+)\.read\.all$' -and $granted -eq ($Matches.prefix + '.readwrite.all')) {
        return $true
    }

    $directoryReadSatisfies = @(
        'organization.read.all',
        'group.read.all',
        'user.read.all',
        'groupmember.read.all',
        'application.read.all'
    )
    if ($required -in $directoryReadSatisfies -and ($granted -eq 'directory.read.all' -or $granted -eq 'directory.readwrite.all')) {
        return $true
    }

    return $false
}

function Get-EnvironmentSecureString {
    param([string]$VariableName)

    if ([string]::IsNullOrWhiteSpace($VariableName)) { return $null }
    $secret = [Environment]::GetEnvironmentVariable($VariableName)
    if ([string]::IsNullOrEmpty($secret)) { return $null }

    $secure = [System.Security.SecureString]::new()
    foreach ($character in $secret.ToCharArray()) {
        $secure.AppendChar($character)
    }
    $secure.MakeReadOnly()
    return $secure
}

function Get-NormalizedCertificateThumbprint {
    param([string]$Thumbprint)

    if ([string]::IsNullOrWhiteSpace($Thumbprint)) { return '' }
    return (($Thumbprint -replace '[^0-9a-fA-F]', '').ToUpperInvariant())
}

function Get-AuthenticationCertificate {
    param(
        [string]$Thumbprint,
        [string]$Path,
        [string]$SecretVariableName
    )

    $resolvedPath = $Path
    if (-not [string]::IsNullOrWhiteSpace($resolvedPath)) {
        if (-not (Test-Path -LiteralPath $resolvedPath)) {
            throw "Certificate file '$resolvedPath' was not found."
        }

        $fullPath = (Resolve-Path -LiteralPath $resolvedPath).Path
        $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
        $secret = $null
        if (-not [string]::IsNullOrWhiteSpace($SecretVariableName)) {
            $secret = [Environment]::GetEnvironmentVariable($SecretVariableName)
            if ([string]::IsNullOrEmpty($secret)) {
                throw "Certificate password environment variable '$SecretVariableName' is not set."
            }
        }

        $certificate = if ([string]::IsNullOrEmpty($secret)) {
            [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($fullPath, $null, $flags)
        }
        else {
            [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($fullPath, $secret, $flags)
        }

        if (-not $certificate.HasPrivateKey) {
            throw "Certificate '$fullPath' does not contain a private key."
        }

        return $certificate
    }

    $normalizedThumbprint = Get-NormalizedCertificateThumbprint -Thumbprint $Thumbprint
    if ([string]::IsNullOrWhiteSpace($normalizedThumbprint)) { return $null }

    foreach ($storeLocation in @(
            [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser,
            [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine
        )) {
        $store = $null
        try {
            $store = [System.Security.Cryptography.X509Certificates.X509Store]::new(
                [System.Security.Cryptography.X509Certificates.StoreName]::My,
                $storeLocation
            )
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
            $certificateMatches = $store.Certificates.Find(
                [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint,
                $normalizedThumbprint,
                $false
            )
            if ($certificateMatches.Count -gt 0) {
                $found = $certificateMatches[0]
                if (-not $found.HasPrivateKey) {
                    throw "Certificate '$normalizedThumbprint' in $storeLocation\My does not contain a private key."
                }
                return $found
            }
        }
        catch {
            if ($_.Exception.Message -match 'does not contain a private key') { throw }
        }
        finally {
            if ($null -ne $store) { $store.Dispose() }
        }
    }

    throw "Certificate thumbprint '$normalizedThumbprint' was not found in CurrentUser\My or LocalMachine\My."
}

function ConvertTo-Base64Url {
    param([Parameter(Mandatory)][byte[]]$Bytes)

    return ([System.Convert]::ToBase64String($Bytes)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Get-CertificateAssertionAccessToken {
    param(
        [Parameter(Mandatory)][string]$Tenant,
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [string]$Scope = 'https://graph.microsoft.com/.default'
    )

    $tokenEndpoint = "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token"
    $x5t = ConvertTo-Base64Url -Bytes $Certificate.GetCertHash()
    $now = [int][DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $headerJson = '{"alg":"RS256","typ":"JWT","x5t":"' + $x5t + '"}'
    $payloadJson = '{"aud":"' + $tokenEndpoint + '","iss":"' + $AppId + '","sub":"' + $AppId + '","jti":"' + [guid]::NewGuid().ToString() + '","nbf":' + $now + ',"exp":' + ($now + 600) + '}'
    $headerSegment = ConvertTo-Base64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes($headerJson))
    $payloadSegment = ConvertTo-Base64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes($payloadJson))
    $signingInput = [System.Text.Encoding]::UTF8.GetBytes("$headerSegment.$payloadSegment")
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if ($null -eq $rsa) { throw 'The authentication certificate does not expose an RSA private key.' }
    $signature = $rsa.SignData($signingInput, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $clientAssertion = "$headerSegment.$payloadSegment.$(ConvertTo-Base64Url -Bytes $signature)"

    try {
        $response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ContentType 'application/x-www-form-urlencoded' -Body @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            scope                 = $Scope
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $clientAssertion
        } -ErrorAction Stop
        return [string]$response.access_token
    }
    catch {
        $detail = $null
        try { $detail = ($_.ErrorDetails.Message | ConvertFrom-Json).error_description } catch { $detail = $null }
        throw "Certificate client-credentials token request failed: $(if ($detail) { $detail } else { $_.Exception.Message })"
    }
}

function Get-GraphErrorCode {
    param([AllowNull()][object]$ErrorRecord)

    $payload = Get-GraphErrorPayload -ErrorRecord $ErrorRecord
    if ($null -eq $payload) { return '' }
    $graphError = Get-ObjectPropertyValue -InputObject $payload -Name 'error' -Default $null
    $code = [string](Get-ObjectPropertyValue -InputObject $graphError -Name 'code' -Default '')
    if ([string]::IsNullOrWhiteSpace($code)) {
        $code = [string](Get-ObjectPropertyValue -InputObject $payload -Name 'code' -Default '')
    }
    return $code
}

function Get-GraphHttpStatusCode {
    param([AllowNull()][object]$ErrorRecord)

    if ($null -eq $ErrorRecord) { return $null }
    $statusCode = $null
    if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.PSObject.Properties.Name -contains 'StatusCode') {
        try { $statusCode = [int]$ErrorRecord.Exception.Response.StatusCode } catch { $statusCode = $null }
    }
    return $statusCode
}

function ConvertTo-GraphReportText {
    param([AllowNull()][object]$Content)

    if ($null -eq $Content) { return '' }

    $text = $null
    if ($Content -is [byte[]]) {
        $text = [System.Text.Encoding]::UTF8.GetString($Content)
    }
    elseif ($Content -is [System.Collections.IList] -and @($Content).Count -gt 0 -and @($Content)[0] -is [byte]) {
        $text = [System.Text.Encoding]::UTF8.GetString(([byte[]]@($Content)))
    }
    else {
        $text = [string]$Content
    }

    if ($text.Length -gt 0 -and [int][char]$text[0] -eq 0xFEFF) {
        $text = $text.Substring(1)
    }

    return $text
}

function ConvertFrom-GraphReportContent {
    param([AllowNull()][object]$Content)

    $text = ConvertTo-GraphReportText -Content $Content
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }

    $trimmed = $text.Trim()
    if ($trimmed.StartsWith('{') -or $trimmed.StartsWith('[')) {
        try {
            $jsonPayload = $trimmed | ConvertFrom-Json -ErrorAction Stop
            $valueProperty = if ($null -ne $jsonPayload) { $jsonPayload.PSObject.Properties['value'] } else { $null }
            if ($null -ne $valueProperty) { return @($valueProperty.Value) }
            return @($jsonPayload)
        }
        catch {
            return @()
        }
    }

    return @($text | ConvertFrom-Csv)
}

function Register-GraphReportDiagnostic {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Message,
        [AllowNull()][object]$StatusCode,
        [string]$Section = 'GraphReport',
        [switch]$DiagnosticOnly
    )

    Initialize-GraphRuntimeState
    $entry = [ordered]@{
        Uri          = $Uri
        Message      = $Message
        StatusCode   = $StatusCode
        TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
        Severity     = if ($DiagnosticOnly) { 'diagnostic' } else { 'error' }
    }
    $script:GraphReportLastError = $entry
    $script:GraphReportDiagnostics.Add($entry) | Out-Null

    # Fallback probes (for example a Copilot period the API rejects before a later
    # period succeeds) stay in the diagnostics list. They are not collector errors.
    if ($DiagnosticOnly) { return }

    $reportVariable = Get-Variable -Name Report -Scope Script -ErrorAction SilentlyContinue
    if ($null -ne $reportVariable -and $null -ne $reportVariable.Value -and $reportVariable.Value.Contains('Errors') -and $null -ne $reportVariable.Value.Errors) {
        $reportVariable.Value.Errors.Add([ordered]@{
            Section      = $Section
            Message      = $Message
            Exception    = 'GraphReportRequestFailed'
            Detail       = $Uri
            StatusCode   = $StatusCode
            TimestampUtc = $entry.TimestampUtc
        }) | Out-Null
    }
}

function Get-GraphReportDiagnosticsForUri {
    param([Parameter(Mandatory)][string]$UriPattern)

    Initialize-GraphRuntimeState
    return @($script:GraphReportDiagnostics | Where-Object { [string]$_.Uri -match $UriPattern })
}

function Get-CopilotPeriodLength {
    param([Parameter(Mandatory)][string]$Period)

    switch ($Period) {
        'D7' { return 7 }
        'D28' { return 28 }
        'D30' { return 30 }
        'D90' { return 90 }
        'D180' { return 180 }
        default { return $null }
    }
}

function Get-CopilotReportCandidateUris {
    # Live headless run, 22 Nov 2026: the first candidate was
    # beta /reports/...UsageUserDetail(period='D30'), and that API returned HTTP 400.
    # /copilot/reports/... accepts D7, D28, D90, and D180, so a requested D30 is
    # tried there as D28 first. The beta /reports/ path still receives the original
    # period as a later fallback. Those fallbacks are diagnostics, not errors.
    # getMicrosoft365CopilotUsageUserCounts is not a Graph report and is intentionally absent.
    param(
        [ValidateSet('D7', 'D30', 'D90', 'D180')]
        [string]$Period
    )

    $copilotPeriod = if ($Period -eq 'D30') { 'D28' } else { $Period }
    return @(
        [ordered]@{ Kind = 'detail'; Uri = "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='$copilotPeriod')" }
        [ordered]@{ Kind = 'detail'; Uri = "https://graph.microsoft.com/beta/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='$copilotPeriod')" }
        [ordered]@{ Kind = 'summary'; Uri = "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUserCountSummary(period='$copilotPeriod')" }
        [ordered]@{ Kind = 'summary'; Uri = "https://graph.microsoft.com/beta/copilot/reports/getMicrosoft365CopilotUserCountSummary(period='$copilotPeriod')" }
        [ordered]@{ Kind = 'detail'; Uri = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='$Period')" }
        [ordered]@{ Kind = 'summary'; Uri = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUserCountSummary(period='$Period')" }
    )
}

function Get-CopilotReportAnchorDate {
    param(
        [AllowNull()][object[]]$Rows,
        [AllowNull()][object]$Fallback
    )

    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $raw = Get-RowValueNormalized -Row $row -CandidateNames @('reportRefreshDate', 'report refresh date', 'reportDate', 'report date')
        if ([string]::IsNullOrWhiteSpace([string]$raw)) { continue }
        try { return ([datetime]$raw).ToUniversalTime() } catch { continue }
    }

    if ($null -ne $Fallback) { return $Fallback.ToUniversalTime() }
    return (Get-Date).ToUniversalTime()
}

function Test-CopilotRowHasActivityDate {
    param([AllowNull()][object]$Row)

    if ($null -eq $Row) { return $false }
    foreach ($property in @($Row.PSObject.Properties)) {
        $normalized = ([string]$property.Name).ToLowerInvariant() -replace '[^a-z0-9]', ''
        if ($normalized -eq 'lastactivitydate') { return $true }
    }
    return $false
}

function Get-CopilotActiveUserCount {
    # Counts enabled rows only when Last Activity Date falls inside the report period.
    # A blank or older activity date is a licensed user who was not active.
    param(
        [AllowNull()][object[]]$Rows,
        [Parameter(Mandatory)][string]$Period,
        [AllowNull()][datetime]$ReportDate
    )

    $rowList = @($Rows | Where-Object { $null -ne $_ })
    $hasActivityColumn = $false
    foreach ($row in $rowList) {
        if (Test-CopilotRowHasActivityDate -Row $row) {
            $hasActivityColumn = $true
            break
        }
    }
    if (-not $hasActivityColumn) { return $null }

    $days = Get-CopilotPeriodLength -Period $Period
    if ($null -eq $days) { return $null }

    $anchor = Get-CopilotReportAnchorDate -Rows $rowList -Fallback $ReportDate
    $cutoff = $anchor.Date.AddDays(-1 * $days)
    $active = 0
    foreach ($row in $rowList) {
        $raw = Get-RowValueNormalized -Row $row -CandidateNames @('lastActivityDate', 'last activity date')
        if ([string]::IsNullOrWhiteSpace([string]$raw)) { continue }
        try { $activity = ([datetime]$raw).ToUniversalTime() } catch { continue }
        if ($activity.Date -ge $cutoff -and $activity.Date -le $anchor.Date) { $active++ }
    }

    return $active
}

function Get-CopilotSummaryActiveUserCount {
    # Uses "active" columns only. "Enabled" columns count licensed users and are ignored.
    param(
        [AllowNull()][object[]]$Rows,
        [Parameter(Mandatory)][string]$Period
    )

    $days = Get-CopilotPeriodLength -Period $Period
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $direct = Get-RowValueNormalized -Row $row -CandidateNames @(
            'anyAppActiveUsers',
            'any app active users',
            'microsoft365CopilotActiveUsers',
            'microsoft 365 copilot active users'
        )
        if ($null -ne $direct -and -not [string]::IsNullOrWhiteSpace([string]$direct)) {
            return (Get-IntValue -Value $direct -Default 0)
        }

        $products = @(Get-ObjectPropertyValue -InputObject $row -Name 'adoptionByProduct' -Default @())
        if (@($products).Count -eq 0) { continue }
        $matched = @($products | Where-Object { (Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'reportPeriod') -Default -1) -eq $days } | Select-Object -First 1)
        if (@($matched).Count -eq 0) { $matched = @($products | Select-Object -First 1) }
        $nested = Get-ObjectPropertyValue -InputObject $matched[0] -Name 'anyAppActiveUsers'
        if ($null -eq $nested) { $nested = Get-ObjectPropertyValue -InputObject $matched[0] -Name 'microsoft365CopilotActiveUsers' }
        if ($null -ne $nested -and -not [string]::IsNullOrWhiteSpace([string]$nested)) {
            return (Get-IntValue -Value $nested -Default 0)
        }
    }

    return $null
}

function Get-SecureScoreAddressedState {
    # Same states as Analysis/Secure Score/o365-secure-score-extract.ps1.
    # A profile is addressed when any controlStateUpdates entry uses one of them.
    param(
        [AllowNull()][object]$Control,
        [System.Collections.Generic.HashSet[string]]$AddressedStates
    )

    $updates = Get-ObjectPropertyValue -InputObject $Control -Name 'controlStateUpdates' -Default $null
    if ($null -eq $updates -or $updates -is [string]) { return $null }

    foreach ($update in @($updates)) {
        if ($null -eq $update) { continue }
        $state = [string](Get-ObjectPropertyValue -InputObject $update -Name 'state' -Default '')
        if (-not [string]::IsNullOrWhiteSpace($state) -and $AddressedStates.Contains($state)) { return $state }
    }

    return $null
}

function Test-SecureScoreControlAddressed {
    param(
        [AllowNull()][object]$Control,
        [System.Collections.Generic.HashSet[string]]$AddressedStates
    )

    return $null -ne (Get-SecureScoreAddressedState -Control $Control -AddressedStates $AddressedStates)
}

function Get-SecureScoreRankSortKey {
    param([AllowNull()][object]$Control)

    $rankValue = Get-ObjectPropertyValue -InputObject $Control -Name 'rank' -Default $null
    if ($null -eq $rankValue) { return [int]::MaxValue }
    $parsed = 0
    if (-not [int]::TryParse([string]$rankValue, [ref]$parsed)) { return [int]::MaxValue }
    if ($parsed -le 0) { return [int]::MaxValue }
    return $parsed
}

function Get-SecureScoreOpenControlAssessment {
    # Same join as Analysis/Secure Score/o365-secure-score-extract.ps1:
    # controlScores.controlName equals the control profile id. Open means the
    # control was scored, pointsAvailable (max - current, floored at 0) is > 0,
    # and the profile is not addressed. Catalog profiles absent from the latest
    # score are not applicable, so they are not open.
    param(
        [AllowNull()][object[]]$ControlProfiles,
        [AllowNull()][object[]]$ControlScores
    )

    $addressedStates = @(
        'completed',
        'ignored',
        'thirdParty',
        'riskAccepted',
        'alternativeMitigation'
    )
    $addressedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($stateName in $addressedStates) { $null = $addressedSet.Add($stateName) }

    $scoreByName = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($scoreEntry in @($ControlScores)) {
        if ($null -eq $scoreEntry) { continue }
        $scoredName = ([string](Get-ObjectPropertyValue -InputObject $scoreEntry -Name 'controlName' -Default '')).Trim()
        if ([string]::IsNullOrWhiteSpace($scoredName)) { continue }
        if (-not $scoreByName.ContainsKey($scoredName)) { $scoreByName[$scoredName] = $scoreEntry }
    }

    $profileById = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $notApplicableCount = 0
    foreach ($controlProfile in @($ControlProfiles)) {
        if ($null -eq $controlProfile) {
            $notApplicableCount++
            continue
        }
        $profileId = ([string](Get-ObjectPropertyValue -InputObject $controlProfile -Name 'id' -Default '')).Trim()
        if ([string]::IsNullOrWhiteSpace($profileId) -or -not $scoreByName.ContainsKey($profileId)) {
            $notApplicableCount++
            continue
        }
        if (-not $profileById.ContainsKey($profileId)) { $profileById[$profileId] = $controlProfile }
    }

    $openList = [System.Collections.Generic.List[object]]::new()
    $gapList = [System.Collections.Generic.List[object]]::new()
    $fullyAchievedCount = 0
    $addressedWithPointsRemaining = 0
    $controlsBelowMax = 0
    $totalPoints = [double]0

    foreach ($controlName in @($scoreByName.Keys)) {
        $scoreEntry = $scoreByName[$controlName]
        $controlProfile = if ($profileById.ContainsKey($controlName)) { $profileById[$controlName] } else { $null }
        $currentScore = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $scoreEntry -Name 'score') -Default -1
        if ($currentScore -lt 0) {
            $currentScore = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $scoreEntry -Name 'currentScore') -Default 0
        }

        $maxScore = $null
        if ($null -ne $controlProfile) {
            $profileMax = Get-ObjectPropertyValue -InputObject $controlProfile -Name 'maxScore' -Default $null
            if ($null -ne $profileMax -and -not [string]::IsNullOrWhiteSpace([string]$profileMax)) {
                $maxScore = Get-DoubleValue -Value $profileMax -Default 0
            }
        }
        if ($null -eq $maxScore) {
            $scoreMax = Get-ObjectPropertyValue -InputObject $scoreEntry -Name 'maxScore' -Default $null
            $maxScore = if ($null -ne $scoreMax -and -not [string]::IsNullOrWhiteSpace([string]$scoreMax)) {
                Get-DoubleValue -Value $scoreMax -Default 0
            }
            else {
                [double]0
            }
        }

        $pointsAvailable = [math]::Round(($maxScore - $currentScore), 2)
        if ($pointsAvailable -lt 0) { $pointsAvailable = [double]0 }
        $totalPoints += $pointsAvailable
        if ($pointsAvailable -gt 0) { $controlsBelowMax++ }

        $addressedState = $null
        if ($null -ne $controlProfile) {
            $addressedState = Get-SecureScoreAddressedState -Control $controlProfile -AddressedStates $addressedSet
        }
        $isAddressed = -not [string]::IsNullOrWhiteSpace([string]$addressedState)

        $gapClassification = 'fully-achieved'
        if ($pointsAvailable -gt 0 -and $isAddressed) {
            $gapClassification = 'addressed-with-points-remaining'
        }
        elseif ($pointsAvailable -gt 0) {
            $gapClassification = 'open'
        }

        $openRecord = [pscustomobject][ordered]@{
            id                          = if ($null -ne $controlProfile) { Get-ObjectPropertyValue -InputObject $controlProfile -Name 'id' -Default $controlName } else { $controlName }
            title                       = if ($null -ne $controlProfile) { Get-ObjectPropertyValue -InputObject $controlProfile -Name 'title' -Default 'not collected' } else { Get-ObjectPropertyValue -InputObject $scoreEntry -Name 'description' -Default $controlName }
            rank                        = if ($null -ne $controlProfile) { Get-ObjectPropertyValue -InputObject $controlProfile -Name 'rank' -Default $null } else { $null }
            maxScore                    = $maxScore
            currentScore                = $currentScore
            pointsAvailable             = $pointsAvailable
            controlCategory             = if ($null -ne $controlProfile) { Get-ObjectPropertyValue -InputObject $controlProfile -Name 'controlCategory' -Default 'not collected' } else { Get-ObjectPropertyValue -InputObject $scoreEntry -Name 'controlCategory' -Default 'not collected' }
            implementationCost          = if ($null -ne $controlProfile) { Get-ObjectPropertyValue -InputObject $controlProfile -Name 'implementationCost' -Default 'not collected' } else { 'not collected' }
            userImpact                  = if ($null -ne $controlProfile) { Get-ObjectPropertyValue -InputObject $controlProfile -Name 'userImpact' -Default 'not collected' } else { 'not collected' }
            service                     = if ($null -ne $controlProfile) { Get-ObjectPropertyValue -InputObject $controlProfile -Name 'service' -Default 'not collected' } else { 'not collected' }
            remediation                 = if ($null -ne $controlProfile) { Get-ObjectPropertyValue -InputObject $controlProfile -Name 'remediation' -Default 'not collected' } else { 'not collected' }
            actionUrl                   = if ($null -ne $controlProfile) { Get-ObjectPropertyValue -InputObject $controlProfile -Name 'actionUrl' -Default 'not collected' } else { 'not collected' }
            gapClassification           = $gapClassification
            addressedState              = if ($isAddressed) { $addressedState } else { '' }
            addressedWithPointsRemaining = ($gapClassification -eq 'addressed-with-points-remaining')
        }

        if ($pointsAvailable -gt 0 -and -not $isAddressed) {
            $openList.Add($openRecord) | Out-Null
            $gapList.Add($openRecord) | Out-Null
        }
        elseif ($pointsAvailable -gt 0 -and $isAddressed) {
            # Ignored, risk-accepted, and other addressed controls can still have
            # points available (for example MFARegistrationV2). They stay out of the
            # open count and stay on the gap list with an explicit flag.
            $addressedWithPointsRemaining++
            $gapList.Add($openRecord) | Out-Null
        }
        else {
            $fullyAchievedCount++
        }
    }

    $sortProperties = @(
        @{
            Expression = { Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'pointsAvailable') -Default 0 }
            Descending = $true
        }, @{
            Expression = { Get-SecureScoreRankSortKey -Control $_ }
            Descending = $false
        }
    )
    $openArray = @($openList.ToArray() | Sort-Object -Property $sortProperties)
    $gapArray = @($gapList.ToArray() | Sort-Object -Property $sortProperties)

    return [ordered]@{
        ScoredControlCount           = $scoreByName.Count
        OpenControlCount             = @($openArray).Count
        FullyAchievedCount           = $fullyAchievedCount
        AddressedWithPointsRemaining = $addressedWithPointsRemaining
        NotApplicableCount           = $notApplicableCount
        ControlsBelowMaxScore        = $controlsBelowMax
        TotalPointsAvailable         = [math]::Round($totalPoints, 2)
        Open                         = @($openArray)
        TopOpen                      = @($openArray | Select-Object -First 10)
        Gaps                         = @($gapArray)
        TopGaps                      = @($gapArray | Select-Object -First 10)
    }
}

function Get-SensitivityLabelSnapshot {
    param(
        [bool]$GraphCollected = $false,
        [AllowNull()][object[]]$GraphLabels,
        [bool]$ComplianceSessionConnected = $false,
        [AllowNull()][object[]]$ComplianceLabels,
        [AllowNull()][object[]]$ComplianceLabelPolicies,
        [AllowNull()][object[]]$AutoSensitivityLabelPolicies,
        [string]$GraphFailureReason = ''
    )

    $published = $null
    $total = $null
    $autoLabel = $null
    $policyPublished = $null
    $policyTotal = $null
    $source = 'not collected'
    $collected = $false
    $reason = ''

    if ($GraphCollected) {
        $collected = $true
        $source = 'graph/informationProtection/policy/labels'
        $labelList = @($GraphLabels | Where-Object { $null -ne $_ })
        $total = @($labelList).Count
        $published = @($labelList | Where-Object {
            $isPublished = Get-ObjectPropertyValue -InputObject $_ -Name 'isPublished' -Default $null
            $isActive = Get-ObjectPropertyValue -InputObject $_ -Name 'isActive' -Default $null
            ($isPublished -eq $true) -or ($null -eq $isPublished -and $isActive -eq $true)
        }).Count
        $autoLabel = @($labelList | Where-Object {
            (Get-ObjectPropertyValue -InputObject $_ -Name 'isLabelAutoClassified') -eq $true -or
            (Get-ObjectPropertyValue -InputObject $_ -Name 'isDefault') -eq $true
        }).Count
    }
    elseif ($ComplianceSessionConnected -and ($null -ne $ComplianceLabels -or $null -ne $ComplianceLabelPolicies -or $null -ne $AutoSensitivityLabelPolicies)) {
        $collected = $true
        $source = 'ipps/Get-Label'
        if ($null -ne $ComplianceLabels) { $total = @($ComplianceLabels | Where-Object { $null -ne $_ }).Count }
        if ($null -ne $ComplianceLabelPolicies) {
            $publishedNames = [System.Collections.Generic.List[string]]::new()
            $policyList = @($ComplianceLabelPolicies | Where-Object { $null -ne $_ })
            $policyTotal = @($policyList).Count
            $policyPublished = 0
            foreach ($policy in $policyList) {
                $labelNames = @(Get-ObjectPropertyValue -InputObject $policy -Name 'Labels' -Default @())
                $policyPublishesLabel = $false
                foreach ($labelName in $labelNames) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$labelName)) {
                        $publishedNames.Add([string]$labelName)
                        $policyPublishesLabel = $true
                    }
                }
                if ($policyPublishesLabel) { $policyPublished++ }
            }
            $published = @($publishedNames | Sort-Object -Unique).Count
        }
        if ($null -ne $AutoSensitivityLabelPolicies) {
            $autoLabel = @($AutoSensitivityLabelPolicies | Where-Object { $null -ne $_ }).Count
        }
        if ($null -eq $published) {
            $reason = 'Purview label policies were not returned, so the published label count is not collected.'
        }
    }
    else {
        $reason = if (-not [string]::IsNullOrWhiteSpace($GraphFailureReason)) {
            $GraphFailureReason
        }
        elseif ($ComplianceSessionConnected) {
            'Sensitivity labels were not collected from Graph or the Purview session.'
        }
        else {
            'Sensitivity labels were not collected. Graph label read was unavailable and no Purview compliance session was connected.'
        }
    }

    return [ordered]@{
        Collected       = $collected
        Source          = $source
        Published       = $published
        Total           = $total
        AutoLabel       = $autoLabel
        PolicyPublished = $policyPublished
        PolicyTotal     = $policyTotal
        Reason          = $reason
    }
}

function Get-MfaRegistrationCohort {
    param([AllowNull()][object[]]$RegistrationDetails)

    $rows = @($RegistrationDetails | Where-Object { $null -ne $_ })
    $userTypeAvailable = $false
    $adminFlagAvailable = $false
    $memberPopulation = 0
    $memberRegistered = 0
    $guestPopulation = 0
    $guestRegistered = 0
    $adminPopulation = 0
    $adminRegistered = 0
    $registered = 0
    $capable = 0
    $passwordless = 0

    foreach ($row in $rows) {
        $isRegistered = (Get-ObjectPropertyValue -InputObject $row -Name 'isMfaRegistered' -Default $false) -eq $true
        $isCapable = (Get-ObjectPropertyValue -InputObject $row -Name 'isMfaCapable' -Default $false) -eq $true
        $isPasswordless = (Get-ObjectPropertyValue -InputObject $row -Name 'isPasswordlessCapable' -Default $false) -eq $true
        if ($isRegistered) { $registered++ }
        if ($isCapable) { $capable++ }
        if ($isPasswordless) { $passwordless++ }

        $userTypeProperty = $row.PSObject.Properties['userType']
        if ($null -eq $userTypeProperty) { $userTypeProperty = $row.PSObject.Properties['UserType'] }
        if ($null -ne $userTypeProperty) {
            $userTypeAvailable = $true
            $userType = [string]$userTypeProperty.Value
            if ($userType -eq 'guest') {
                $guestPopulation++
                if ($isRegistered) { $guestRegistered++ }
            }
            elseif ($userType -eq 'member') {
                $memberPopulation++
                if ($isRegistered) { $memberRegistered++ }
            }
        }

        $adminProperty = $row.PSObject.Properties['isAdmin']
        if ($null -eq $adminProperty) { $adminProperty = $row.PSObject.Properties['IsAdmin'] }
        if ($null -ne $adminProperty) {
            $adminFlagAvailable = $true
            if ($adminProperty.Value -eq $true) {
                $adminPopulation++
                if ($isRegistered) { $adminRegistered++ }
            }
        }
    }

    return [ordered]@{
        TotalCount          = $rows.Count
        RegisteredCount     = $registered
        CapableCount        = $capable
        PasswordlessCount   = $passwordless
        UserTypeAvailable   = $userTypeAvailable
        AdminFlagAvailable  = $adminFlagAvailable
        MemberPopulation    = if ($userTypeAvailable) { $memberPopulation } else { $null }
        MemberRegistered    = if ($userTypeAvailable) { $memberRegistered } else { $null }
        GuestPopulation     = if ($userTypeAvailable) { $guestPopulation } else { $null }
        GuestRegistered     = if ($userTypeAvailable) { $guestRegistered } else { $null }
        AdminPopulation     = if ($adminFlagAvailable) { $adminPopulation } else { $null }
        AdminRegistered     = if ($adminFlagAvailable) { $adminRegistered } else { $null }
    }
}

function Get-RegistrationPercent {
    param(
        [AllowNull()][object]$Registered,
        [AllowNull()][object]$Population
    )

    if ($null -eq $Registered -or $null -eq $Population) { return 'not collected' }
    $denominator = Get-IntValue -Value $Population -Default 0
    if ($denominator -le 0) { return 'not collected' }
    return [math]::Round(((Get-IntValue -Value $Registered -Default 0) / $denominator) * 100, 2)
}

function Get-EnabledMemberMfaRegistration {
    # Enabled members come from the directory. Registration details do not
    # carry accountEnabled, so guests and disabled members in that report
    # must not set the percentage used for scoring.
    param([AllowNull()][object[]]$RegistrationDetails)

    $enabledMembers = @(Get-EnabledMemberUsers)
    $population = @($enabledMembers).Count
    $enabledIds = @{}
    foreach ($user in @($enabledMembers)) {
        $id = Get-DirectoryObjectId -User $user
        if (-not [string]::IsNullOrWhiteSpace($id)) {
            $enabledIds[$id.ToLowerInvariant()] = $true
        }
    }

    $registeredIds = @{}
    foreach ($row in @($RegistrationDetails)) {
        if ($null -eq $row) { continue }
        $rowId = Get-DirectoryObjectId -User $row
        if ([string]::IsNullOrWhiteSpace($rowId)) { continue }
        $key = $rowId.ToLowerInvariant()
        if (-not $enabledIds.ContainsKey($key)) { continue }

        $isRegistered = (Get-ObjectPropertyValue -InputObject $row -Name 'isMfaRegistered' -Default $false) -eq $true
        if (-not $isRegistered) {
            $isRegistered = (Get-ObjectPropertyValue -InputObject $row -Name 'IsMfaRegistered' -Default $false) -eq $true
        }
        if ($isRegistered) { $registeredIds[$key] = $true }
    }

    $registered = @($registeredIds.Keys).Count
    return [ordered]@{
        Population = $population
        Registered = $registered
        Percent    = if ($population -gt 0) { [math]::Round(($registered / $population) * 100, 2) } else { 'not collected' }
    }
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
    # Keep the interactive default delegated scope set explicit so the startup guidance,
    # consent expectations, and requested permissions stay aligned.
    return @(
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
            # Tenant sharingCapability is weaker evidence than a site inventory.
            TenantOnlyAnyoneLinkCompletenessFactor = 0.5
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

    $providedScopes = [System.Collections.Generic.List[string]]::new()
    $ctx = $null
    try { $ctx = Get-MgContext -ErrorAction SilentlyContinue } catch { $ctx = $null }

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

    $missing = foreach ($requiredScope in $requiredScopeList) {
        $covered = $false
        foreach ($grantedScope in $providedScopes) {
            if (Test-GraphPermissionSatisfied -GrantedPermission $grantedScope -RequiredPermission $requiredScope) {
                $covered = $true
                break
            }
        }
        if (-not $covered) { $requiredScope }
    }

    return @($missing | Sort-Object -Unique)
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

    if ([string]::IsNullOrWhiteSpace($script:GraphClientId)) { return $false }

    $clientServicePrincipalId = Get-GraphClientServicePrincipalId -ClientAppId $script:GraphClientId
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

        $consentUrl = Get-GraphAdminConsentUrl -TenantId $script:TenantId -ClientId $script:GraphClientId
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
        $percentComplete = [math]::Min(100, [int][math]::Round(($script:CollectorStepIndex / $script:CollectorStepTotal) * 100, 0))
        $progressParams = @{
            Id              = 1
            Activity        = 'Microsoft 365 AI Readiness Export'
            Status          = ("[{0}/{1}] Reading section '{2}'" -f $script:CollectorStepIndex, $script:CollectorStepTotal, $Section)
            PercentComplete = $percentComplete
        }
        Write-Progress @progressParams
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
function Get-GraphCertificateMaterialRequested {
    $thumbprint = Get-ScriptStringParameter -Name 'GraphCertificateThumbprint'
    $path = Get-ScriptStringParameter -Name 'GraphCertificatePath'
    return (-not [string]::IsNullOrWhiteSpace($thumbprint) -or -not [string]::IsNullOrWhiteSpace($path))
}

function Connect-GraphAppOnlyCertificate {
    param(
        [string]$Tenant = 'organizations',
        [string]$ClientId
    )

    if ([string]::IsNullOrWhiteSpace($Tenant) -or $Tenant -eq 'organizations') {
        throw 'App-only certificate authentication requires -TenantId.'
    }
    if ([string]::IsNullOrWhiteSpace($ClientId) -or $ClientId -eq '04b07795-8ddb-461a-bbee-02f9e1bf7b46') {
        throw 'App-only certificate authentication requires -GraphClientId for your app registration. The Microsoft public client cannot use client credentials.'
    }

    $certificate = Get-AuthenticationCertificate -Thumbprint (Get-ScriptStringParameter -Name 'GraphCertificateThumbprint') -Path (Get-ScriptStringParameter -Name 'GraphCertificatePath') -SecretVariableName (Get-ScriptStringParameter -Name 'GraphCertificatePasswordEnvironmentVariable')
    $script:GraphAuthenticationCertificate = $certificate
    $script:GraphAppClientId = $ClientId
    $token = Get-CertificateAssertionAccessToken -Tenant $Tenant -AppId $ClientId -Certificate $certificate
    $script:GraphAccessToken = $token
    $script:GraphAccessTokenExpiresAtUtc = Get-GraphTokenExpirationUtc -AccessToken $token
    $script:GraphAuthenticationMode = 'AppOnlyCertificate'

    $claims = Get-GraphTokenClaims -AccessToken $token
    if ($claims) {
        foreach ($claimName in @('app_displayname', 'appDisplayName', 'name')) {
            if ($claims.ContainsKey($claimName) -and -not [string]::IsNullOrWhiteSpace([string]$claims[$claimName])) {
                $script:GraphAppDisplayName = [string]$claims[$claimName]
                break
            }
        }
        if ([string]::IsNullOrWhiteSpace($script:GraphAppClientId)) {
            foreach ($claimName in @('appid', 'azp')) {
                if ($claims.ContainsKey($claimName) -and -not [string]::IsNullOrWhiteSpace([string]$claims[$claimName])) {
                    $script:GraphAppClientId = [string]$claims[$claimName]
                    break
                }
            }
        }
    }

    return $token
}

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

    if (Get-GraphCertificateMaterialRequested) {
        return (Connect-GraphAppOnlyCertificate -Tenant $TenantId -ClientId $ClientId)
    }

    if ($script:GraphAuthenticationMode -eq 'ExistingSession') {
        return $null
    }

    if (Test-NonInteractiveMode) {
        throw 'NonInteractive/Headless mode will not start device-code sign-in. Connect with an existing Microsoft Graph session or pass -GraphClientId and -GraphCertificateThumbprint or -GraphCertificatePath.'
    }

    $deviceCodeUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode"
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $scopeString = $Scopes -join ' '

    $deviceCodeResponse = Invoke-RestMethod -Method Post -Uri $deviceCodeUrl -Body @{ client_id = $ClientId; scope = $scopeString } -ErrorAction Stop

    $verificationUri = [string]$deviceCodeResponse.verification_uri
    $userCode = [string]$deviceCodeResponse.user_code

    if (-not [string]::IsNullOrWhiteSpace($userCode)) {
        if (Test-NonInteractiveMode) {
            throw 'NonInteractive/Headless mode will not copy a device code or wait for interactive sign-in.'
        }
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
        if (Test-NonInteractiveMode) {
            throw 'NonInteractive/Headless mode will not open a browser.'
        }
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
    $script:GraphAuthenticationMode = 'DeviceCode'
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
    foreach ($user in $signInUsers) {
        if ($null -eq $user) { continue }
        $idProp = $user.PSObject.Properties['Id']
        if ($null -eq $idProp -or [string]::IsNullOrWhiteSpace([string]$idProp.Value)) { continue }

        $signInActProp = $user.PSObject.Properties['SignInActivity']
        $signInById[[string]$idProp.Value] = if ($null -ne $signInActProp) { $signInActProp.Value } else { $null }
    }

    $baseList = $script:GraphUsersCache[$baseCacheKey]
    $mergedUsers = [System.Collections.Generic.List[object]]::new($baseList.Count)
    foreach ($baseUser in $baseList) {
        $mergedMap = [ordered]@{}
        foreach ($property in $baseUser.PSObject.Properties) {
            $mergedMap[$property.Name] = $property.Value
        }

        $signInActivityValue = $null
        $baseIdProp = $baseUser.PSObject.Properties['Id']
        if ($null -ne $baseIdProp -and $null -ne $baseIdProp.Value) {
            $baseIdText = [string]$baseIdProp.Value
            if ($signInById.ContainsKey($baseIdText)) {
                $signInActivityValue = $signInById[$baseIdText]
            }
        }

        $mergedMap['SignInActivity'] = $signInActivityValue
        $mergedUsers.Add([pscustomobject]$mergedMap)
    }

    $mergedArray = $mergedUsers.ToArray()
    $script:GraphUsersCache[$signInCacheKey] = $mergedArray
    return $mergedArray
}

function Get-WebResponseBody {
    param([AllowNull()][object]$Response)

    if ($null -eq $Response) { return $null }
    $content = $Response.Content
    $hasBytes = ($content -is [byte[]] -and $content.Length -gt 0) -or ($content -is [System.Collections.IList] -and @($content).Count -gt 0 -and @($content)[0] -is [byte])
    $hasText = ($content -is [string] -and -not [string]::IsNullOrWhiteSpace($content))
    if ($hasBytes -or $hasText) { return $content }

    if ($Response.PSObject.Properties.Name -contains 'RawContentStream' -and $null -ne $Response.RawContentStream) {
        $stream = $Response.RawContentStream
        if ($stream.CanSeek) { $stream.Position = 0 }
        $memory = [System.IO.MemoryStream]::new()
        try {
            $stream.CopyTo($memory)
            return $memory.ToArray()
        }
        finally {
            $memory.Dispose()
        }
    }

    return $content
}

function Get-GraphReportRows {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Query = @{},
        [switch]$DiagnosticOnly
    )

    Initialize-GraphRuntimeState
    $requestUri = Add-GraphQueryString -Uri $Uri -Query $Query
    if ($script:GraphReportRowsCache.ContainsKey($requestUri)) {
        return @($script:GraphReportRowsCache[$requestUri])
    }

    $token = $null
    try { $token = Get-GraphAccessToken } catch {
        Register-GraphReportDiagnostic -Uri $requestUri -Message $_.Exception.Message -StatusCode $null -DiagnosticOnly:$DiagnosticOnly
        return @()
    }

    $maxAttempts = 5
    for ($attempt = 0; $attempt -lt $maxAttempts; $attempt++) {
        try {
            $body = $null
            if (-not [string]::IsNullOrWhiteSpace([string]$token)) {
                $response = Invoke-WebRequest -Method Get -Uri $requestUri -Headers @{ Authorization = "Bearer $token"; Accept = 'text/csv' } -ErrorAction Stop
                $body = Get-WebResponseBody -Response $response
            }
            elseif (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
                $body = Invoke-MgGraphRequest -Method GET -Uri $requestUri -Headers @{ Accept = 'text/csv' } -ErrorAction Stop
            }
            else {
                Register-GraphReportDiagnostic -Uri $requestUri -Message 'No Graph access token or existing Microsoft Graph session was available for the report request.' -StatusCode $null -DiagnosticOnly:$DiagnosticOnly
                return @()
            }

            $preview = ConvertTo-GraphReportText -Content $body
            if ([string]::IsNullOrWhiteSpace($preview)) { return @() }

            # Some report endpoints return a pre-authenticated download URL as the response body.
            if ($preview.Trim() -match '^https?://') {
                $downloadResponse = Invoke-WebRequest -Method Get -Uri $preview.Trim() -ErrorAction Stop
                $body = Get-WebResponseBody -Response $downloadResponse
            }

            $rows = @(ConvertFrom-GraphReportContent -Content $body)
            $script:GraphReportRowsCache[$requestUri] = $rows
            $script:GraphReportLastError = $null
            return @($rows)
        }
        catch {
            $statusCode = Get-GraphHttpStatusCode -ErrorRecord $_
            $script:GraphLastHttpStatus = $statusCode

            if (($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)) -and $attempt -lt ($maxAttempts - 1)) {
                $delaySeconds = Get-GraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt
                Start-Sleep -Seconds $delaySeconds
                continue
            }

            Register-GraphReportDiagnostic -Uri $requestUri -Message $_.Exception.Message -StatusCode $statusCode -DiagnosticOnly:$DiagnosticOnly
            return @()
        }
    }

    Register-GraphReportDiagnostic -Uri $requestUri -Message 'Graph report request did not return rows.' -StatusCode $script:GraphLastHttpStatus -DiagnosticOnly:$DiagnosticOnly
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
        [ValidateSet('D7', 'D30', 'D90', 'D180')]
        [string]$Period = 'D30'
    )

    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('Reports.Read.All'))) {
        return [ordered]@{
            Status      = 'not collected'
            Reason      = 'Missing Graph permission: Reports.Read.All'
            ActiveUsers = 'not collected'
            ReportDate  = 'not collected'
            RecordCount = 0
            Source      = 'not collected'
            ReportDiagnostics = @()
        }
    }

    if ($script:CopilotUsageCache.ContainsKey($Period)) {
        return $script:CopilotUsageCache[$Period]
    }

    $candidateReports = @(Get-CopilotReportCandidateUris -Period $Period)
    $attemptDiagnostics = [System.Collections.Generic.List[string]]::new()

    foreach ($candidate in $candidateReports) {
        $reportUri = [string]$candidate.Uri
        # A rejected period is a fallback probe. Record it as a diagnostic so a later
        # successful candidate is not reported as a failed collector.
        $rows = @(Get-GraphReportRows -Uri $reportUri -DiagnosticOnly)
        if (@($rows).Count -eq 0) {
            $diagnostic = Get-GraphReportDiagnosticsForUri -UriPattern ([regex]::Escape($reportUri)) | Select-Object -Last 1
            if ($diagnostic) { $attemptDiagnostics.Add("$reportUri : $($diagnostic.Message)") | Out-Null }
            continue
        }

        $reportDate = Get-CopilotReportAnchorDate -Rows $rows -Fallback $null
        $activeUsers = $null
        $source = $null
        if ([string]$candidate.Kind -eq 'detail') {
            $activeUsers = Get-CopilotActiveUserCount -Rows $rows -Period $Period -ReportDate $reportDate
            $source = 'microsoft365CopilotUsageUserDetail'
        }
        if ($null -eq $activeUsers) {
            $activeUsers = Get-CopilotSummaryActiveUserCount -Rows $rows -Period $Period
            if ($null -ne $activeUsers) { $source = 'microsoft365CopilotUserCountSummary' }
        }
        if ($null -eq $activeUsers) { continue }

        $result = [ordered]@{
            Status      = 'reported'
            Reason      = ''
            ActiveUsers = [int]$activeUsers
            ReportDate  = if ($reportDate) { $reportDate.ToString('yyyy-MM-dd') } else { 'not collected' }
            RecordCount = @($rows).Count
            Source      = $source
            ReportDiagnostics = @($attemptDiagnostics)
        }
        $script:CopilotUsageCache[$Period] = $result
        return $result
    }

    $diagnosticReason = if ($attemptDiagnostics.Count -gt 0) { ' ' + ($attemptDiagnostics -join ' | ') } else { '' }
    $result = [ordered]@{
        Status      = 'not collected'
        Reason      = "Copilot usage report endpoint returned no data (not enabled, unsupported, or not licensed in this tenant).$diagnosticReason"
        ActiveUsers = 'not collected'
        ReportDate  = 'not collected'
        RecordCount = 0
        Source      = 'not collected'
        ReportDiagnostics = @($attemptDiagnostics)
    }
    $script:CopilotUsageCache[$Period] = $result
    return $result
}

function Get-GraphRest {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Query = @{}
    )

    $token = $null
    try { $token = Get-GraphAccessToken } catch { $token = $null }
    $existingGraphContext = $null
    if ([string]::IsNullOrWhiteSpace([string]$token)) {
        try { $existingGraphContext = Get-MgContext -ErrorAction SilentlyContinue } catch { $existingGraphContext = $null }
        if ($null -eq $existingGraphContext -or -not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
            $script:GraphLastHttpStatus = $null
            return $null
        }
    }

    $requestUri = Add-GraphQueryString -Uri $Uri -Query $Query

    $maxAttempts = 5
    for ($attempt = 0; $attempt -lt $maxAttempts; $attempt++) {
        try {
            $result = $null
            if (-not [string]::IsNullOrWhiteSpace([string]$token)) {
                $result = Invoke-RestMethod -Method Get -Uri $requestUri -Headers @{ Authorization = "Bearer $token" } -ErrorAction Stop
            }
            else {
                $result = Invoke-MgGraphRequest -Method GET -Uri $requestUri -ErrorAction Stop
            }
            $script:GraphLastHttpStatus = 200
            $script:GraphLastErrorMessage = $null
            $script:GraphLastErrorCode = $null
            return $result
        }
        catch {
            $message = $_.Exception.Message
            $statusCode = Get-GraphHttpStatusCode -ErrorRecord $_
            $script:GraphLastHttpStatus = $statusCode
            $script:GraphLastErrorMessage = [string]$message
            $script:GraphLastErrorCode = Get-GraphErrorCode -ErrorRecord $_

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
            $valueProp = $response.PSObject.Properties['value']
            if ($null -ne $valueProp) {
                foreach ($item in @($valueProp.Value)) {
                    if ($null -ne $item) { $items.Add([pscustomobject]$item) }
                }
                $nextProp = $response.PSObject.Properties['@odata.nextLink']
                $nextUri = if ($null -ne $nextProp -and $null -ne $nextProp.Value) { [string]$nextProp.Value } else { $null }
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

function Get-DirectoryObjectId {
    param([AllowNull()][object]$User)

    if ($null -eq $User) { return '' }
    foreach ($name in @('id', 'Id')) {
        $value = [string](Get-ObjectPropertyValue -InputObject $User -Name $name -Default '')
        if (-not [string]::IsNullOrWhiteSpace($value)) { return $value }
    }
    return ''
}

function Test-IsEnabledMemberUser {
    # Same population as the licence denominator: guests are excluded, and a
    # disabled account is excluded only when accountEnabled is present.
    param([AllowNull()][object]$User)

    if ($null -eq $User) { return $false }
    $userType = [string](Get-ObjectPropertyValue -InputObject $User -Name 'userType' -Default '')
    if ([string]::IsNullOrWhiteSpace($userType)) {
        $userType = [string](Get-ObjectPropertyValue -InputObject $User -Name 'UserType' -Default '')
    }
    if ($userType.Equals('Guest', [System.StringComparison]::OrdinalIgnoreCase)) { return $false }

    $accountEnabled = Get-ObjectPropertyValue -InputObject $User -Name 'accountEnabled' -Default $null
    if ($null -eq $accountEnabled) {
        $accountEnabled = Get-ObjectPropertyValue -InputObject $User -Name 'AccountEnabled' -Default $null
    }
    if ($accountEnabled -eq $false) { return $false }
    return $true
}

function Get-EnabledMemberUsers {
    $eligible = [System.Collections.Generic.List[object]]::new()
    foreach ($candidateUser in @(Get-DirectoryUsersForReadiness)) {
        if (Test-IsEnabledMemberUser -User $candidateUser) {
            $eligible.Add($candidateUser) | Out-Null
        }
    }

    # Callers wrap this in @(), which keeps an empty or single-user result as a list.
    return $eligible.ToArray()
}

function Get-WorkloadCertificateSelection {
    param(
        [string]$AppId,
        [string]$Organization,
        [string]$Thumbprint,
        [string]$Path,
        [string]$SecretVariableName,
        [switch]$AllowGraphFallback
    )

    $selectedAppId = $AppId
    $selectedOrganization = if (-not [string]::IsNullOrWhiteSpace($Organization)) { $Organization } else { (Get-ScriptStringParameter -Name 'TenantId') }
    $selectedThumbprint = $Thumbprint
    $selectedPath = $Path
    $selectedPasswordVariable = $SecretVariableName

    $hasExplicitCertificate = -not [string]::IsNullOrWhiteSpace($selectedThumbprint) -or -not [string]::IsNullOrWhiteSpace($selectedPath)
    # Parentheses are required. Without them PowerShell binds -or as an argument
    # of Test-NonInteractiveMode and the AppOnlyCertificate fallback never runs.
    if (-not $hasExplicitCertificate -and $AllowGraphFallback -and ((Test-NonInteractiveMode) -or $script:GraphAuthenticationMode -eq 'AppOnlyCertificate')) {
        $selectedThumbprint = Get-ScriptStringParameter -Name 'GraphCertificateThumbprint'
        $selectedPath = Get-ScriptStringParameter -Name 'GraphCertificatePath'
        $selectedPasswordVariable = Get-ScriptStringParameter -Name 'GraphCertificatePasswordEnvironmentVariable'
        if ([string]::IsNullOrWhiteSpace($selectedAppId)) { $selectedAppId = Get-ScriptStringParameter -Name 'GraphClientId' }
        $hasExplicitCertificate = -not [string]::IsNullOrWhiteSpace($selectedThumbprint) -or -not [string]::IsNullOrWhiteSpace($selectedPath)
    }

    if (-not $hasExplicitCertificate -or [string]::IsNullOrWhiteSpace($selectedAppId) -or $selectedAppId -eq '04b07795-8ddb-461a-bbee-02f9e1bf7b46') {
        return $null
    }

    return [ordered]@{
        AppId = $selectedAppId
        Organization = $selectedOrganization
        Thumbprint = $selectedThumbprint
        Path = $selectedPath
        PasswordEnvironmentVariable = $selectedPasswordVariable
    }
}

function Connect-ExchangeOrIppsWithCertificate {
    param(
        [Parameter(Mandatory)][string]$CommandName,
        [Parameter(Mandatory)][hashtable]$Selection
    )

    $command = Get-Command $CommandName -ErrorAction Stop
    $params = @{ ErrorAction = 'Stop'; AppId = $Selection.AppId }
    if ($command.Parameters.ContainsKey('ShowBanner')) { $params['ShowBanner'] = $false }
    if (-not [string]::IsNullOrWhiteSpace([string]$Selection.Organization) -and $command.Parameters.ContainsKey('Organization')) {
        $params['Organization'] = $Selection.Organization
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Selection.Path) -and $command.Parameters.ContainsKey('CertificateFilePath')) {
        $params['CertificateFilePath'] = $Selection.Path
        if ($command.Parameters.ContainsKey('CertificatePassword')) {
            $password = Get-EnvironmentSecureString -VariableName $Selection.PasswordEnvironmentVariable
            if ($null -ne $password) { $params['CertificatePassword'] = $password }
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$Selection.Thumbprint) -and $command.Parameters.ContainsKey('CertificateThumbprint')) {
        $params['CertificateThumbprint'] = Get-NormalizedCertificateThumbprint -Thumbprint $Selection.Thumbprint
    }
    else {
        throw "$CommandName does not expose a supported certificate parameter for the supplied certificate."
    }

    & $CommandName @params -WarningAction SilentlyContinue -InformationAction SilentlyContinue 3>$null 6>$null | Out-Null
}

function Connect-TeamsWithCertificate {
    param([Parameter(Mandatory)][hashtable]$Selection)

    $params = @{
        ErrorAction = 'Stop'
        ApplicationId = $Selection.AppId
        TenantId = $Selection.Organization
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Selection.Path)) {
        $certificate = Get-AuthenticationCertificate -Path $Selection.Path -SecretVariableName $Selection.PasswordEnvironmentVariable
        $params['Certificate'] = $certificate
    }
    else {
        $params['CertificateThumbprint'] = Get-NormalizedCertificateThumbprint -Thumbprint $Selection.Thumbprint
    }

    $null = Connect-MicrosoftTeams @params -WarningAction SilentlyContinue -InformationAction SilentlyContinue 3>$null 6>$null
}

function Connect-GraphForExport {
    $isDefaultMsClient = ($script:GraphClientId -eq '04b07795-8ddb-461a-bbee-02f9e1bf7b46')
    $graphScopes = if ($script:GraphScopes -and $script:GraphScopes.Count -gt 0) {
        @($script:GraphScopes)
    }
    elseif ($script:RequestRequiredGraphScopesUpFront) {
        Get-DefaultGraphScopes
    }
    elseif ($isDefaultMsClient) {
        @('https://graph.microsoft.com/.default')
    }
    else {
        Get-DefaultGraphScopes
    }

    $existingContext = $null
    try { $existingContext = Get-MgContext -ErrorAction SilentlyContinue } catch { $existingContext = $null }
    $contextAuthType = [string](Get-ObjectPropertyValue -InputObject $existingContext -Name 'AuthType' -Default '')
    $contextIsAppOnly = $contextAuthType -eq 'AppOnly'

    if ($existingContext) {
        Write-Step 'Detected an existing Microsoft Graph session; continuing with the current token flow.' -Color Green
    }

    if (Get-GraphCertificateMaterialRequested -and -not $contextIsAppOnly) {
        $null = Connect-GraphAppOnlyCertificate -Tenant $script:TenantId -ClientId $script:GraphClientId
        Write-Step 'Connected to Microsoft Graph with an app-only certificate.' -Color Green
        return
    }

    $existingToken = $null
    if ($script:GraphAccessToken) {
        $existingToken = $script:GraphAccessToken
    }
    elseif ($existingContext -and $existingContext.PSObject.Properties.Name -contains 'AccessToken') {
        $existingToken = [string]$existingContext.AccessToken
    }

    $existingTokenExpiresAt = $null
    if (-not [string]::IsNullOrWhiteSpace($existingToken)) {
        $existingTokenExpiresAt = Get-GraphTokenExpirationUtc -AccessToken $existingToken
    }

    if ($script:RequestRequiredGraphScopesUpFront) {
        $existingToken = $null
        $existingTokenExpiresAt = $null
        $script:GraphAccessToken = $null
        $script:GraphAccessTokenExpiresAtUtc = $null
        Write-Step 'Upfront Graph scope mode is enabled; cached Graph tokens will not be reused.' -Color DarkYellow
    }

    $tokenIsUsable = $existingToken -and $existingTokenExpiresAt -and $existingTokenExpiresAt -gt (Get-Date).ToUniversalTime().AddMinutes(1)
    $shouldReuseExistingSession = $false
    if ($existingContext -and ($contextIsAppOnly -or (Test-NonInteractiveMode) -or ($tokenIsUsable -and -not ($graphScopes -contains 'https://graph.microsoft.com/.default')))) {
        $shouldReuseExistingSession = $true
    }

    if ($shouldReuseExistingSession) {
        if ($tokenIsUsable) {
            $script:GraphAccessToken = $existingToken
            $script:GraphAccessTokenExpiresAtUtc = $existingTokenExpiresAt
            Write-Step 'Reusing the current Microsoft Graph access token from the existing session.' -Color Green
        }
        else {
            Write-Step 'Reusing the existing Microsoft Graph session through Invoke-MgGraphRequest.' -Color Green
        }
        $script:GraphAuthenticationMode = 'ExistingSession'
        $script:MgContext = $existingContext
        if ($contextIsAppOnly -and [string]::IsNullOrWhiteSpace($script:GraphAppDisplayName)) {
            $script:GraphAppDisplayName = [string](Get-ObjectPropertyValue -InputObject $existingContext -Name 'AppName' -Default '')
            $script:GraphAppClientId = [string](Get-ObjectPropertyValue -InputObject $existingContext -Name 'ClientId' -Default '')
        }
        Write-Step 'Microsoft Graph authentication is ready.' -Color Green
        return
    }

    if (Test-NonInteractiveMode) {
        throw 'NonInteractive/Headless mode requires an existing Microsoft Graph session or -GraphCertificateThumbprint/-GraphCertificatePath. Device-code sign-in is disabled.'
    }

    if ($script:RequestRequiredGraphScopesUpFront -and $isDefaultMsClient) {
        Write-Warning 'Upfront Graph scope mode works best with a custom app registration. The default Microsoft public client cannot reliably request this tenant-specific delegated scope set directly, so a custom GraphClientId is recommended.'
        Write-Step 'Rerun for full delegated consent with: -GraphClientId <your-app-id> -RequestRequiredGraphScopesUpFront' -Color DarkYellow
    }

    if ($isDefaultMsClient -and ($graphScopes -notcontains 'https://graph.microsoft.com/.default')) {
        Write-Warning 'The Microsoft first-party client ID cannot request tenant-specific delegated Graph scopes directly. Falling back to https://graph.microsoft.com/.default to avoid AADSTS65002.'
        Write-Step 'For the full exporter delegated scope set, rerun with: -GraphClientId <your-app-id> -RequestRequiredGraphScopesUpFront' -Color DarkYellow
        $graphScopes = @('https://graph.microsoft.com/.default')
    }

    if ($graphScopes -contains 'https://graph.microsoft.com/.default') {
        Write-Step 'Using Graph default scope set for device-code authentication.' -Color DarkYellow
        if ($isDefaultMsClient) {
            Write-Step 'To request all exporter delegated Graph permissions explicitly, rerun with: -GraphClientId <your-app-id> -RequestRequiredGraphScopesUpFront' -Color DarkYellow
        }
    }
    else {
        $consentUrl = Get-GraphAdminConsentUrl -TenantId $script:TenantId -ClientId $script:GraphClientId
        if ($script:RequestRequiredGraphScopesUpFront) {
            Write-Step 'Requesting the exporter''s full delegated Graph scope set up front so consent can be granted before collection starts.' -Color DarkYellow
        }
        else {
            Write-Step 'Requesting the required delegated Graph permissions so the tenant can consent to them if needed.' -Color DarkYellow
            Write-Step 'To force consent up front, rerun with: -GraphClientId <your-app-id> -RequestRequiredGraphScopesUpFront' -Color DarkYellow
        }
        Write-Step "If sign-in fails due to missing consent, grant admin consent for the custom app registration: $consentUrl" -Color DarkYellow
    }

    Write-Step 'Starting device-code sign-in and opening the login page automatically...' -Color DarkYellow
    $null = Get-GraphAccessToken -TenantId $script:TenantId -ClientId $script:GraphClientId -Scopes $graphScopes
    if ([string]::IsNullOrWhiteSpace($script:GraphAuthenticationMode)) { $script:GraphAuthenticationMode = 'DeviceCode' }
    Write-Step 'Connected to Microsoft Graph.' -Color Green
}

function Add-ServiceConnectionFailure {
    param(
        [Parameter(Mandatory)][string]$Section,
        [Parameter(Mandatory)][string]$Message,
        [AllowNull()][object]$ErrorRecord
    )

    $reportVariable = Get-Variable -Name Report -Scope Script -ErrorAction SilentlyContinue
    if ($null -eq $reportVariable -or $null -eq $reportVariable.Value) { return }
    $report = $reportVariable.Value
    if (-not ($report -is [System.Collections.IDictionary]) -or -not $report.Contains('Errors') -or $null -eq $report.Errors) { return }

    $exceptionName = 'ConnectionFailed'
    $detail = $Message
    if ($null -ne $ErrorRecord) {
        $exception = Get-ObjectPropertyValue -InputObject $ErrorRecord -Name 'Exception' -Default $null
        if ($null -ne $exception) {
            $exceptionName = $exception.GetType().FullName
        }
        $detailText = ($ErrorRecord | Out-String -Width 200).Trim()
        if (-not [string]::IsNullOrWhiteSpace($detailText)) { $detail = $detailText }
    }

    $report.Errors.Add([ordered]@{
        Section      = $Section
        Message      = $Message
        Exception    = $exceptionName
        Detail       = $detail
        TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
    }) | Out-Null
}

function Set-ServiceConnectionResult {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSUseShouldProcessForStateChangingFunctions',
        '',
        Justification='Records script-scoped connection status and collector errors. It does not change the tenant.')]
    param(
        [Parameter(Mandatory)][ValidateSet('Exchange', 'Compliance')][string]$Name,
        [Parameter(Mandatory)][ValidateSet('connected', 'unavailable', 'failed')][string]$Status,
        [string]$Message = '',
        [AllowNull()][object]$ErrorRecord
    )

    Set-Variable -Name "${Name}Connected" -Scope Script -Value ($Status -eq 'connected')
    Set-Variable -Name "${Name}ConnectionStatus" -Scope Script -Value $Status
    Set-Variable -Name "${Name}ConnectionError" -Scope Script -Value $(if ($Status -eq 'failed') { $Message } else { '' })
    if ($Status -ne 'failed') { return }

    $section = if ($Name -eq 'Compliance') { 'Purview' } else { 'Exchange' }
    Add-ServiceConnectionFailure -Section $section -Message $Message -ErrorRecord $ErrorRecord
}

# ----------------------------------------------------------------------------
# Connections (read-only scopes)
# ----------------------------------------------------------------------------
function Connect-Services {
    Write-Step 'Connecting to Microsoft Graph and Microsoft 365 services...' -Color Yellow

    # PnP.PowerShell is loaded only in the SharePoint child process. Importing it
    # here registers an assembly resolver that recurses when Exchange Online loads
    # MSAL and identity assemblies (Linux stack overflow, exit 134).
    # ExchangeOnlineManagement must load before the MicrosoftTeams module.
    # Teams pins Microsoft.Identity.Client 4.82; Exchange Online Management 3.10.1
    # needs 4.83.1. Importing Teams first makes Connect-ExchangeOnline and
    # Connect-IPPSSession fail to load Microsoft.Identity.Client.dll.
    try { Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue } catch { Write-Debug "ExchangeOnlineManagement optional import skipped: $($_.Exception.Message)" }

    try {
        Connect-GraphForExport
    }
    catch {
        Write-Warning "Graph sign-in failed: $($_.Exception.Message)"
    }

    # Exchange and Purview connect before any later step can touch PnP.
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
            $exchangeCertificate = Get-WorkloadCertificateSelection -AppId $script:ExchangeAppId -Organization $script:ExchangeOrganization -Thumbprint $script:ExchangeCertificateThumbprint -Path $script:ExchangeCertificatePath -SecretVariableName $script:ExchangeCertificatePasswordEnvironmentVariable -AllowGraphFallback
            if ($exchangeCertificate) {
                Write-Step 'Connecting to Exchange Online with a certificate...' -Color Cyan
                Connect-ExchangeOrIppsWithCertificate -CommandName 'Connect-ExchangeOnline' -Selection $exchangeCertificate
                $exoConnected = $true
                Set-ServiceConnectionResult -Name 'Exchange' -Status 'connected'
                Write-Step 'Connected to Exchange Online with a certificate.' -Color Green
            }
            elseif (Test-NonInteractiveMode) {
                Set-ServiceConnectionResult -Name 'Exchange' -Status 'unavailable'
                Write-Step 'Headless mode will not start an interactive Exchange Online sign-in. Pass -ExchangeAppId and an Exchange certificate thumbprint or path.' -Color DarkYellow
            }
            else {
                Write-Step 'Connecting to Exchange Online for mailbox and compliance signals...' -Color Cyan
                $exoParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
                if (-not [string]::IsNullOrWhiteSpace($script:TenantId)) {
                    $exoParams['Organization'] = $script:TenantId
                }
                Connect-ExchangeOnline @exoParams -WarningAction SilentlyContinue -InformationAction SilentlyContinue 3>$null 6>$null | Out-Null
                $exoConnected = $true
                Set-ServiceConnectionResult -Name 'Exchange' -Status 'connected'
                Write-Step 'Connected to Exchange Online.' -Color Green
            }
        }
        elseif ($exoConnected) {
            Set-ServiceConnectionResult -Name 'Exchange' -Status 'connected'
            Write-Step 'Reusing existing Exchange Online session.' -Color Green
        }
        else {
            Set-ServiceConnectionResult -Name 'Exchange' -Status 'unavailable'
            Write-Step 'ExchangeOnlineManagement cmdlets are not available in this session.' -Color DarkYellow
        }
    }
    catch {
        $exchangeConnectionMessage = [string]$_.Exception.Message
        Set-ServiceConnectionResult -Name 'Exchange' -Status 'failed' -Message $exchangeConnectionMessage -ErrorRecord $_
        Write-Warning "Exchange Online sign-in was not available during startup: $exchangeConnectionMessage"
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
            $ippsCertificate = Get-WorkloadCertificateSelection -AppId $script:IppsAppId -Organization $script:IppsOrganization -Thumbprint $script:IppsCertificateThumbprint -Path $script:IppsCertificatePath -SecretVariableName $script:IppsCertificatePasswordEnvironmentVariable -AllowGraphFallback
            if ($ippsCertificate) {
                Write-Step 'Connecting to Microsoft Purview with a certificate...' -Color Cyan
                Connect-ExchangeOrIppsWithCertificate -CommandName 'Connect-IPPSSession' -Selection $ippsCertificate
                $ippsConnected = $true
                Set-ServiceConnectionResult -Name 'Compliance' -Status 'connected'
                Write-Step 'Connected to Microsoft Purview compliance session with a certificate.' -Color Green
            }
            elseif (Test-NonInteractiveMode) {
                Set-ServiceConnectionResult -Name 'Compliance' -Status 'unavailable'
                Write-Step 'Headless mode will not start an interactive Purview sign-in. Pass -IppsAppId and an IPPS certificate thumbprint or path.' -Color DarkYellow
            }
            else {
                Write-Step 'Connecting to Microsoft Purview compliance session for DLP and retention signals...' -Color Cyan
                $ippsCommand = Get-Command Connect-IPPSSession -ErrorAction SilentlyContinue
                $ippsParams = @{ ErrorAction = 'Stop' }
                if ($ippsCommand -and $ippsCommand.Parameters.ContainsKey('ShowBanner')) {
                    $ippsParams['ShowBanner'] = $false
                }
                if (-not [string]::IsNullOrWhiteSpace($script:TenantId)) {
                    $ippsParams['Organization'] = $script:TenantId
                }
                Connect-IPPSSession @ippsParams -WarningAction SilentlyContinue -InformationAction SilentlyContinue 3>$null 6>$null | Out-Null
                $ippsConnected = $true
                Set-ServiceConnectionResult -Name 'Compliance' -Status 'connected'
                Write-Step 'Connected to Microsoft Purview compliance session.' -Color Green
            }
        }
        elseif ($ippsConnected) {
            Set-ServiceConnectionResult -Name 'Compliance' -Status 'connected'
            Write-Step 'Reusing existing Microsoft Purview compliance session.' -Color Green
        }
        else {
            Set-ServiceConnectionResult -Name 'Compliance' -Status 'unavailable'
            Write-Step 'Compliance cmdlets are not available in this session.' -Color DarkYellow
        }
    }
    catch {
        $purviewConnectionMessage = [string]$_.Exception.Message
        Set-ServiceConnectionResult -Name 'Compliance' -Status 'failed' -Message $purviewConnectionMessage -ErrorRecord $_
        Write-Warning "Compliance (IPPS) sign-in was not available during startup: $purviewConnectionMessage"
    }

    # Import Teams only after Exchange and Purview have connected, so their
    # Microsoft.Identity.Client 4.83.1 assembly is already loaded.
    $teamsImportError = $null
    try { Import-Module MicrosoftTeams -ErrorAction Stop } catch { $teamsImportError = $_.Exception.Message }

    $script:TeamsConnected = $false
    try {
        $teamsModuleAvailable = [bool](Get-Module -ListAvailable -Name MicrosoftTeams -ErrorAction SilentlyContinue)
        $teamsConnectCommand = Get-Command Connect-MicrosoftTeams -ErrorAction SilentlyContinue
        $teamsCmdletAvailable = $null -ne $teamsConnectCommand

        if ($teamsCmdletAvailable) {
            $existingTeamsSession = $false
            if (-not $script:ForceInteractiveTeamsPnPLogin) {
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
                $teamsCertificate = Get-WorkloadCertificateSelection -AppId $script:TeamsAppId -Organization $script:TeamsOrganization -Thumbprint $script:TeamsCertificateThumbprint -Path $script:TeamsCertificatePath -SecretVariableName $script:TeamsCertificatePasswordEnvironmentVariable -AllowGraphFallback
                if ($teamsCertificate) {
                    try {
                        Connect-TeamsWithCertificate -Selection $teamsCertificate
                        $script:TeamsConnected = $true
                        Write-Step 'Connected to Microsoft Teams with a certificate.' -Color Green
                    }
                    catch {
                        $script:TeamsConnected = $false
                        Write-Warning "Teams certificate sign-in was not available during startup: $($_.Exception.Message)"
                    }
                }
                elseif (Test-NonInteractiveMode) {
                    $script:TeamsConnected = $false
                    Write-Step 'Headless mode will not start an interactive Microsoft Teams sign-in. Pass -TeamsAppId and a Teams certificate thumbprint or path.' -Color DarkYellow
                }
                else {
                    if ($script:ForceInteractiveTeamsPnPLogin) {
                        Write-Step 'Force interactive mode enabled: starting a fresh Microsoft Teams sign-in attempt.' -Color DarkYellow
                    }
                    try {
                        $teamsParams = @{ ErrorAction = 'Stop' }
                        if (-not [string]::IsNullOrWhiteSpace($script:TenantId)) {
                            $teamsParams['TenantId'] = $script:TenantId
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

    try { $script:MgContext = Get-MgContext -ErrorAction Stop } catch { $script:MgContext = $null }
}

function Disconnect-Services {
    Write-Step 'Disconnecting services and clearing sessions...' -Color Yellow
    try {
        if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
            $null = Disconnect-MgGraph -ErrorAction SilentlyContinue -InformationAction SilentlyContinue 6>$null
        }
    }
    catch {
        Write-Debug "Disconnect-MgGraph cleanup ignored: $($_.Exception.Message)"
    }
    try {
        if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
            $null = Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue -InformationAction SilentlyContinue 6>$null
        }
    }
    catch {
        Write-Debug "Disconnect-ExchangeOnline cleanup ignored: $($_.Exception.Message)"
    }
    try {
        if (Get-Command Disconnect-IPPSSession -ErrorAction SilentlyContinue) {
            $null = Disconnect-IPPSSession -Confirm:$false -ErrorAction SilentlyContinue -InformationAction SilentlyContinue 6>$null
        }
    }
    catch {
        Write-Debug "Disconnect-IPPSSession cleanup ignored: $($_.Exception.Message)"
    }
    try {
        if (Get-Command Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue) {
            $null = Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue -InformationAction SilentlyContinue 6>$null
        }
    }
    catch {
        Write-Debug "Disconnect-MicrosoftTeams cleanup ignored: $($_.Exception.Message)"
    }
    try {
        # Get-Command would autoload PnP.PowerShell into this process. Only disconnect
        # when the child path already imported it, which the parent never does.
        if ((Get-Module -Name PnP.PowerShell) -and (Get-Command Disconnect-PnPOnline -ErrorAction SilentlyContinue)) {
            $null = Disconnect-PnPOnline -ErrorAction SilentlyContinue -InformationAction SilentlyContinue 6>$null
        }
    }
    catch {
        Write-Debug "Disconnect-PnPOnline cleanup ignored: $($_.Exception.Message)"
    }
    try {
        if (Get-Command Disconnect-SPOService -ErrorAction SilentlyContinue) {
            $null = Disconnect-SPOService -ErrorAction SilentlyContinue -InformationAction SilentlyContinue 6>$null
        }
    }
    catch {
        Write-Debug "Disconnect-SPOService cleanup ignored: $($_.Exception.Message)"
    }
}

# ============================================================================
# COLLECTORS - each returns an [ordered] hashtable; no side effects.
# ============================================================================

function Get-ExecutingAdminLabel {
    $account = [string](Get-ObjectPropertyValue -InputObject $script:MgContext -Name 'Account' -Default '')
    if (-not [string]::IsNullOrWhiteSpace($account)) { return $account }

    $clientId = [string]$script:GraphAppClientId
    if ([string]::IsNullOrWhiteSpace($clientId)) {
        $clientId = [string](Get-ObjectPropertyValue -InputObject $script:MgContext -Name 'ClientId' -Default '')
    }
    if ([string]::IsNullOrWhiteSpace($clientId)) { $clientId = Get-ScriptStringParameter -Name 'GraphClientId' }

    $appName = [string]$script:GraphAppDisplayName
    if ([string]::IsNullOrWhiteSpace($appName)) {
        $appName = [string](Get-ObjectPropertyValue -InputObject $script:MgContext -Name 'AppName' -Default '')
    }
    if ([string]::IsNullOrWhiteSpace($appName) -and -not [string]::IsNullOrWhiteSpace([string]$script:GraphAccessToken)) {
        $claims = Get-GraphTokenClaims -AccessToken $script:GraphAccessToken
        if ($claims) {
            foreach ($claimName in @('app_displayname', 'appDisplayName', 'name')) {
                if ($claims.ContainsKey($claimName) -and -not [string]::IsNullOrWhiteSpace([string]$claims[$claimName])) {
                    $appName = [string]$claims[$claimName]
                    break
                }
            }
        }
    }

    $authType = [string](Get-ObjectPropertyValue -InputObject $script:MgContext -Name 'AuthType' -Default '')
    $isAppOnly = $script:GraphAuthenticationMode -eq 'AppOnlyCertificate' -or $authType -eq 'AppOnly' -or ($script:GraphAuthenticationMode -eq 'ExistingSession' -and [string]::IsNullOrWhiteSpace($account) -and -not [string]::IsNullOrWhiteSpace($clientId) -and $clientId -ne '04b07795-8ddb-461a-bbee-02f9e1bf7b46')
    if (-not $isAppOnly) { return 'not collected' }
    if ([string]::IsNullOrWhiteSpace($appName)) { $appName = 'App' }
    if ([string]::IsNullOrWhiteSpace($clientId)) { $clientId = 'unknown-client' }
    return "$appName ($clientId) [AppOnly]"
}

function Get-TenantMetadata {
    $org = $null
    $users = @()
    $userCount = 'not collected'
    $memberCount = 'not collected'
    $guestCount = 'not collected'
    $defaultDomain = 'not collected'
    $geoLocation = 'not collected'
    $organizationDenied = $false
    $usersDenied = $false

    $missingMetadataScopes = @(Get-GraphMissingScopes -RequiredScopes @('Directory.Read.All', 'Organization.Read.All'))

    try {
        $org = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/organization')) | Select-Object -First 1
        if ($null -eq $org -and $script:GraphLastHttpStatus -eq 403) { $organizationDenied = $true }
        if ($null -ne $org) {
            $defaultDomainValue = @($org.VerifiedDomains | Where-Object { $_.IsDefault -eq $true } | Select-Object -ExpandProperty Name -First 1)
            if (-not [string]::IsNullOrWhiteSpace([string]$defaultDomainValue)) {
                $defaultDomain = [string]$defaultDomainValue
            }

            $geoLocationValue = Get-ObjectPropertyValue -InputObject $org -Name 'CountryLetterCode' -Default 'not collected'
            if (Test-ValueCollected -Value $geoLocationValue) {
                $geoLocation = [string]$geoLocationValue
            }
        }
    }
    catch {
        $org = $null
        if ($script:GraphLastHttpStatus -eq 403) { $organizationDenied = $true }
    }

    try {
        $users = @(Get-DirectoryUsersForReadiness)
        if (@($users).Count -eq 0 -and $script:GraphLastHttpStatus -eq 403) {
            $usersDenied = $true
            $userCount = 'not collected'
            $memberCount = 'not collected'
            $guestCount = 'not collected'
        }
        else {
            $userCount = @($users).Count
            $memberCount = @($users | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'userType' -Default '')) -ne 'Guest' }).Count
            $guestCount = @($users | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'userType' -Default '')) -eq 'Guest' }).Count
        }
    }
    catch {
        $userCount = 'not collected'
        $memberCount = 'not collected'
        $guestCount = 'not collected'
        if ($script:GraphLastHttpStatus -eq 403) { $usersDenied = $true }
    }

    $contextTenantId = Get-ObjectPropertyValue -InputObject $script:MgContext -Name 'TenantId' -Default 'not collected'
    $tenantIdValue = Get-ObjectPropertyValue -InputObject $org -Name 'Id' -Default 'not collected'
    if (-not (Test-ValueCollected -Value $tenantIdValue)) {
        if (-not [string]::IsNullOrWhiteSpace($script:TenantId)) {
            $tenantIdValue = $script:TenantId
        }
        elseif (Test-ValueCollected -Value $contextTenantId) {
            $tenantIdValue = $contextTenantId
        }
    }

    $tenantNameValue = Get-ObjectPropertyValue -InputObject $org -Name 'DisplayName' -Default 'not collected'
    if (-not (Test-ValueCollected -Value $tenantNameValue)) {
        $adminUpn = [string](Get-ObjectPropertyValue -InputObject $script:MgContext -Name 'Account' -Default '')
        if (-not [string]::IsNullOrWhiteSpace($adminUpn) -and $adminUpn.Contains('@')) {
            $tenantNameValue = ($adminUpn.Split('@')[-1]).ToLowerInvariant()
        }
        elseif (Test-ValueCollected -Value $defaultDomain) {
            $tenantNameValue = [string]$defaultDomain
        }
    }

    if (-not (Test-ValueCollected -Value $defaultDomain)) {
        $defaultDomainFromAccount = [string](Get-ObjectPropertyValue -InputObject $script:MgContext -Name 'Account' -Default '')
        if (-not [string]::IsNullOrWhiteSpace($defaultDomainFromAccount) -and $defaultDomainFromAccount.Contains('@')) {
            $defaultDomain = ($defaultDomainFromAccount.Split('@')[-1]).ToLowerInvariant()
        }
    }

    $reason = ''
    $status = 'reported'
    if ($null -ne $org) {
        $missingMetadataScopes = @($missingMetadataScopes | Where-Object { $_ -ne 'organization.read.all' })
    }
    if ($userCount -isnot [string]) {
        $missingMetadataScopes = @($missingMetadataScopes | Where-Object { $_ -notin @('directory.read.all', 'directory.readwrite.all', 'user.read.all') })
    }
    if (($organizationDenied -or $usersDenied) -and @($missingMetadataScopes).Count -gt 0) {
        $status = 'partial'
        $reason = "Missing Graph permission(s): $(@($missingMetadataScopes) -join ', '). Metadata fields were populated from available fallback sources where possible."
    }

    if (-not (Test-ValueCollected -Value $tenantIdValue) -and -not (Test-ValueCollected -Value $tenantNameValue)) {
        $status = 'not collected'
        if ([string]::IsNullOrWhiteSpace($reason)) {
            $reason = 'Tenant metadata was unavailable from both Microsoft Graph and current auth context.'
        }
    }

    return [ordered]@{
        Status             = $status
        Reason             = $reason
        TenantName         = if (Test-ValueCollected -Value $tenantNameValue) { $tenantNameValue } else { 'not collected' }
        TenantId           = if (Test-ValueCollected -Value $tenantIdValue) { $tenantIdValue } else { 'not collected' }
        DefaultDomain      = if (Test-ValueCollected -Value $defaultDomain) { $defaultDomain } else { 'not collected' }
        GeoLocation        = $geoLocation
        ScriptVersion      = '1.3'
        RunTimestampUtc    = (Get-Date).ToUniversalTime().ToString('o')
        ExecutingAdmin     = Get-ExecutingAdminLabel
        GrantedScopes      = if ($script:MgContext -and $script:MgContext.Scopes) { @($script:MgContext.Scopes) } elseif (@($script:GraphProvidedScopes).Count -gt 0) { @($script:GraphProvidedScopes) } else { 'not collected' }
        ConsentStatus      = if (@($script:GraphMissingScopes).Count -gt 0) { 'partial' } else { 'complete' }
        MissingGraphPermissions = @($script:GraphMissingScopes)
        AuthenticationMode = if (-not [string]::IsNullOrWhiteSpace([string]$script:GraphAuthenticationMode)) { [string]$script:GraphAuthenticationMode } else { 'DeviceCode' }
        UnattendedExecutionSupported = ([string]$script:GraphAuthenticationMode -eq 'AppOnlyCertificate')
        AuthenticationModeNote = switch ([string]$(if ($script:GraphAuthenticationMode) { $script:GraphAuthenticationMode } else { 'DeviceCode' })) {
            'AppOnlyCertificate' { 'App-only client-credentials certificate authentication.' }
            'ExistingSession' { 'Reused an existing Microsoft Graph session. App-only sessions are read with Invoke-MgGraphRequest when the access token is not exposed.' }
            default { 'Delegated device-code sign-in.' }
        }
        TotalUsers         = $userCount
        MemberUsers        = $memberCount
        GuestUsers         = $guestCount
        IncludeSampling    = [bool]$script:IncludeSampling
        IncludeDetailedData = [bool]$script:IncludeDetailedData
        SampleSize         = [int]$script:SampleSize
    }
}

function Get-LicensingReadiness {
    $scopeOk = Test-GraphPermissionsAvailable -RequiredPermissions @('Directory.Read.All', 'Organization.Read.All')
    $skus = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus'))
    if (@($skus).Count -eq 0 -and $script:GraphLastHttpStatus -eq 403) {
        return [ordered]@{
            Status = 'not collected'
            Reason = 'Missing Graph permission: Directory.Read.All and/or Organization.Read.All'
            Skus = @()
            CopilotSku = 'not collected'
            QualifyingBaseLicenses = @()
            PrereqReadyUserCount = 'not collected'
            PrereqReadyDenominator = 'not collected'
            PrereqReadyDenominatorDefinition = 'Enabled member users. Guests are excluded. Disabled accounts are excluded when accountEnabled is present.'
            PrereqReadyPercent = 'not collected'
        }
    }
    if (@($skus).Count -eq 0 -and -not $scopeOk -and $null -eq $script:GraphLastHttpStatus) {
        return [ordered]@{
            Status = 'not collected'
            Reason = 'Missing Graph permission: Directory.Read.All and/or Organization.Read.All'
            Skus = @()
            CopilotSku = 'not collected'
            QualifyingBaseLicenses = @()
            PrereqReadyUserCount = 'not collected'
            PrereqReadyDenominator = 'not collected'
            PrereqReadyDenominatorDefinition = 'Enabled member users. Guests are excluded. Disabled accounts are excluded when accountEnabled is present.'
            PrereqReadyPercent = 'not collected'
        }
    }

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

    $eligibleUsers = @(Get-EnabledMemberUsers)

    $ready = 0
    foreach ($u in @($eligibleUsers)) {
        if ($null -eq $u) { continue }
        $plansProp = $u.PSObject.Properties['assignedPlans']
        if ($null -eq $plansProp -or $null -eq $plansProp.Value) { continue }

        # Graph assignedPlans expose capabilityStatus + service (not provisioningStatus/servicePlanName).
        $hasApps = $false
        $hasExchange = $false
        $hasSharePoint = $false
        $hasTeams = $false

        foreach ($plan in $plansProp.Value) {
            if ($null -eq $plan) { continue }
            $capProp = $plan.PSObject.Properties['capabilityStatus']
            if ($null -eq $capProp -or [string]$capProp.Value -ne 'Enabled') { continue }

            $svcProp = $plan.PSObject.Properties['service']
            if ($null -eq $svcProp) { continue }
            $svc = [string]$svcProp.Value
            if ([string]::IsNullOrWhiteSpace($svc)) { continue }

            if (-not $hasApps -and ($svc -eq 'MicrosoftOffice' -or $svc -eq 'OfficeForWeb')) {
                $hasApps = $true
            }
            elseif (-not $hasExchange -and $svc -eq 'exchange') {
                $hasExchange = $true
            }
            elseif (-not $hasSharePoint -and $svc.StartsWith('SharePoint', [System.StringComparison]::OrdinalIgnoreCase)) {
                $hasSharePoint = $true
            }
            elseif (-not $hasTeams -and ($svc.StartsWith('TeamspaceAPI', [System.StringComparison]::OrdinalIgnoreCase) -or $svc.StartsWith('MicrosoftTeams', [System.StringComparison]::OrdinalIgnoreCase))) {
                $hasTeams = $true
            }

            if ($hasApps -and $hasExchange -and $hasSharePoint -and $hasTeams) {
                $ready++
                break
            }
        }
    }

    return [ordered]@{
        Status                = 'reported'
        Reason                = ''
        Skus                  = $skuList
        CopilotSku            = if (@($copilotSkuMatches).Count -gt 0) { @($copilotSkuMatches) } else { 'not collected' }
        QualifyingBaseLicenses = @($qualifying)
        PrereqReadyUserCount  = $ready
        PrereqReadyDenominator = @($eligibleUsers).Count
        PrereqReadyDenominatorDefinition = 'Enabled member users. Guests are excluded. Disabled accounts are excluded when accountEnabled is present.'
        PrereqReadyPercent   = if (@($eligibleUsers).Count -gt 0) { [math]::Round(($ready / @($eligibleUsers).Count) * 100, 2) } else { 0 }
    }
}

function Get-IdentityReadiness {
    $missingCoreScopes = @(Get-GraphMissingScopes -RequiredScopes @('AuditLog.Read.All', 'Directory.Read.All'))
    $hasAuditLogRead = @((Get-GraphMissingScopes -RequiredScopes @('AuditLog.Read.All'))).Count -eq 0
    $hasDirectoryRead = @((Get-GraphMissingScopes -RequiredScopes @('Directory.Read.All'))).Count -eq 0

    $auth = @()
    $featureSummary = $null
    $mfaCapable = 'not collected'
    $mfaRegistered = 'not collected'
    $passwordless = 'not collected'
    $mfaPopulationUserCount = 'not collected'
    $mfaPopulationSource = 'not collected'
    $mfaSummaryTotalUserCount = 'not collected'
    $mfaSummaryUserTypes = 'not collected'
    $mfaSummaryUserRoles = 'not collected'
    $mfaCapableSummaryCount = 'not collected'
    $mfaRegisteredMemberCount = 'not collected'
    $mfaRegisteredGuestCount = 'not collected'
    $mfaRegisteredAdminCount = 'not collected'
    $mfaMemberPopulationUserCount = 'not collected'
    $mfaGuestPopulationUserCount = 'not collected'
    $mfaAdminPopulationUserCount = 'not collected'
    $mfaRegisteredMemberPercent = 'not collected'
    $mfaRegisteredGuestPercent = 'not collected'
    $mfaRegisteredAdminPercent = 'not collected'
    $mfaRegisteredEnabledMemberCount = 'not collected'
    $mfaEnabledMemberPopulationUserCount = 'not collected'
    $mfaRegisteredEnabledMemberPercent = 'not collected'

    $registrationDenied = $false
    try {
        $auth = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails'))
        if (@($auth).Count -eq 0 -and $script:GraphLastHttpStatus -eq 403) { $registrationDenied = $true }
    }
    catch {
        $auth = @()
        if ($script:GraphLastHttpStatus -eq 403) { $registrationDenied = $true }
    }

    if (-not $registrationDenied -and ($hasAuditLogRead -or @($auth).Count -gt 0 -or $script:GraphLastHttpStatus -eq 200)) {

        try {
            $featureSummary = Get-UserRegistrationFeatureSummary -IncludedUserTypes 'all' -IncludedUserRoles 'all'
        }
        catch {
            $featureSummary = $null
        }

        $cohorts = Get-MfaRegistrationCohort -RegistrationDetails $auth
        $mfaCapable = $cohorts.CapableCount
        $mfaRegistered = $cohorts.RegisteredCount
        $passwordless = $cohorts.PasswordlessCount
        $mfaRegisteredMemberCount = if ($null -eq $cohorts.MemberRegistered) { 'not collected' } else { $cohorts.MemberRegistered }
        $mfaRegisteredGuestCount = if ($null -eq $cohorts.GuestRegistered) { 'not collected' } else { $cohorts.GuestRegistered }
        $mfaRegisteredAdminCount = if ($null -eq $cohorts.AdminRegistered) { 'not collected' } else { $cohorts.AdminRegistered }
        $mfaMemberPopulationUserCount = if ($null -eq $cohorts.MemberPopulation) { 'not collected' } else { $cohorts.MemberPopulation }
        $mfaGuestPopulationUserCount = if ($null -eq $cohorts.GuestPopulation) { 'not collected' } else { $cohorts.GuestPopulation }
        $mfaAdminPopulationUserCount = if ($null -eq $cohorts.AdminPopulation) { 'not collected' } else { $cohorts.AdminPopulation }
        $mfaRegisteredMemberPercent = Get-RegistrationPercent -Registered $cohorts.MemberRegistered -Population $cohorts.MemberPopulation
        $mfaRegisteredGuestPercent = Get-RegistrationPercent -Registered $cohorts.GuestRegistered -Population $cohorts.GuestPopulation
        $mfaRegisteredAdminPercent = Get-RegistrationPercent -Registered $cohorts.AdminRegistered -Population $cohorts.AdminPopulation
        if ($hasDirectoryRead) {
            try {
                $enabledMemberMfa = Get-EnabledMemberMfaRegistration -RegistrationDetails $auth
                if ($null -ne $enabledMemberMfa -and (Test-ValueCollected -Value $enabledMemberMfa.Percent)) {
                    $mfaRegisteredEnabledMemberCount = $enabledMemberMfa.Registered
                    $mfaEnabledMemberPopulationUserCount = $enabledMemberMfa.Population
                    $mfaRegisteredEnabledMemberPercent = $enabledMemberMfa.Percent
                }
            }
            catch {
                Write-Debug "Enabled-member MFA percentage was not calculated: $($_.Exception.Message)"
            }
        }

        if (@($auth).Count -gt 0) {
            $mfaPopulationUserCount = @($auth).Count
            $mfaPopulationSource = 'reports/authenticationMethods/userRegistrationDetails row count'
        }
        elseif ($featureSummary -and $featureSummary.PSObject.Properties.Name -contains 'totalUserCount') {
            $mfaPopulationUserCount = Get-IntValue -Value $featureSummary.totalUserCount -Default 0
            $mfaPopulationSource = 'reports/authenticationMethods/usersRegisteredByFeature totalUserCount fallback'
        }

        if ($featureSummary -and $featureSummary.PSObject.Properties.Name -contains 'totalUserCount') {
            $mfaSummaryTotalUserCount = Get-IntValue -Value $featureSummary.totalUserCount -Default 0
        }
        if ($featureSummary -and $featureSummary.PSObject.Properties.Name -contains 'userTypes') {
            $mfaSummaryUserTypes = [string]$featureSummary.userTypes
        }
        if ($featureSummary -and $featureSummary.PSObject.Properties.Name -contains 'userRoles') {
            $mfaSummaryUserRoles = [string]$featureSummary.userRoles
        }
        if ($featureSummary -and $featureSummary.PSObject.Properties.Name -contains 'userRegistrationFeatureCounts') {
            $mfaCapableSummary = Get-UserRegistrationFeatureCount -FeatureCounts $featureSummary.userRegistrationFeatureCounts -FeatureName 'mfaCapable'
            if ($null -ne $mfaCapableSummary) {
                $mfaCapableSummaryCount = $mfaCapableSummary
            }
        }
    }

    $phishingFailureReason = ''
    if ($registrationDenied -or (@($auth).Count -eq 0 -and $script:GraphLastHttpStatus -eq 403)) {
        $phishingFailureReason = Get-ScriptStringParameter -Name 'GraphLastErrorMessage'
        if ([string]::IsNullOrWhiteSpace($phishingFailureReason)) {
            $phishingFailureReason = 'HTTP 403. UserAuthenticationMethod.Read.All or AuditLog.Read.All is required.'
        }
    }
    elseif (-not ($hasAuditLogRead -or @($auth).Count -gt 0 -or $script:GraphLastHttpStatus -eq 200)) {
        $phishingFailureReason = 'Authentication method registration details were not returned.'
    }
    $phishingCoverage = Get-PhishingResistantMfaCoverageResult -RegistrationDetails $(if ($phishingFailureReason) { $null } else { $auth }) -FailureReason $phishingFailureReason

    $globalAdminCount = 'not collected'
    $privilegedRoleCount = 'not collected'
    $guestUsers = 'not collected'
    $staleUsers = 'not collected'

    if ($hasDirectoryRead) {
        $targetRoleNames = @(
            'Global Administrator','Exchange Administrator','SharePoint Administrator',
            'Teams Administrator','Security Administrator','Compliance Administrator',
            'Conditional Access Administrator','Privileged Role Administrator'
        )
        $globalAdminCount = 0
        $privilegedRoleCount = 0

        try {
            $roles = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/directoryRoles'))
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
        }
        catch {
            $globalAdminCount = 'not collected'
            $privilegedRoleCount = 'not collected'
        }

        try {
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
        }
        catch {
            $guestUsers = 'not collected'
            $staleUsers = 'not collected'
        }
    }

    $securityDefaults = $null
    if (Test-GraphPermissionsAvailable -RequiredPermissions @('Policy.Read.All')) {
        try {
            $securityDefaults = Get-GraphRest -Uri 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy'
        }
        catch {
            $securityDefaults = $null
        }
    }

    $reasons = [System.Collections.Generic.List[string]]::new()
    if (@($missingCoreScopes).Count -gt 0) {
        $reasons.Add("Missing Graph permission(s): $(@($missingCoreScopes) -join ', ')") | Out-Null
    }
    if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('Policy.Read.All'))) {
        $reasons.Add('Security defaults could not be collected because Policy.Read.All is missing.') | Out-Null
    }

    $hasAnyCoreData =
        (Test-ValueCollected -Value $mfaCapable) -or
        (Test-ValueCollected -Value $mfaRegistered) -or
        (Test-ValueCollected -Value $passwordless) -or
        (Test-ValueCollected -Value $guestUsers) -or
        (Test-ValueCollected -Value $staleUsers) -or
        (Test-ValueCollected -Value $globalAdminCount) -or
        (Test-ValueCollected -Value $privilegedRoleCount)

    $identityStatus = if ($hasAnyCoreData) {
        if ($reasons.Count -gt 0) { 'partial' } else { 'reported' }
    }
    else {
        'not collected'
    }

    $identityReason = if ($reasons.Count -gt 0) { $reasons -join ' ' } else { '' }

    return [ordered]@{
        Status               = $identityStatus
        Reason               = $identityReason
        MfaCapableCount      = $mfaCapable
        MfaRegisteredCount   = $mfaRegistered
        MfaRegisteredMemberCount = $mfaRegisteredMemberCount
        MfaRegisteredGuestCount = $mfaRegisteredGuestCount
        MfaRegisteredAdminCount = $mfaRegisteredAdminCount
        MfaMemberPopulationUserCount = $mfaMemberPopulationUserCount
        MfaGuestPopulationUserCount = $mfaGuestPopulationUserCount
        MfaAdminPopulationUserCount = $mfaAdminPopulationUserCount
        MfaRegisteredMemberPercent = $mfaRegisteredMemberPercent
        MfaRegisteredGuestPercent = $mfaRegisteredGuestPercent
        MfaRegisteredAdminPercent = $mfaRegisteredAdminPercent
        MfaRegisteredEnabledMemberCount = $mfaRegisteredEnabledMemberCount
        MfaEnabledMemberPopulationUserCount = $mfaEnabledMemberPopulationUserCount
        MfaRegisteredEnabledMemberPercent = $mfaRegisteredEnabledMemberPercent
        MfaRegistrationPercentNote = 'MfaRegisteredCount and the overall percent include guests and are informational. Scoring uses MfaRegisteredEnabledMemberPercent: enabled members only. Guests and disabled accounts are excluded.'
        PhishingResistantMfaCoverage = $phishingCoverage.Coverage
        PhishingResistantMfaUserCount = $phishingCoverage.UserCount
        PhishingResistantMfaDenominator = $phishingCoverage.Denominator
        PhishingResistantMfaReason = $phishingCoverage.Reason
        PasswordlessCount    = $passwordless
        MfaPopulationUserCount = $mfaPopulationUserCount
        MfaPopulationSource  = $mfaPopulationSource
        MfaSummaryTotalUserCount = $mfaSummaryTotalUserCount
        MfaSummaryUserTypes  = $mfaSummaryUserTypes
        MfaSummaryUserRoles  = $mfaSummaryUserRoles
        MfaCapableSummaryCount = $mfaCapableSummaryCount
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
            Status = 'partial'
            Reason = 'Missing Graph permission: Policy.Read.All. Conditional Access policy inventory and coverage signals were not collected.'
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

function Get-CompliancePolicyWorkloads {
    param([AllowNull()][object]$Policy)

    $workload = Get-ObjectPropertyValue -InputObject $Policy -Name 'Workload' -Default $null
    if ($null -eq $workload) { $workload = Get-ObjectPropertyValue -InputObject $Policy -Name 'Workloads' -Default $null }
    if ($null -eq $workload) { return 'not collected' }

    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($item in @($workload)) {
        $text = [string]$item
        if (-not [string]::IsNullOrWhiteSpace($text)) { $names.Add($text.Trim()) }
    }
    if ($names.Count -eq 0) { return 'not collected' }
    return ($names -join ', ')
}

function ConvertTo-CompliancePolicyList {
    # Always returns a list, including when the cmdlet emits one policy or none.
    param([AllowNull()][object[]]$Policies)

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($policy in @($Policies)) {
        if ($null -eq $policy) { continue }
        $mode = Get-ObjectPropertyValue -InputObject $policy -Name 'Mode' -Default $null
        $state = Get-ObjectPropertyValue -InputObject $policy -Name 'State' -Default $null
        if ($null -eq $state -or [string]::IsNullOrWhiteSpace([string]$state)) {
            if ($null -ne $mode -and -not [string]::IsNullOrWhiteSpace([string]$mode)) {
                $state = [string]$mode
            }
        }

        $enabled = Get-ObjectPropertyValue -InputObject $policy -Name 'Enabled' -Default $null
        if ($null -eq $enabled -or ($enabled -is [string] -and [string]::IsNullOrWhiteSpace([string]$enabled))) {
            $modeText = [string]$mode
            if ($modeText -match '^(Enable|Enabled|Enforce)') { $enabled = $true }
            elseif (-not [string]::IsNullOrWhiteSpace($modeText)) { $enabled = $false }
            else { $enabled = 'not collected' }
        }

        $list.Add([ordered]@{
            Name      = [string](Get-ObjectPropertyValue -InputObject $policy -Name 'Name' -Default '')
            Enabled   = $enabled
            Mode      = if ($null -ne $mode -and -not [string]::IsNullOrWhiteSpace([string]$mode)) { [string]$mode } else { 'not collected' }
            State     = if ($null -ne $state -and -not [string]::IsNullOrWhiteSpace([string]$state)) { [string]$state } else { 'not collected' }
            Workloads = Get-CompliancePolicyWorkloads -Policy $policy
        }) | Out-Null
    }

    # The unary comma keeps a 0- or 1-item result as a list instead of unwrapping it.
    return ,$list.ToArray()
}

function Get-MissingComplianceDataReason {
    # Shared by the not-collected and partial Data governance paths so a failed
    # Purview connect is not dropped when Graph labels are also unavailable.
    param([bool]$HasIppsSession)

    $purviewConnectionError = Get-ScriptStringParameter -Name 'ComplianceConnectionError'
    $purviewConnectionStatus = Get-ScriptStringParameter -Name 'ComplianceConnectionStatus'
    if (-not $HasIppsSession -and $purviewConnectionStatus -eq 'failed' -and -not [string]::IsNullOrWhiteSpace($purviewConnectionError)) {
        return "Purview connection failed: $purviewConnectionError"
    }
    if (-not $HasIppsSession) {
        return 'DLP and retention were not collected because no Purview compliance session was available.'
    }
    return 'DLP or retention policy data was not collected.'
}

function Get-ExportExitCode {
    param(
        [int]$FailedCollectorCount,
        [int]$ProcessedSectionCount,
        [bool]$FailOnCollectorError
    )

    if ($FailedCollectorCount -gt 0 -and $ProcessedSectionCount -eq 0) { return 2 }
    if ($FailOnCollectorError -and $FailedCollectorCount -gt 0) { return 3 }
    return 0
}

function Get-PurviewLabelCommandResult {
    # $null means the cmdlet is unavailable or failed. An empty array means the
    # cmdlet succeeded and returned nothing. Callers must not treat $null as zero.
    param([Parameter(Mandatory)][string]$CommandName)

    if (-not (Get-Command $CommandName -ErrorAction SilentlyContinue)) { return $null }
    try {
        $items = @(& $CommandName -ErrorAction Stop | Where-Object { $null -ne $_ })
        return ,$items
    }
    catch {
        $message = "$CommandName failed: $($_.Exception.Message)"
        Register-GraphReportDiagnostic -Uri $CommandName -Message $message -Section 'DataGovernance'
        return $null
    }
}

function Get-DataGovernanceReadiness {
    $labels = @()
    $graphCollected = $false
    $graphFailureReason = ''
    try {
        $labels = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/informationProtection/policy/labels'))
        if ($script:GraphLastHttpStatus -eq 403) {
            $graphFailureReason = 'Graph sensitivity label read returned HTTP 403. InformationProtectionPolicy.Read.All is required.'
            $labels = @()
        }
        elseif ($script:GraphLastHttpStatus -eq 200 -or @($labels).Count -gt 0) {
            $graphCollected = $true
        }
        elseif ([string]::IsNullOrWhiteSpace($graphFailureReason)) {
            $graphFailureReason = 'Graph sensitivity label read did not return a successful response.'
        }
    }
    catch {
        $graphFailureReason = $_.Exception.Message
        $labels = @()
    }

    $hasIppsSession = $false
    if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        $connections = @((Get-ConnectionInformation -ErrorAction SilentlyContinue))
        $hasIppsSession = @($connections | Where-Object {
            ($_.PSObject.Properties.Name -contains 'Name' -and [string]$_.Name -match 'IPPSSession|SecurityCompliance') -or
            ($_.PSObject.Properties.Name -contains 'ConnectionUri' -and [string]$_.ConnectionUri -match 'ps\.compliance\.protection\.outlook\.com')
        }).Count -gt 0
    }

    $complianceLabels = $null
    $compliancePolicies = $null
    $autoPolicies = $null
    $labelPolicyNote = ''
    if (-not $graphCollected -and $hasIppsSession) {
        $complianceLabels = Get-PurviewLabelCommandResult -CommandName 'Get-Label'
        $compliancePolicies = Get-PurviewLabelCommandResult -CommandName 'Get-LabelPolicy'
        $labelCount = if ($null -eq $complianceLabels) { 0 } else { @($complianceLabels).Count }
        $labelPolicyFailureNote = 'Get-LabelPolicy failed; label policy counts are not collected.'
        # A live run returned labels and then an empty policy list. The next run
        # returned the policies. Do not score that first empty list as zero.
        # A thrown read is $null. With labels present, that is partial, not a clean report.
        if ($labelCount -gt 0 -and $null -eq $compliancePolicies) {
            $labelPolicyNote = $labelPolicyFailureNote
        }
        elseif ($labelCount -gt 0 -and @($compliancePolicies).Count -eq 0) {
            $retryDelaySeconds = 2
            $retryDelayVariable = Get-Variable -Name LabelPolicyRetryDelaySeconds -Scope Script -ErrorAction SilentlyContinue
            if ($null -ne $retryDelayVariable -and $null -ne $retryDelayVariable.Value) {
                $retryDelaySeconds = [int]$retryDelayVariable.Value
            }
            if ($retryDelaySeconds -gt 0) { Start-Sleep -Seconds $retryDelaySeconds }
            $retriedPolicies = Get-PurviewLabelCommandResult -CommandName 'Get-LabelPolicy'
            if ($null -eq $retriedPolicies) {
                $compliancePolicies = $null
                $labelPolicyNote = $labelPolicyFailureNote
            }
            elseif (@($retriedPolicies).Count -eq 0) {
                $compliancePolicies = $null
                $labelPolicyNote = 'Get-Label returned labels, and Get-LabelPolicy still returned no policies after one retry. Label policy counts are not collected.'
                Register-GraphReportDiagnostic -Uri 'Get-LabelPolicy' -Message $labelPolicyNote -Section 'DataGovernance' -DiagnosticOnly
            }
            else {
                $compliancePolicies = $retriedPolicies
            }
        }
        $autoPolicies = Get-PurviewLabelCommandResult -CommandName 'Get-AutoSensitivityLabelPolicy'
    }

    $labelSnapshot = Get-SensitivityLabelSnapshot -GraphCollected $graphCollected -GraphLabels $labels -ComplianceSessionConnected $hasIppsSession -ComplianceLabels $complianceLabels -ComplianceLabelPolicies $compliancePolicies -AutoSensitivityLabelPolicies $autoPolicies -GraphFailureReason $graphFailureReason
    if (-not [string]::IsNullOrWhiteSpace($labelPolicyNote)) {
        $labelSnapshot.Reason = $labelPolicyNote
    }
    $published = $labelSnapshot.Published
    $labelTotal = $labelSnapshot.Total
    $autoLabel = $labelSnapshot.AutoLabel

    $dlp = 'not collected'
    $dlpFailureMessage = ''
    if ($hasIppsSession -and (Get-Command Get-DlpCompliancePolicy -ErrorAction SilentlyContinue)) {
        try {
            $dlp = ConvertTo-CompliancePolicyList -Policies @(Get-DlpCompliancePolicy -ErrorAction Stop)
        }
        catch {
            $dlp = 'not collected'
            $dlpFailureMessage = [string]$_.Exception.Message
            Register-GraphReportDiagnostic -Uri 'Get-DlpCompliancePolicy' -Message $dlpFailureMessage -Section 'DataGovernance'
        }
    }
    $retention = 'not collected'
    $retentionFailureMessage = ''
    if ($hasIppsSession -and (Get-Command Get-RetentionCompliancePolicy -ErrorAction SilentlyContinue)) {
        try {
            $retention = ConvertTo-CompliancePolicyList -Policies @(Get-RetentionCompliancePolicy -ErrorAction Stop)
        }
        catch {
            $retention = 'not collected'
            $retentionFailureMessage = [string]$_.Exception.Message
            Register-GraphReportDiagnostic -Uri 'Get-RetentionCompliancePolicy' -Message $retentionFailureMessage -Section 'DataGovernance'
        }
    }
    $dlpCollected = Test-ValueCollected -Value $dlp
    $retentionCollected = Test-ValueCollected -Value $retention
    $hasComplianceData = $dlpCollected -or $retentionCollected

    $labelsCollected = [bool]$labelSnapshot.Collected
    $status = 'reported'
    $reason = ''
    $complianceGapReason = if ($hasComplianceData) { '' } else { Get-MissingComplianceDataReason -HasIppsSession $hasIppsSession }
    if (-not $labelsCollected -and -not $hasComplianceData) {
        $status = 'not collected'
        $reasonParts = [System.Collections.Generic.List[string]]::new()
        if (-not [string]::IsNullOrWhiteSpace([string]$labelSnapshot.Reason)) { $reasonParts.Add([string]$labelSnapshot.Reason) }
        else { $reasonParts.Add('Neither labels nor compliance policy data could be collected.') }
        if (-not [string]::IsNullOrWhiteSpace($complianceGapReason)) { $reasonParts.Add($complianceGapReason) }
        $reason = $reasonParts -join ' '
    }
    elseif (-not $labelsCollected -or -not $hasComplianceData) {
        $status = 'partial'
        $reasonParts = [System.Collections.Generic.List[string]]::new()
        if (-not $labelsCollected -and -not [string]::IsNullOrWhiteSpace([string]$labelSnapshot.Reason)) { $reasonParts.Add([string]$labelSnapshot.Reason) }
        if (-not $hasComplianceData -and -not [string]::IsNullOrWhiteSpace($complianceGapReason)) { $reasonParts.Add($complianceGapReason) }
        $reason = $reasonParts -join ' '
    }
    if (-not [string]::IsNullOrWhiteSpace($labelPolicyNote) -and $reason -notlike "*$labelPolicyNote*") {
        if ($status -eq 'reported') { $status = 'partial' }
        $reason = ("$reason $labelPolicyNote").Trim()
    }
    # One successful compliance read must not hide the other read's failure.
    $complianceReadFailures = [System.Collections.Generic.List[string]]::new()
    if (-not $dlpCollected -and -not [string]::IsNullOrWhiteSpace($dlpFailureMessage)) {
        $complianceReadFailures.Add("DLP policy read failed: $dlpFailureMessage")
    }
    if (-not $retentionCollected -and -not [string]::IsNullOrWhiteSpace($retentionFailureMessage)) {
        $complianceReadFailures.Add("Retention policy read failed: $retentionFailureMessage")
    }
    if ($complianceReadFailures.Count -gt 0) {
        if ($status -eq 'reported') { $status = 'partial' }
        $reason = ("$reason $($complianceReadFailures -join ' ')").Trim()
    }

    return [ordered]@{
        Status                    = $status
        Reason                    = $reason
        SensitivityLabelsPublished = if ($labelsCollected) { $published } else { $null }
        SensitivityLabelsTotal     = if ($labelsCollected) { $labelTotal } else { $null }
        LabelPoliciesPublished     = if ($null -ne $labelSnapshot.PolicyPublished) { $labelSnapshot.PolicyPublished } else { $null }
        LabelPoliciesTotal         = if ($null -ne $labelSnapshot.PolicyTotal) { $labelSnapshot.PolicyTotal } else { $null }
        AutoLabelingIndicators     = if ($labelsCollected -and $null -ne $autoLabel) { $autoLabel } else { $null }
        LabelCollectionSource      = $labelSnapshot.Source
        LabelCollectionReason      = if ([string]::IsNullOrWhiteSpace([string]$labelSnapshot.Reason)) { 'reported' } else { $labelSnapshot.Reason }
        DlpPolicies                = $dlp
        RetentionPolicies          = $retention
    }
}

function Get-SharePointSettingsFromGraph {
    $settings = $null
    try {
        $settings = Get-GraphRest -Uri 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings'
    }
    catch {
        $settings = $null
    }

    if ($null -eq $settings) {
        return [ordered]@{
            Collected = $false
            StatusCode = $script:GraphLastHttpStatus
            Reason = if ($script:GraphLastHttpStatus -eq 403) {
                'Graph SharePoint tenant settings returned HTTP 403. SharePointTenantSettings.Read.All is required for the Graph fallback.'
            }
            else {
                'Graph SharePoint tenant settings were not returned.'
            }
        }
    }

    $sharing = ConvertTo-SharingCapabilityName -Value (Get-ObjectPropertyValue -InputObject $settings -Name 'sharingCapability' -Default $null)
    return [ordered]@{
        Collected = $true
        StatusCode = 200
        SharingCapability = if (Test-ValueCollected -Value $sharing) { [string]$sharing } else { 'not collected' }
        Reason = 'Tenant sharing capability was read from GET /admin/sharepoint/settings. Site inventory was not collected because PnP was unavailable.'
    }
}

function Resolve-SharePointGraphSettingsFallback {
    param([AllowNull()][object]$Section)

    $sharingCollected = $Section -and (Test-ValueCollected -Value (Get-ObjectPropertyValue -InputObject $Section -Name 'ExternalSharingLevel' -Default 'not collected'))
    if ($sharingCollected -and [string](Get-ObjectPropertyValue -InputObject $Section -Name 'Status' -Default '') -ne 'not collected') {
        return $Section
    }

    $graphSettings = Get-SharePointSettingsFromGraph
    if (-not $graphSettings.Collected) {
        if ($Section -and $graphSettings.StatusCode -eq 403) {
            $existingReason = [string](Get-ObjectPropertyValue -InputObject $Section -Name 'Reason' -Default '')
            $Section.Reason = ("$existingReason $($graphSettings.Reason)").Trim()
        }
        return $Section
    }

    $pnpReason = [string](Get-ObjectPropertyValue -InputObject $Section -Name 'Reason' -Default '')
    return [ordered]@{
        Status                = 'partial'
        Reason                = ("$($graphSettings.Reason) $pnpReason").Trim()
        ExternalSharingLevel  = $graphSettings.SharingCapability
        DefaultSharingLinkType = 'not collected'
        RestrictedSearchEnabled = 'not collected'
        TotalSiteCount        = 'not collected'
        Truncated             = 'not collected'
        Sites                 = @()
        OneDriveCoveragePercent = 'not collected'
        InactiveSiteCount     = 'not collected'
        SharePointSettingsSource = 'graph/admin/sharepoint/settings'
        SiteInventoryStatus   = 'not collected'
        SiteInventoryReason   = 'Site inventory was not collected because PnP site collection did not run in this process.'
        OversharingSignals    = [ordered]@{
            Status                    = 'not collected'
            CollectionMethod          = 'Graph admin/sharepoint/settings tenant sharing capability only'
            CollectionScope           = 'tenant'
            CollectionPathUsed        = 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings'
            CollectionReason          = $graphSettings.Reason
            RequiredForReliableCollection = 'PnP site inventory is still required for per-site sharing and oversharing counters'
            DocumentationNote         = 'GET /admin/sharepoint/settings requires SharePointTenantSettings.Read.All and does not return per-site sharing links.'
            SampledAnyoneLinkCount    = 'not collected'
            SampledOrgWideLinkCount   = 'not collected'
            SitesWithAnyoneLinksCount = 'not collected'
        }
    }
}

function Get-SharePointOneDriveReadiness {
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
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

    $childArgs = [System.Collections.Generic.List[string]]::new()
    foreach ($arg in @(
        '-NoProfile',
        '-NoLogo',
        '-ExecutionPolicy', 'Bypass',
        '-File', $PSCommandPath,
        '-SharePointCollectionChild',
        '-SharePointAdminUrl', $adminUrl
    )) {
        $childArgs.Add($arg)
    }
    if (-not [string]::IsNullOrWhiteSpace($script:TenantId)) { $childArgs.Add('-TenantId'); $childArgs.Add($script:TenantId) }
    if (-not [string]::IsNullOrWhiteSpace($script:PnPClientId)) { $childArgs.Add('-PnPClientId'); $childArgs.Add($script:PnPClientId) }
    if (-not [string]::IsNullOrWhiteSpace($script:GraphClientId)) { $childArgs.Add('-GraphClientId'); $childArgs.Add($script:GraphClientId) }
    if (Test-NonInteractiveMode) { $childArgs.Add('-NonInteractive') }
    if (-not [string]::IsNullOrWhiteSpace($script:PnPCertificateThumbprint)) { $childArgs.Add('-PnPCertificateThumbprint'); $childArgs.Add($script:PnPCertificateThumbprint) }
    if (-not [string]::IsNullOrWhiteSpace($script:PnPCertificatePath)) { $childArgs.Add('-PnPCertificatePath'); $childArgs.Add($script:PnPCertificatePath) }
    if (-not [string]::IsNullOrWhiteSpace($script:PnPCertificatePasswordEnvironmentVariable)) { $childArgs.Add('-PnPCertificatePasswordEnvironmentVariable'); $childArgs.Add($script:PnPCertificatePasswordEnvironmentVariable) }
    if ($Sample) { $childArgs.Add('-IncludeSampling') }
    if ($Max -ne 200) { $childArgs.Add('-SampleSize'); $childArgs.Add([string]$Max) }

    $childArgArray = $childArgs.ToArray()
    $childOutput = & $pwshPath @childArgArray 2>&1 | Out-String
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
    $tenantCollectionError = $null
    $siteCollectionError = $null

    if (Get-Command Get-PnPTenant -ErrorAction SilentlyContinue) {
        try {
            $tenant = Get-PnPTenant -ErrorAction Stop
        }
        catch {
            $tenantCollectionError = "Tenant settings could not be collected: $($_.Exception.Message)"
        }
    }
    else {
        $tenantCollectionError = 'Tenant settings could not be collected because Get-PnPTenant is unavailable.'
    }
    if ($null -eq $tenant -and [string]::IsNullOrWhiteSpace($tenantCollectionError)) {
        $tenantCollectionError = 'Tenant settings could not be collected because Get-PnPTenant returned no data.'
    }

    if (Get-Command Get-PnPTenantSite -ErrorAction SilentlyContinue) {
        try {
            $sites = @(Get-PnPTenantSite -IncludeOneDriveSites -ErrorAction Stop)
        }
        catch {
            $siteCollectionError = "Site inventory could not be collected: $($_.Exception.Message)"
        }
    }
    else {
        $siteCollectionError = 'Site inventory could not be collected because Get-PnPTenantSite is unavailable.'
    }

    $siteInventoryCollected = $null -eq $siteCollectionError
    if ($null -eq $tenant -and -not $siteInventoryCollected) {
        $collectionErrors = @(@($tenantCollectionError, $siteCollectionError) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        return [ordered]@{
            Status                = 'not collected'
            Reason                = $collectionErrors -join ' '
            SiteInventoryStatus   = 'not collected'
            SiteInventoryReason   = $siteCollectionError
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

    $totalCount = if ($siteInventoryCollected) { @($sites).Count } else { 'not collected' }
    $siteList = if ($siteInventoryCollected) { @($sites) } else { @() }

    if ($Sample -and $siteInventoryCollected -and $totalCount -gt $Max) {
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
            AnyoneLinkEnabled   = (ConvertTo-SharingCapabilityName -Value $site.SharingCapability) -eq 'ExternalUserAndGuestSharing'
            LastContentModified = $site.LastContentModifiedDate
        }
    }

    $externalSharingLevel = 'not collected'
    if ($tenant) {
        $rawSharingCapability = Get-ObjectPropertyValue -InputObject $tenant -Name 'SharingCapability' -Default 'not collected'
        $sharingCapabilityName = ConvertTo-SharingCapabilityName -Value $rawSharingCapability
        if (-not [string]::IsNullOrWhiteSpace([string]$sharingCapabilityName)) {
            $externalSharingLevel = [string]$sharingCapabilityName
        }
    }

    $collectionErrors = @(@($tenantCollectionError, $siteCollectionError) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    return [ordered]@{
        Status                = if ($collectionErrors.Count -gt 0) { 'partial' } else { 'reported' }
        Reason                = $collectionErrors -join ' '
        SiteInventoryStatus   = if ($siteInventoryCollected) { 'reported' } else { 'not collected' }
        SiteInventoryReason   = if ($siteInventoryCollected) { '' } else { $siteCollectionError }
        ExternalSharingLevel   = $externalSharingLevel
        DefaultSharingLinkType = if ($tenant) { [string](Get-ObjectPropertyValue -InputObject $tenant -Name 'DefaultSharingLinkType' -Default 'not collected') } else { 'not collected' }
        RestrictedSearchEnabled = if ($tenant -and $null -ne (Get-ObjectPropertyValue -InputObject $tenant -Name 'RestrictedSharePointSearch')) { [bool](Get-ObjectPropertyValue -InputObject $tenant -Name 'RestrictedSharePointSearch') } else { 'not collected' }
        TotalSiteCount         = $totalCount
        Truncated              = if ($siteInventoryCollected) { [bool]($Sample -and $totalCount -gt $Max) } else { 'not collected' }
        Sites                  = @($siteSummaries)
        OneDrivePersonalSiteCount = if ($siteInventoryCollected) { $oneDriveCoverage } else { 'not collected' }
        OneDrivePersonalSiteSharePercent = if (-not $siteInventoryCollected) { 'not collected' } elseif (@($siteList).Count -gt 0) { [math]::Round(($oneDriveCoverage / @($siteList).Count) * 100, 2) } else { 0 }
        OneDriveCoverageMetric = 'Deprecated alias: OneDriveCoveragePercent reflects the share of sampled SharePoint/PnP tenant sites whose URL is a OneDrive personal site, not licensed-user OneDrive provisioning coverage.'
        OneDriveCoveragePercent = if (-not $siteInventoryCollected) { 'not collected' } elseif (@($siteList).Count -gt 0) { [math]::Round(($oneDriveCoverage / @($siteList).Count) * 100, 2) } else { 0 }
        InactiveSiteCount      = if ($siteInventoryCollected) { $inactive } else { 'not collected' }
        OversharingSignals     = [ordered]@{
            Status                     = if (-not $siteInventoryCollected) { 'not collected' } elseif ($oversharingSignalAvailable) { 'reported' } else { 'not collected' }
            CollectionMethod           = if (-not $siteInventoryCollected) {
                'Site inventory could not be collected.'
            }
            elseif ($oversharingSignalAvailable) {
                'Best-effort from PnP tenant site metadata counters exposed by current cmdlet/runtime.'
            }
            else {
                'No oversharing counters collected from standard Get-PnPTenantSite output.'
            }
            CollectionScope            = if (-not $siteInventoryCollected) { 'not collected' } elseif ($oversharingSignalAvailable) { 'sampled site metadata counters only (not guaranteed tenant-wide)' } else { 'standard PnP tenant site metadata only' }
            CollectionPathUsed         = if (-not $siteInventoryCollected) { 'not collected' } elseif ($oversharingSignalAvailable) { $oversharingPropertySource } else { 'none' }
            CollectionReason           = if (-not $siteInventoryCollected) {
                $siteCollectionError
            }
            elseif ($oversharingSignalAvailable) {
                'reported from tenant site metadata counters exposed by current cmdlet/runtime'
            }
            else {
                'standard Get-PnPTenantSite output does not normally include per-site sharing-link counters'
            }
            RequiredForReliableCollection = if (-not $siteInventoryCollected) {
                'A successful Get-PnPTenantSite collection is required.'
            }
            elseif ($oversharingSignalAvailable) {
                'SharePoint Advanced Management oversharing telemetry is still recommended for reliable tenant-wide results'
            }
            else {
                'SharePoint Advanced Management oversharing telemetry (recommended), or costly per-site sharing-link enumeration'
            }
            DocumentationNote          = if (-not $siteInventoryCollected) {
                'Unknown site inventory is not evidence of low oversharing risk.'
            }
            elseif ($oversharingSignalAvailable) {
                'Counters were exposed in this run, but this is not guaranteed in standard PnP output across tenants.'
            }
            else {
                'This run did not collect oversharing link counters; absence here does not imply low oversharing risk.'
            }
            SampledAnyoneLinkCount     = if ($siteInventoryCollected -and $oversharingSignalAvailable) { $anyoneLinkCount } else { 'not collected' }
            SampledOrgWideLinkCount    = if ($siteInventoryCollected -and $oversharingSignalAvailable) { $orgWideLinkCount } else { 'not collected' }
            SitesWithAnyoneLinksCount  = if ($siteInventoryCollected -and $oversharingSignalAvailable) { $sitesWithAnyoneLinks } else { 'not collected' }
        }
    }
}

function Get-ExchangeReadiness {
    try {
        $mailboxes = @()
        $hasExoSession = $false
        $exoCollectionError = $null

        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $hasExoSession = @((Get-ConnectionInformation -ErrorAction SilentlyContinue)).Count -gt 0
        }

        if ($hasExoSession -and (Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue)) {
            # ArchiveStatus and LitigationHoldEnabled are not in the default EXO minimum property set.
            try {
                $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails, ArchiveStatus, LitigationHoldEnabled -ErrorAction Stop)
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
            catch {
                $exoCollectionError = "Exchange Online mailbox collection failed: $($_.Exception.Message)"
            }
        }
        elseif ($hasExoSession) {
            $exoCollectionError = 'Exchange Online mailbox collection failed because Get-EXOMailbox is unavailable.'
        }
        else {
            $exchangeConnectionError = Get-ScriptStringParameter -Name 'ExchangeConnectionError'
            $exchangeConnectionStatus = Get-ScriptStringParameter -Name 'ExchangeConnectionStatus'
            if ($exchangeConnectionStatus -eq 'failed' -and -not [string]::IsNullOrWhiteSpace($exchangeConnectionError)) {
                $exoCollectionError = "Exchange Online connection failed: $exchangeConnectionError"
            }
            else {
                $exoCollectionError = 'Exchange Online session was unavailable.'
            }
        }

        if (-not (Test-GraphPermissionsAvailable -RequiredPermissions @('Directory.Read.All'))) {
            return [ordered]@{
                Status               = 'not collected'
                Reason               = "$exoCollectionError Graph directory read access is also missing, so mailbox data could not be collected."
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
            Reason               = "$exoCollectionError Mailbox count was estimated from Graph users because Exchange Online mailbox data was unavailable."
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
            Reason               = "Exchange mailbox data could not be collected: $($_.Exception.Message)"
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
    $groupScopeOk = Test-GraphPermissionsAvailable -RequiredPermissions @('Group.Read.All')
    $groupDenied = $false
    $teams = @()
    try {
        $teams = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/groups' -Query @{ '$filter' = "resourceProvisioningOptions/any(x:x eq 'Team')"; '$select' = 'id,displayName,visibility,renewedDateTime' }))
        if (@($teams).Count -eq 0 -and $script:GraphLastHttpStatus -eq 403) { $groupDenied = $true }
    }
    catch {
        $teams = @()
        if ($script:GraphLastHttpStatus -eq 403) { $groupDenied = $true }
    }

    if ($groupDenied -and -not $groupScopeOk -and -not $script:TeamsConnected) {
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

    $teamCount = if ($groupDenied) { 'not collected' } else { @($teams).Count }
    $privateTeams = if ($groupDenied) { 'not collected' } else { @($teams | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'visibility' -Default '')) -eq 'Private' }).Count }
    $publicTeams = if ($groupDenied) { 'not collected' } else { @($teams | Where-Object { ([string](Get-ObjectPropertyValue -InputObject $_ -Name 'visibility' -Default '')) -eq 'Public' }).Count }
    $staleTeams = if ($groupDenied) {
        'not collected'
    }
    else {
        @($teams | Where-Object {
            $renewed = Get-ObjectPropertyValue -InputObject $_ -Name 'renewedDateTime'
            $renewed -and ((Get-Date).ToUniversalTime() - [DateTime]$renewed) -gt [TimeSpan]::FromDays(180)
        }).Count
    }

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

    $recordingPropertyName = $null
    foreach ($candidateName in @('AllowCloudRecording', 'AllowRecording')) {
        if ($null -ne $meetingPolicy -and $meetingPolicy.PSObject.Properties.Name -contains $candidateName) {
            $recordingPropertyName = $candidateName
            break
        }
    }
    $transcriptionPropertyName = $null
    foreach ($candidateName in @('AllowTranscription', 'AllowCloudTranscription')) {
        if ($null -ne $meetingPolicy -and $meetingPolicy.PSObject.Properties.Name -contains $candidateName) {
            $transcriptionPropertyName = $candidateName
            break
        }
    }
    $hasAllowRecording = -not [string]::IsNullOrWhiteSpace($recordingPropertyName)
    $hasAllowTranscription = -not [string]::IsNullOrWhiteSpace($transcriptionPropertyName)

    $recordingEnabled = if ($hasAllowRecording) { [bool]$meetingPolicy.$recordingPropertyName } else { 'not collected' }
    $transcriptionEnabled = if ($hasAllowTranscription) { [bool]$meetingPolicy.$transcriptionPropertyName } else { 'not collected' }
    $meetingPolicyState = if ($hasAllowRecording -or $hasAllowTranscription) {
        $allowRecordingValue = if ($hasAllowRecording) { [bool]$meetingPolicy.$recordingPropertyName } else { $false }
        $allowTranscriptionValue = if ($hasAllowTranscription) { [bool]$meetingPolicy.$transcriptionPropertyName } else { $false }
        if ($allowRecordingValue -or $allowTranscriptionValue) { 'enabled' } else { 'disabled' }
    }
    else {
        'not collected'
    }

    $teamsStatus = if ($groupDenied) { 'partial' } elseif ($hasAllowRecording -or $hasAllowTranscription) { 'reported' } else { 'partial' }
    $teamsReason = if ($groupDenied -and ($hasAllowRecording -or $hasAllowTranscription)) {
        'Team inventory returned HTTP 403 (Group.Read.All). Meeting policy was read from the connected Teams session.'
    }
    elseif ($groupDenied) {
        'Missing Graph permission: Group.Read.All. Teams meeting policy was not collected because no Teams session was connected.'
    }
    elseif ($teamsStatus -eq 'partial') {
        'Teams meeting policy metadata was not collected.'
    }
    else {
        ''
    }

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
    $m365AppsUri = "https://graph.microsoft.com/v1.0/reports/getM365AppUserCounts(period='D30')"
    try {
        # getM365AppUserCounts is the supported report endpoint (returns CSV rows).
        $m365Apps = @(Get-GraphReportRows -Uri $m365AppsUri)
    }
    catch {
        $m365Apps = @()
        Register-GraphReportDiagnostic -Uri $m365AppsUri -Message $_.Exception.Message -StatusCode $script:GraphLastHttpStatus
    }

    $m365Diagnostics = @(Get-GraphReportDiagnosticsForUri -UriPattern 'getM365AppUserCounts')
    $m365Reason = ''
    if (@($m365Apps).Count -eq 0) {
        $m365Reason = if (@($m365Diagnostics).Count -gt 0) {
            "M365 Apps report request failed: $($m365Diagnostics[-1].Message)"
        }
        else {
            'M365 Apps report endpoint returned no data.'
        }
    }

    return [ordered]@{
        Status                = if (@($m365Apps).Count -gt 0) { 'reported' } else { 'not collected' }
        Reason                = $m365Reason
        M365AppsUpdateChannel = if (@($m365Apps).Count -gt 0) { 'reported' } else { 'not collected' }
        VersionReadiness      = if (@($m365Apps).Count -gt 0) { 'reported' } else { 'not collected' }
        BrowserSignals        = 'not collected'
        EndpointSignals       = 'not collected'
        ReportCount           = @($m365Apps).Count
        ReportDiagnostics     = @($m365Diagnostics | ForEach-Object { $_.Message })
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
            ControlProfilesDetailedIncluded = [bool]$script:IncludeDetailedData
            ControlProfiles = @()
        }
    }

    $historyCount = if ($script:IncludeDetailedData) { 10 } else { 5 }
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
            ControlProfilesDetailedIncluded = [bool]$script:IncludeDetailedData
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
    $controlProfiles = @(
        Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles' -Query @{ '$top' = '200' }
    )
    $scoreAssessment = Get-SecureScoreOpenControlAssessment -ControlProfiles $controlProfiles -ControlScores $controlScores
    $gapSource = @($scoreAssessment.TopGaps)
    $controlGaps = @($gapSource | ForEach-Object {
        $gapClassification = [string](Get-ObjectPropertyValue -InputObject $_ -Name 'gapClassification' -Default 'open')
        [ordered]@{
            ControlName = Get-ObjectPropertyValue -InputObject $_ -Name 'id' -Default 'not collected'
            Score = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'currentScore') -Default 0
            MaxScore = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'maxScore') -Default 0
            Gap = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'pointsAvailable') -Default 0
            PointsAvailable = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'pointsAvailable') -Default 0
            GapClassification = $gapClassification
            AddressedWithPointsRemaining = ($gapClassification -eq 'addressed-with-points-remaining')
            AddressedState = Get-ObjectPropertyValue -InputObject $_ -Name 'addressedState' -Default ''
        }
    })

    $topOpenProfiles = @($gapSource | ForEach-Object {
        $gapClassification = [string](Get-ObjectPropertyValue -InputObject $_ -Name 'gapClassification' -Default 'open')
        [ordered]@{
            Id = Get-ObjectPropertyValue -InputObject $_ -Name 'id' -Default 'not collected'
            Title = Get-ObjectPropertyValue -InputObject $_ -Name 'title' -Default 'not collected'
            Rank = Get-IntValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'rank') -Default 0
            MaxScore = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'maxScore') -Default 0
            CurrentScore = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'currentScore') -Default 0
            PointsAvailable = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'pointsAvailable') -Default 0
            ControlCategory = Get-ObjectPropertyValue -InputObject $_ -Name 'controlCategory' -Default 'not collected'
            ImplementationCost = Get-ObjectPropertyValue -InputObject $_ -Name 'implementationCost' -Default 'not collected'
            UserImpact = Get-ObjectPropertyValue -InputObject $_ -Name 'userImpact' -Default 'not collected'
            Service = Get-ObjectPropertyValue -InputObject $_ -Name 'service' -Default 'not collected'
            Remediation = Get-ObjectPropertyValue -InputObject $_ -Name 'remediation' -Default 'not collected'
            ActionUrl = Get-ObjectPropertyValue -InputObject $_ -Name 'actionUrl' -Default 'not collected'
            GapClassification = $gapClassification
            AddressedWithPointsRemaining = ($gapClassification -eq 'addressed-with-points-remaining')
            AddressedState = Get-ObjectPropertyValue -InputObject $_ -Name 'addressedState' -Default ''
        }
    })

    $controlProfilesPayload = if ($script:IncludeDetailedData) {
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
            TotalControls = $scoreAssessment.ScoredControlCount
            ControlsBelowMaxScore = $scoreAssessment.ControlsBelowMaxScore
            TopControlGaps = @($controlGaps)
        }
        ControlProfilesSummary = [ordered]@{
            TotalProfiles = @($controlProfiles).Count
            OpenProfiles = $scoreAssessment.OpenControlCount
            ScoredControlCount = $scoreAssessment.ScoredControlCount
            FullyAchievedCount = $scoreAssessment.FullyAchievedCount
            AddressedWithPointsRemaining = $scoreAssessment.AddressedWithPointsRemaining
            NotApplicableCount = $scoreAssessment.NotApplicableCount
            TotalPointsAvailable = $scoreAssessment.TotalPointsAvailable
            TopOpenProfiles = $topOpenProfiles
        }
        ControlProfilesDetailedIncluded = [bool]$script:IncludeDetailedData
        ControlProfiles = $controlProfilesPayload
    }
}

function Get-Office365ActiveUserCountsLatestRow {
    # Daily getOffice365ActiveUserCounts rows often leave the current Office 365
    # cell blank. Adoption counts use getOffice365ServicesUserCounts instead.
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

function Get-Office365ServicesUserCountsRow {
    # One summary row for the period. Prefer a row whose Office 365 Active cell is populated.
    param([ValidateSet('D7', 'D30', 'D90', 'D180')][string]$Period = 'D30')

    $rows = @(Get-GraphReportRows -Uri "https://graph.microsoft.com/v1.0/reports/getOffice365ServicesUserCounts(period='$Period')")
    $candidates = @($rows | Where-Object { $null -ne $_ })
    if ($candidates.Count -eq 0) { return $null }

    $withActive = @($candidates | Where-Object {
        $activeValue = Get-RowValueNormalized -Row $_ -CandidateNames @('Office 365 Active', 'office365Active')
        $null -ne $activeValue -and -not [string]::IsNullOrWhiteSpace([string]$activeValue)
    })
    $pool = if ($withActive.Count -gt 0) { $withActive } else { $candidates }
    return $pool |
        Sort-Object {
            $reportDateValue = Get-RowValueNormalized -Row $_ -CandidateNames @('reportRefreshDate', 'report refresh date', 'reportDate', 'report date')
            try { [datetime]$reportDateValue } catch { [datetime]::MinValue }
        } -Descending |
        Select-Object -First 1
}

function Get-ActiveUserCountFromRow {
    # A blank report cell is not collected. It is never coerced to zero.
    param(
        [AllowNull()][object]$Row,
        [Parameter(Mandatory)][string[]]$CandidateNames
    )

    if ($null -eq $Row) { return $null }
    $value = Get-RowValueNormalized -Row $Row -CandidateNames $CandidateNames
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return $null }

    $parsedInt = 0
    if ([int]::TryParse([string]$value, [ref]$parsedInt)) { return $parsedInt }

    $parsedDouble = 0.0
    if ([double]::TryParse([string]$value, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedDouble)) {
        return [int][math]::Round($parsedDouble, 0)
    }

    return $null
}

function Get-AdoptionSignals {
    $latest = $null
    try {
        $latest = Get-Office365ServicesUserCountsRow -Period 'D30'
    }
    catch {
        $latest = $null
    }

    $copilot = Get-CopilotUsageSnapshot -Period 'D30'
    $copilotReported = ([string](Get-ObjectPropertyValue -InputObject $copilot -Name 'Status' -Default 'not collected') -eq 'reported')
    $copilotReason = [string](Get-ObjectPropertyValue -InputObject $copilot -Name 'Reason' -Default '')

    $activeUserDiagnostics = @(Get-GraphReportDiagnosticsForUri -UriPattern 'getOffice365ServicesUserCounts')
    if ($null -eq $latest) {
        $status = if ($copilotReported) { 'partial' } else { 'not collected' }
        $reason = 'Office 365 services user count report was unavailable.'
        if (@($activeUserDiagnostics).Count -gt 0) {
            $reason = "$reason $($activeUserDiagnostics[-1].Message)"
        }
        if (-not $copilotReported -and -not [string]::IsNullOrWhiteSpace($copilotReason)) {
            $reason = "$reason Copilot usage was not collected: $copilotReason"
        }

        return [ordered]@{
            Status                  = $status
            Reason                  = $reason
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
            ReportDiagnostics       = @(
                @($activeUserDiagnostics | ForEach-Object { $_.Message }) +
                @(Get-ObjectPropertyValue -InputObject $copilot -Name 'ReportDiagnostics' -Default @())
            )
        }
    }

    $status = if ($copilotReported) { 'reported' } else { 'partial' }
    $reason = if ($copilotReported) {
        ''
    }
    elseif (-not [string]::IsNullOrWhiteSpace($copilotReason)) {
        "Copilot usage was not collected: $copilotReason"
    }
    else {
        'Copilot usage was not collected.'
    }

    return [ordered]@{
        Status                  = $status
        Reason                  = $reason
        # getOffice365ServicesUserCounts(period='D30') reports users active in the
        # period. Blank Active cells stay null. The daily active-user report's last
        # row often has a blank Office 365 column and is not used as zero.
        ActiveUsers30Days      = Get-ActiveUserCountFromRow -Row $latest -CandidateNames @('Office 365 Active', 'office365Active')
        TeamsActiveUsers       = Get-ActiveUserCountFromRow -Row $latest -CandidateNames @('Teams Active', 'teamsActive')
        SharePointActiveUsers  = Get-ActiveUserCountFromRow -Row $latest -CandidateNames @('SharePoint Active', 'sharePointActive')
        OneDriveActiveUsers    = Get-ActiveUserCountFromRow -Row $latest -CandidateNames @('OneDrive Active', 'oneDriveActive')
        OutlookActiveUsers     = Get-ActiveUserCountFromRow -Row $latest -CandidateNames @('Exchange Active', 'exchangeActive')
        ReportSource           = 'office365ServicesUserCounts'
        CopilotActiveUsers30Days = $copilot.ActiveUsers
        CopilotUsageStatus      = $copilot.Status
        CopilotReportSource     = $copilot.Source
        CopilotReportDate       = $copilot.ReportDate
        CopilotCollectionReason = if ([string]::IsNullOrWhiteSpace($copilot.Reason)) { 'reported' } else { $copilot.Reason }
        ReportDiagnostics       = @(
            @($activeUserDiagnostics | ForEach-Object { $_.Message }) +
            @(Get-ObjectPropertyValue -InputObject $copilot -Name 'ReportDiagnostics' -Default @())
        )
    }
}

function Test-PhishingResistantAuthenticationMethod {
    param([AllowNull()][object]$Methods)

    foreach ($method in @($Methods)) {
        $name = [string]$method
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($name -match '^(fido2|windowsHelloForBusiness|passKeyDeviceBound|passKeyDeviceBoundAuthenticator|certificateBasedAuthentication)$') {
            return $true
        }
        if ($name -like 'x509Certificate*') { return $true }
    }
    return $false
}

function Get-PhishingResistantMfaCoverageResult {
    param(
        [AllowNull()][object[]]$RegistrationDetails,
        [string]$FailureReason = ''
    )

    if (-not [string]::IsNullOrWhiteSpace($FailureReason)) {
        Register-GraphReportDiagnostic -Uri 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails' -Message $FailureReason -Section 'IdentityAccessAdvanced' -DiagnosticOnly
        return [ordered]@{
            Coverage    = 'not collected'
            UserCount   = 'not collected'
            Denominator = 'not collected'
            Reason      = $FailureReason
        }
    }

    $enabledMembers = @(Get-EnabledMemberUsers)
    $population = @($enabledMembers).Count
    $enabledIds = @{}
    foreach ($user in @($enabledMembers)) {
        $id = Get-DirectoryObjectId -User $user
        if (-not [string]::IsNullOrWhiteSpace($id)) { $enabledIds[$id.ToLowerInvariant()] = $true }
    }

    $coveredIds = @{}
    foreach ($row in @($RegistrationDetails)) {
        if ($null -eq $row) { continue }
        $rowId = Get-DirectoryObjectId -User $row
        if ([string]::IsNullOrWhiteSpace($rowId)) { continue }
        $key = $rowId.ToLowerInvariant()
        if (-not $enabledIds.ContainsKey($key)) { continue }
        $methods = Get-ObjectPropertyValue -InputObject $row -Name 'methodsRegistered' -Default $null
        if ($null -eq $methods) { $methods = Get-ObjectPropertyValue -InputObject $row -Name 'MethodsRegistered' -Default $null }
        if (Test-PhishingResistantAuthenticationMethod -Methods $methods) { $coveredIds[$key] = $true }
    }

    $covered = @($coveredIds.Keys).Count
    if ($population -le 0) {
        return [ordered]@{
            Coverage    = 'not collected'
            UserCount   = $covered
            Denominator = 'not collected'
            Reason      = 'Enabled member users were not available, so phishing-resistant MFA coverage was not collected.'
        }
    }

    return [ordered]@{
        Coverage    = [math]::Round(($covered / $population) * 100, 2)
        UserCount   = $covered
        Denominator = $population
        Reason      = ''
    }
}

function Get-PrivilegedDirectoryRoleTemplateIds {
    return @(
        '62e90394-69f5-4237-9190-012177145e10' # Global Administrator
        'e8611ab8-c189-46e8-94e1-60213ab1f814' # Privileged Role Administrator
        '194ae4cb-b126-40b2-bd5b-6091b380977d' # Security Administrator
        '29232cdf-9323-42fd-ade2-1d097af3e4de' # Exchange Administrator
        'f28a1f50-f6e7-4571-818b-6a12f2af6b6c' # SharePoint Administrator
        'fe930be7-5e62-47db-91af-98c3a49a38b1' # User Administrator
        '69091246-20e8-4a56-aa4d-066075b2a7a8' # Teams Administrator
        '17315797-102d-40b4-93e0-432062caca18' # Compliance Administrator
        'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9' # Conditional Access Administrator
        'c4e39bd9-1100-46d3-8c65-fb160da0071f' # Authentication Administrator
        '7be44c8a-adaf-4e2a-84d6-ab2649e08a13' # Privileged Authentication Administrator
    )
}

function Get-DirectoryRoleDefinitionId {
    param([AllowNull()][object]$Assignment)

    $roleId = [string](Get-ObjectPropertyValue -InputObject $Assignment -Name 'roleDefinitionId' -Default '')
    if ([string]::IsNullOrWhiteSpace($roleId)) {
        $roleId = [string](Get-ObjectPropertyValue -InputObject $Assignment -Name 'RoleDefinitionId' -Default '')
    }
    return $roleId.ToLowerInvariant()
}

function Get-PimDirectorySignal {
    $notCollected = [ordered]@{
        Signal = 'not collected'
        EligibleAssignmentCount = 'not collected'
        PrivilegedRoleEligibleCount = 'not collected'
        PermanentPrivilegedAssignmentCount = 'not collected'
        ActiveAssignmentScheduleCount = 'not collected'
        Reason = 'PIM role eligibility was not collected.'
    }

    try {
        $eligibility = @(Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilitySchedules')
    }
    catch {
        $eligibility = @()
        $script:GraphLastErrorMessage = [string]$_.Exception.Message
    }

    $eligibilityStatusVariable = Get-Variable -Name GraphLastHttpStatus -Scope Script -ErrorAction SilentlyContinue
    $eligibilityStatus = if ($null -ne $eligibilityStatusVariable) { $eligibilityStatusVariable.Value } else { $null }
    $eligibilityError = Get-ScriptStringParameter -Name 'GraphLastErrorMessage'
    $eligibilityCode = Get-ScriptStringParameter -Name 'GraphLastErrorCode'
    $licenseOrDenied = ($eligibilityStatus -in @(400, 401, 403)) -or (
        $eligibilityStatus -ne 200 -and "$eligibilityError $eligibilityCode" -match 'AadPremiumLicenseRequired|Premium license|AadPremium'
    )
    $eligibilityFailed = (@($eligibility).Count -eq 0) -and ($licenseOrDenied -or $null -eq $eligibilityStatus -or $eligibilityStatus -ne 200)
    if ($eligibilityFailed) {
        $reason = $eligibilityError
        if ([string]::IsNullOrWhiteSpace($reason)) {
            if (-not [string]::IsNullOrWhiteSpace($eligibilityCode)) { $reason = $eligibilityCode }
            elseif ($eligibilityStatus) { $reason = "HTTP $eligibilityStatus. RoleEligibilitySchedule.Read.Directory is required, or the tenant does not have Entra ID P2." }
            else { $reason = 'Role eligibility schedules were not returned.' }
        }
        Register-GraphReportDiagnostic -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilitySchedules' -Message $reason -StatusCode $eligibilityStatus -Section 'IdentityAccessAdvanced' -DiagnosticOnly
        $notCollected.Reason = $reason
        return $notCollected
    }

    $privileged = @{}
    foreach ($roleId in @(Get-PrivilegedDirectoryRoleTemplateIds)) { $privileged[$roleId.ToLowerInvariant()] = $true }

    $privilegedRoles = @{}
    $privilegedEligible = 0
    foreach ($schedule in @($eligibility)) {
        if ($null -eq $schedule) { continue }
        $roleId = Get-DirectoryRoleDefinitionId -Assignment $schedule
        if ($privileged.ContainsKey($roleId)) {
            $privilegedEligible++
            $privilegedRoles[$roleId] = $true
        }
    }

    $activeScheduleCount = 'not collected'
    try {
        $activeSchedules = @(Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentSchedules')
        $activeStatusVariable = Get-Variable -Name GraphLastHttpStatus -Scope Script -ErrorAction SilentlyContinue
        $activeStatus = if ($null -ne $activeStatusVariable) { $activeStatusVariable.Value } else { $null }
        if (@($activeSchedules).Count -gt 0 -or $activeStatus -eq 200) {
            $activeScheduleCount = @($activeSchedules).Count
        }
        else {
            $activeError = Get-ScriptStringParameter -Name 'GraphLastErrorMessage'
            if ([string]::IsNullOrWhiteSpace($activeError)) { $activeError = "HTTP $activeStatus" }
            Register-GraphReportDiagnostic -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentSchedules' -Message $activeError -StatusCode $activeStatus -Section 'IdentityAccessAdvanced' -DiagnosticOnly
        }
    }
    catch {
        Register-GraphReportDiagnostic -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentSchedules' -Message $_.Exception.Message -Section 'IdentityAccessAdvanced' -DiagnosticOnly
    }

    $permanentCount = 'not collected'
    $permanentReason = ''
    try {
        $assignments = @(Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments')
        $assignmentStatusVariable = Get-Variable -Name GraphLastHttpStatus -Scope Script -ErrorAction SilentlyContinue
        $assignmentStatus = if ($null -ne $assignmentStatusVariable) { $assignmentStatusVariable.Value } else { $null }
        if (@($assignments).Count -gt 0 -or $assignmentStatus -eq 200) {
            $permanentCount = @($assignments | Where-Object { $privileged.ContainsKey((Get-DirectoryRoleDefinitionId -Assignment $_)) }).Count
        }
        else {
            $permanentReason = Get-ScriptStringParameter -Name 'GraphLastErrorMessage'
            if ([string]::IsNullOrWhiteSpace($permanentReason)) { $permanentReason = "HTTP $assignmentStatus. Permanent role assignments were not collected." }
            Register-GraphReportDiagnostic -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments' -Message $permanentReason -StatusCode $assignmentStatus -Section 'IdentityAccessAdvanced' -DiagnosticOnly
        }
    }
    catch {
        $permanentReason = [string]$_.Exception.Message
        Register-GraphReportDiagnostic -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments' -Message $permanentReason -Section 'IdentityAccessAdvanced' -DiagnosticOnly
    }

    return [ordered]@{
        Signal = if ($privilegedEligible -gt 0) { 'in use' } else { 'not in use' }
        EligibleAssignmentCount = @($eligibility).Count
        PrivilegedRoleEligibleCount = @($privilegedRoles.Keys).Count
        PermanentPrivilegedAssignmentCount = $permanentCount
        ActiveAssignmentScheduleCount = $activeScheduleCount
        Reason = $permanentReason
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
            PhishingResistantMfaUserCount = 'not collected'
            PhishingResistantMfaDenominator = 'not collected'
            PhishingResistantMfaReason = 'Identity baseline data was not collected.'
            PimSignal            = 'not collected'
            PimEligibleAssignmentCount = 'not collected'
            PimPrivilegedRoleEligibleCount = 'not collected'
            PimPermanentPrivilegedAssignmentCount = 'not collected'
            PimCollectionReason = 'Identity baseline data was not collected.'
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

    $phishingCoverage = Get-ObjectPropertyValue -InputObject $identity -Name 'PhishingResistantMfaCoverage' -Default $null
    $phishingCount = Get-ObjectPropertyValue -InputObject $identity -Name 'PhishingResistantMfaUserCount' -Default $null
    $phishingDenominator = Get-ObjectPropertyValue -InputObject $identity -Name 'PhishingResistantMfaDenominator' -Default $null
    $phishingReason = [string](Get-ObjectPropertyValue -InputObject $identity -Name 'PhishingResistantMfaReason' -Default '')
    if ($null -eq $phishingCoverage) {
        $registrationRows = @(Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails')
        $fallbackFailure = ''
        if (@($registrationRows).Count -eq 0 -and $script:GraphLastHttpStatus -eq 403) {
            $fallbackFailure = Get-ScriptStringParameter -Name 'GraphLastErrorMessage'
            if ([string]::IsNullOrWhiteSpace($fallbackFailure)) {
                $fallbackFailure = 'HTTP 403. UserAuthenticationMethod.Read.All or AuditLog.Read.All is required.'
            }
        }
        $phishing = Get-PhishingResistantMfaCoverageResult -RegistrationDetails $(if ($fallbackFailure) { $null } else { $registrationRows }) -FailureReason $fallbackFailure
        $phishingCoverage = $phishing.Coverage
        $phishingCount = $phishing.UserCount
        $phishingDenominator = $phishing.Denominator
        $phishingReason = [string]$phishing.Reason
    }

    $pim = Get-PimDirectorySignal
    $advancedReasons = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-ValueCollected -Value $phishingCoverage)) {
        $advancedReasons.Add($(if ($phishingReason) { $phishingReason } else { 'Phishing-resistant MFA coverage was not collected.' })) | Out-Null
    }
    if (-not (Test-ValueCollected -Value $pim.Signal)) {
        $advancedReasons.Add([string]$pim.Reason) | Out-Null
    }
    $advancedStatus = if ($advancedReasons.Count -eq 0) { 'reported' } else { 'partial' }

    return [ordered]@{
        Status               = $advancedStatus
        Reason               = ($advancedReasons -join ' ')
        AdminRoleObjectsCount = $adminRoleCount
        MfaCapableUsers       = $mfaCapable
        MfaRegisteredUsers    = $mfaRegistered
        PasswordlessCapableUsers = $passwordless
        MfaPopulationUserCount = if ($mfaDenominator -gt 0) { $mfaDenominator } else { 'not collected' }
        MfaPopulationSource  = Get-ObjectPropertyValue -InputObject $identity -Name 'MfaPopulationSource' -Default 'not collected'
        MfaRegistrationPercent = if ($mfaDenominator -gt 0) { [math]::Round(($mfaRegistered / $mfaDenominator) * 100, 2) } else { 0 }
        MfaRegisteredMemberPercent = Get-ObjectPropertyValue -InputObject $identity -Name 'MfaRegisteredMemberPercent' -Default 'not collected'
        MfaRegisteredGuestPercent = Get-ObjectPropertyValue -InputObject $identity -Name 'MfaRegisteredGuestPercent' -Default 'not collected'
        MfaRegisteredAdminPercent = Get-ObjectPropertyValue -InputObject $identity -Name 'MfaRegisteredAdminPercent' -Default 'not collected'
        MfaRegisteredEnabledMemberPercent = Get-ObjectPropertyValue -InputObject $identity -Name 'MfaRegisteredEnabledMemberPercent' -Default 'not collected'
        MfaRegisteredEnabledMemberCount = Get-ObjectPropertyValue -InputObject $identity -Name 'MfaRegisteredEnabledMemberCount' -Default 'not collected'
        MfaEnabledMemberPopulationUserCount = Get-ObjectPropertyValue -InputObject $identity -Name 'MfaEnabledMemberPopulationUserCount' -Default 'not collected'
        MfaRegisteredMemberCount = Get-ObjectPropertyValue -InputObject $identity -Name 'MfaRegisteredMemberCount' -Default 'not collected'
        MfaRegisteredGuestCount = Get-ObjectPropertyValue -InputObject $identity -Name 'MfaRegisteredGuestCount' -Default 'not collected'
        MfaRegisteredAdminCount = Get-ObjectPropertyValue -InputObject $identity -Name 'MfaRegisteredAdminCount' -Default 'not collected'
        MfaRegistrationPercentNote = 'MfaRegistrationPercent includes guests and is informational. Scoring uses MfaRegisteredEnabledMemberPercent. Guests and disabled accounts are excluded from that percentage.'
        PhishingResistantMfaCoverage = $phishingCoverage
        PhishingResistantMfaUserCount = $phishingCount
        PhishingResistantMfaDenominator = $phishingDenominator
        PhishingResistantMfaReason = if ($phishingReason) { $phishingReason } else { 'reported' }
        PimSignal = $pim.Signal
        PimEligibleAssignmentCount = $pim.EligibleAssignmentCount
        PimPrivilegedRoleEligibleCount = $pim.PrivilegedRoleEligibleCount
        PimPermanentPrivilegedAssignmentCount = $pim.PermanentPrivilegedAssignmentCount
        PimActiveAssignmentScheduleCount = $pim.ActiveAssignmentScheduleCount
        PimCollectionReason = if ([string]::IsNullOrWhiteSpace([string]$pim.Reason)) { 'reported' } else { $pim.Reason }
        AdminRoleCoverageNote = 'Admin role counts are derived from activated directory roles and active memberships only; PIM-eligible assignments are not included.'
    }
}

function Get-DataProtectionAdvanced {
    $governance = $script:Report.DataGovernance

    if (-not (Test-SectionCollected -SectionData $governance)) {
        return [ordered]@{
            Status                 = 'not collected'
            Reason                 = 'Data governance baseline data was not collected.'
            SensitivityLabelsPublished = $null
            SensitivityLabelsTotal = $null
            LabelPoliciesPublished = $null
            LabelPoliciesTotal     = $null
            LabelCollectionReason  = 'Data governance baseline data was not collected.'
            DlpPoliciesEnforced    = 'not collected'
            DlpPoliciesTotal       = 'not collected'
            RetentionPoliciesEnabled = 'not collected'
            RetentionPoliciesTotal = 'not collected'
            eDiscoverySignal       = 'not collected'
            InsiderRiskSignal      = 'not collected'
        }
    }

    $published = ConvertTo-CollectedInt -Value (Get-ObjectPropertyValue -InputObject $governance -Name 'SensitivityLabelsPublished')
    $labelsTotal = ConvertTo-CollectedInt -Value (Get-ObjectPropertyValue -InputObject $governance -Name 'SensitivityLabelsTotal')
    $labelPoliciesPublished = ConvertTo-CollectedInt -Value (Get-ObjectPropertyValue -InputObject $governance -Name 'LabelPoliciesPublished')
    $labelPoliciesTotal = ConvertTo-CollectedInt -Value (Get-ObjectPropertyValue -InputObject $governance -Name 'LabelPoliciesTotal')

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
        SensitivityLabelsPublished = $published
        SensitivityLabelsTotal = $labelsTotal
        LabelPoliciesPublished = $labelPoliciesPublished
        LabelPoliciesTotal     = $labelPoliciesTotal
        LabelCollectionReason  = if ($null -eq $published) { [string](Get-ObjectPropertyValue -InputObject $governance -Name 'LabelCollectionReason' -Default 'Sensitivity labels were not collected.') } else { 'reported' }
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

    $explicitInventoryStatus = [string](Get-ObjectPropertyValue -InputObject $sharePoint -Name 'SiteInventoryStatus' -Default '')
    if (-not [string]::IsNullOrWhiteSpace($explicitInventoryStatus)) {
        $siteInventoryStatus = $explicitInventoryStatus
    }
    else {
        $totalSiteCountValue = Get-ObjectPropertyValue -InputObject $sharePoint -Name 'TotalSiteCount' -Default 'not collected'
        $hasSiteCollection = $sharePoint -and (Test-HashtableKey -InputObject $sharePoint -Key 'Sites') -and $sharePoint.Sites -isnot [string]
        $suppliedSiteCount = if ($hasSiteCollection) { @($sharePoint.Sites).Count } else { 0 }
        # A numeric total, or a supplied site list, is a collected inventory.
        # An empty list with no total is not a sample of zero sites.
        if ((Test-ValueCollected -Value $totalSiteCountValue) -or $suppliedSiteCount -gt 0) {
            $siteInventoryStatus = 'reported'
        }
        else {
            $siteInventoryStatus = 'not collected'
        }
    }
    if ($siteInventoryStatus -ne 'reported') {
        $externalSharingLevel = [string](Get-ObjectPropertyValue -InputObject $sharePoint -Name 'ExternalSharingLevel' -Default 'not collected')
        $anyoneLinksPossible = 'not collected'
        $anyoneLinkMeasurement = 'not collected'
        if ($externalSharingLevel -eq 'ExternalUserAndGuestSharing') {
            $anyoneLinksPossible = $true
            $anyoneLinkMeasurement = 'tenant-settings'
        }
        elseif ($externalSharingLevel -in @('Disabled', 'ExternalUserSharingOnly', 'ExistingExternalUserSharingOnly')) {
            $anyoneLinksPossible = $false
            $anyoneLinkMeasurement = 'tenant-settings'
        }

        $inventoryReason = [string](Get-ObjectPropertyValue -InputObject $sharePoint -Name 'SiteInventoryReason' -Default 'SharePoint site inventory was not collected.')
        $exposureReason = $inventoryReason
        if ($anyoneLinksPossible -eq $true) {
            $exposureReason = "$inventoryReason Tenant sharingCapability ExternalUserAndGuestSharing means anyone links are possible. Anyone-link and external-sharing site counts were not collected."
        }
        elseif ($anyoneLinkMeasurement -eq 'tenant-settings') {
            $exposureReason = "$inventoryReason Anyone-link and external-sharing site counts were not collected."
        }

        return [ordered]@{
            Status = 'partial'
            Reason = $exposureReason.Trim()
            SampledSites = 'not collected'
            ExternalSharingSites = 'not collected'
            AnyoneLinkSites = 'not collected'
            AnyoneLinksPossible = $anyoneLinksPossible
            AnyoneLinkMeasurement = $anyoneLinkMeasurement
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

    # Anyone links are ExternalUserAndGuestSharing (2) only. External sharing is 1, 2, or 3.
    # Values arrive as enums, ints, numeric strings, or names; AnonymousAccess is not valid.
    $sharingCapabilityNames = @($sampleSites | ForEach-Object {
        ConvertTo-SharingCapabilityName -Value (Get-ObjectPropertyValue -InputObject $_ -Name 'SharingCapability' -Default $null)
    })
    $externalSharingSites = @($sharingCapabilityNames | Where-Object { $_ -in @('ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly') }).Count
    $anyoneLinkSites = @($sharingCapabilityNames | Where-Object { $_ -eq 'ExternalUserAndGuestSharing' }).Count
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
        AnyoneLinksPossible = ($anyoneLinkSites -gt 0)
        AnyoneLinkMeasurement = 'site-inventory'
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
    $hasReportsReadPermission = Test-GraphPermissionsAvailable -RequiredPermissions @('Reports.Read.All')
    $managedDevices = @()
    if ($hasDeviceReadPermission) {
        $managedDevices = @((Get-GraphRestAll -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices' -Query @{ '$select' = 'id,complianceState,operatingSystem' }))
    }
    $managedDevicesReadable = $hasDeviceReadPermission -and (@($managedDevices).Count -gt 0)
    $compliant = @($managedDevices | Where-Object { $_.PSObject.Properties.Name -contains 'ComplianceState' -and $_.ComplianceState -eq 'compliant' }).Count
    $windows = @($managedDevices | Where-Object { $_.PSObject.Properties.Name -contains 'OperatingSystem' -and $_.OperatingSystem -match 'Windows' }).Count

    $appsReport = if ($hasReportsReadPermission) {
        @(Get-GraphReportRows -Uri "https://graph.microsoft.com/v1.0/reports/getM365AppUserCounts(period='D30')")
    }
    else {
        @()
    }

    if (-not $managedDevicesReadable -and @($appsReport).Count -eq 0) {
        $reasonParts = [System.Collections.Generic.List[string]]::new()
        if (-not $hasDeviceReadPermission) {
            $reasonParts.Add('Device management signal is unavailable because Graph permission DeviceManagementManagedDevices.Read.All is missing.') | Out-Null
        }
        else {
            $reasonParts.Add('Device management data could not be collected from managedDevices.') | Out-Null
        }

        if (-not $hasReportsReadPermission) {
            $reasonParts.Add('M365 Apps report data is unavailable because Graph permission Reports.Read.All is missing.') | Out-Null
        }
        else {
            $reasonParts.Add('M365 Apps report data was unavailable.') | Out-Null
        }

        return [ordered]@{
            Status                = 'not collected'
            Reason                = ($reasonParts -join ' ')
            ManagedDeviceCount    = 'not collected'
            CompliantDeviceCount  = 'not collected'
            WindowsDeviceCount    = 'not collected'
            DeviceCompliancePercent = 'not collected'
            M365AppReportAvailable = if ($hasReportsReadPermission) { $false } else { 'not collected' }
            DeviceManagementSignal = if (-not $hasDeviceReadPermission) { 'not collected (missing Graph permission: DeviceManagementManagedDevices.Read.All)' } else { 'not collected (requires DeviceManagement managedDevices read permissions)' }
            BrowserReadinessSignal = 'not collected'
        }
    }

    $endpointReason = 'Browser readiness signal was not collected.'
    if (-not $hasDeviceReadPermission) {
        $endpointReason = "Device management signal is unavailable because Graph permission DeviceManagementManagedDevices.Read.All is missing. $endpointReason"
    }

    return [ordered]@{
        Status = 'partial'
        Reason = $endpointReason
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
    try { $latestD7 = Get-Office365ServicesUserCountsRow -Period 'D7' } catch { $latestD7 = $null }
    try { $latestD30 = Get-Office365ServicesUserCountsRow -Period 'D30' } catch { $latestD30 = $null }
    try { $latestD90 = Get-Office365ServicesUserCountsRow -Period 'D90' } catch { $latestD90 = $null }
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
        ActiveUsersD7  = if ($latestD7) { Get-ActiveUserCountFromRow -Row $latestD7 -CandidateNames @('Office 365 Active', 'office365Active') } else { 'not collected' }
        ActiveUsersD30 = if ($latestD30) { Get-ActiveUserCountFromRow -Row $latestD30 -CandidateNames @('Office 365 Active', 'office365Active') } else { 'not collected' }
        ActiveUsersD90 = if ($latestD90) { Get-ActiveUserCountFromRow -Row $latestD90 -CandidateNames @('Office 365 Active', 'office365Active') } else { 'not collected' }
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
            Reason = 'Missing Graph permission: Directory.Read.All. Service principals, OAuth grants, high-risk grants, stale grant signal, and ownerless app signal were not collected.'
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
    # All-user percent, including guests. Kept as informational output. Scoring uses
    # MfaRegisteredEnabledMemberPercent instead.
    $mfaPercent = if ($mfaDenominator -gt 0 -and $mfaRegisteredCountCollected) { [math]::Round(($mfaRegisteredCount / $mfaDenominator) * 100, 2) } else { 'not collected' }
    $enabledMemberCountValue = Get-ObjectPropertyValue -InputObject $script:Report.Licensing -Name 'PrereqReadyDenominator' -Default 'not collected'
    $enabledMemberCount = 0
    if (Test-ValueCollected -Value $enabledMemberCountValue) {
        $enabledMemberCount = Get-IntValue -Value $enabledMemberCountValue -Default 0
    }
    else {
        try {
            $enabledMemberCount = @(Get-EnabledMemberUsers).Count
        }
        catch {
            $enabledMemberCount = 0
        }
    }
    # NOTE: getOffice365ServicesUserCounts(period='D30') supplies Office 365 Active and
    # the per-service Active columns for the period. Blank cells stay null, not zero.
    # Divide by enabled members. Guests and disabled accounts are not the adoption base.
    $activeUserPercent = if ($enabledMemberCount -gt 0 -and (Test-ValueCollected -Value $activeUsers30DaysValue)) { [math]::Round((([double](Get-IntValue -Value $activeUsers30DaysValue -Default 0)) / $enabledMemberCount) * 100, 2) } else { 'not collected' }

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
        MfaRegisteredEnabledMemberPercent = Get-ObjectPropertyValue -InputObject $id -Name 'MfaRegisteredEnabledMemberPercent' -Default 'not collected'
        MfaRegisteredMemberPercent = Get-ObjectPropertyValue -InputObject $id -Name 'MfaRegisteredMemberPercent' -Default 'not collected'
        MfaRegisteredGuestPercent = Get-ObjectPropertyValue -InputObject $id -Name 'MfaRegisteredGuestPercent' -Default 'not collected'
        MfaRegisteredAdminPercent = Get-ObjectPropertyValue -InputObject $id -Name 'MfaRegisteredAdminPercent' -Default 'not collected'
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
        ActiveUserPercentDenominator = if ($enabledMemberCount -gt 0) { $enabledMemberCount } else { 'not collected' }
        ActiveUserPercentDenominatorDefinition = 'Enabled member users. Guests are excluded. Disabled accounts are excluded when accountEnabled is present.'
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
    $mfaEnabledMemberPercentValue = Get-ObjectPropertyValue -InputObject $flags -Name 'MfaRegisteredEnabledMemberPercent' -Default 'not collected'
    $caEnabledPoliciesValue = Get-ObjectPropertyValue -InputObject $flags -Name 'CaPoliciesEnabled' -Default 'not collected'
    $mfaRegisteredPercent = if (Test-ValueCollected -Value $mfaRegisteredPercentValue) { Get-DoubleValue -Value $mfaRegisteredPercentValue -Default 0 } else { $null }
    $mfaEnabledMemberPercent = if (Test-ValueCollected -Value $mfaEnabledMemberPercentValue) { Get-DoubleValue -Value $mfaEnabledMemberPercentValue -Default 0 } else { $null }
    $caEnabledPolicies = if (Test-ValueCollected -Value $caEnabledPoliciesValue) { Get-IntValue -Value $caEnabledPoliciesValue -Default 0 } else { $null }
    $identityAdvancedMfaPercent = if (Test-ValueCollected -Value $advancedMfaCoverage) { Get-DoubleValue -Value $advancedMfaCoverage -Default 0 } else { $null }

    $identityScoreRaw = 0
    $identityMeasuredWeight = 0
    $identityMeasured = $false
    # Score enabled members. The all-user percent includes guests and stays informational.
    if ($null -ne $mfaEnabledMemberPercent) {
        $identityMeasured = $true
        $identityMeasuredWeight += [int]$scoring.IdentityAccess.MfaRegisteredWeight
        if ($mfaEnabledMemberPercent -ge [double]$scoring.IdentityAccess.MfaRegisteredPercentThreshold) { $identityScoreRaw += [int]$scoring.IdentityAccess.MfaRegisteredWeight }
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
    $anyoneLinkMeasurement = [string](Get-ObjectPropertyValue -InputObject $spo -Name 'AnyoneLinkMeasurement' -Default '')
    $anyoneLinksPossibleValue = Get-ObjectPropertyValue -InputObject $spo -Name 'AnyoneLinksPossible' -Default 'not collected'
    $tenantSharingLevel = [string](Get-ObjectPropertyValue -InputObject $flags -Name 'ExternalSharingLevel' -Default 'not collected')
    $anyoneLinksTenantOnly = $false
    $anyoneLinksScored = $false
    $anyoneLinksPresent = $false
    if ($null -ne $anyoneLinkSites) {
        $anyoneLinksScored = $true
        $anyoneLinksPresent = ($anyoneLinkSites -gt 0)
    }
    elseif ($null -eq $anyoneLinkSites -and (
        ($anyoneLinkMeasurement -eq 'tenant-settings' -and $anyoneLinksPossibleValue -eq $true) -or
        ($tenantSharingLevel -eq 'ExternalUserAndGuestSharing' -and $anyoneLinkMeasurement -ne 'site-inventory')
    )) {
        # Site counts were not collected. ExternalUserAndGuestSharing means anyone links are possible.
        $anyoneLinksScored = $true
        $anyoneLinksPresent = $true
        $anyoneLinksTenantOnly = $true
    }
    elseif ($null -eq $anyoneLinkSites -and $anyoneLinkMeasurement -eq 'tenant-settings' -and $anyoneLinksPossibleValue -eq $false) {
        # Tenant settings block anyone links. Full credit for this factor.
        # Completeness stays scaled by TenantOnlyAnyoneLinkCompletenessFactor.
        $anyoneLinksScored = $true
        $anyoneLinksPresent = $false
        $anyoneLinksTenantOnly = $true
    }
    $restrictedSearchEnabledValue = Get-ObjectPropertyValue -InputObject $flags -Name 'RestrictedSearchEnabled' -Default 'not collected'
    $restrictedSearchEnabled = if (Test-ValueCollected -Value $restrictedSearchEnabledValue) { [bool]$restrictedSearchEnabledValue } else { $null }
    $contentMeasured = $false
    if ($null -ne $oneDriveCoveragePercent) {
        $contentMeasured = $true
        $contentMeasuredWeight += [int]$scoring.ContentExposure.OneDriveCoverageWeight
        if ($oneDriveCoveragePercent -ge [double]$scoring.ContentExposure.OneDriveCoverageThreshold) { $contentScoreRaw += [int]$scoring.ContentExposure.OneDriveCoverageWeight }
    }
    if ($anyoneLinksScored) {
        $contentMeasured = $true
        $contentMeasuredWeight += [int]$scoring.ContentExposure.NoAnyoneLinksWeight
        if (-not $anyoneLinksPresent) { $contentScoreRaw += [int]$scoring.ContentExposure.NoAnyoneLinksWeight } else { $contentScoreRaw += [int]$scoring.ContentExposure.AnyoneLinksPresentWeight }
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

    $identityPossibleWeight = [int]$scoring.IdentityAccess.MfaRegisteredWeight + [int]$scoring.IdentityAccess.AdvancedMfaWeight + [int]$scoring.IdentityAccess.ConditionalAccessWeight
    $governancePossibleWeight = [int]$scoring.DataGovernance.LabelsPublishedWeight + [int]$scoring.DataGovernance.DlpEnforcedWeight + [int]$scoring.DataGovernance.RetentionEnabledWeight
    $contentPossibleWeight = [int]$scoring.ContentExposure.OneDriveCoverageWeight + [int]$scoring.ContentExposure.NoAnyoneLinksWeight + [int]$scoring.ContentExposure.RestrictedSearchWeight
    $adoptionPossibleWeight = [int]$scoring.Adoption.ActiveUserHighWeight + [int]$scoring.Adoption.TrendReportedWeight
    $endpointPossibleWeight = 100
    $possibleMetricWeight = $identityPossibleWeight + $governancePossibleWeight + $contentPossibleWeight + $adoptionPossibleWeight + $endpointPossibleWeight

    $contentCompletenessWeight = [double]$contentMeasuredWeight
    if ($anyoneLinksTenantOnly) {
        $tenantFactor = Get-DoubleValue -Value (Get-ObjectPropertyValue -InputObject $scoring.ContentExposure -Name 'TenantOnlyAnyoneLinkCompletenessFactor' -Default 0.5) -Default 0.5
        if ($tenantFactor -lt 0) { $tenantFactor = 0 }
        if ($tenantFactor -gt 1) { $tenantFactor = 1 }
        $anyoneWeight = [double]$scoring.ContentExposure.NoAnyoneLinksWeight
        $contentCompletenessWeight = $contentCompletenessWeight - $anyoneWeight + ($anyoneWeight * $tenantFactor)
    }
    $endpointCompletenessWeight = if ($null -ne $deviceScore) { [double]$endpointPossibleWeight } else { [double]0 }
    $measuredMetricWeight = [double]$identityMeasuredWeight + [double]$governanceMeasuredWeight + $contentCompletenessWeight + [double]$adoptionMeasuredWeight + $endpointCompletenessWeight
    $completenessPercent = if ($possibleMetricWeight -gt 0) {
        [math]::Round(($measuredMetricWeight / $possibleMetricWeight) * 100, 2)
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
            MfaRegisteredPercent = if ($null -ne $mfaRegisteredPercent) { $mfaRegisteredPercent } else { 'not collected' }
            MfaRegisteredEnabledMemberPercent = if ($null -ne $mfaEnabledMemberPercent) { $mfaEnabledMemberPercent } else { 'not collected' }
            MfaScoreBasis = 'enabled-members'
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
            CompletenessScope = 'metric-level'
            EvaluatedSectionCount = @($evaluatedSections).Count
            CollectedSectionCount = @($collectedSections).Count
            MeasuredMetricWeight = [math]::Round($measuredMetricWeight, 2)
            PossibleMetricWeight = $possibleMetricWeight
            CompletenessPercent = $completenessPercent
            ConfidenceLevel = $confidenceLevel
            NotCollectedSections = @($notCollectedSections)
            AnyoneLinkMeasurement = if ($anyoneLinksTenantOnly) { 'tenant-settings' } elseif ($null -ne $anyoneLinkSites) { 'site-inventory' } else { 'not collected' }
            AnyoneLinkConfidence = if ($anyoneLinksTenantOnly) { 'reduced' } elseif ($null -ne $anyoneLinkSites) { 'standard' } else { 'not collected' }
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
        [string]$CurrentPath = '',
        [int]$Depth = 0,
        [int]$MaxDepth = 12
    )

    $results = [System.Collections.Generic.List[string]]::new()
    if ($null -eq $InputObject) { return @($results) }

    # Strings, dates, enums, and other value types are leaves. DateTime.Date is
    # another DateTime, so walking its properties never terminates.
    if ($InputObject -is [string]) {
        if ([string]$InputObject -like 'not collected*') {
            $results.Add($CurrentPath) | Out-Null
        }
        return @($results)
    }

    if ($InputObject -is [System.ValueType] -or $InputObject -is [System.Enum] -or $InputObject -is [datetime] -or $InputObject -is [datetimeoffset]) {
        return @($results)
    }

    if ($Depth -ge $MaxDepth) { return @($results) }

    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in @($InputObject.Keys)) {
            $childPath = if ([string]::IsNullOrWhiteSpace($CurrentPath)) { [string]$key } else { "$CurrentPath.$key" }
            foreach ($childResult in @(Get-NotCollectedFieldPaths -InputObject $InputObject[$key] -CurrentPath $childPath -Depth ($Depth + 1) -MaxDepth $MaxDepth)) {
                $results.Add($childResult) | Out-Null
            }
        }
        return @($results)
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
        $index = 0
        foreach ($item in @($InputObject)) {
            $childPath = if ([string]::IsNullOrWhiteSpace($CurrentPath)) { "[$index]" } else { "$CurrentPath[$index]" }
            foreach ($childResult in @(Get-NotCollectedFieldPaths -InputObject $item -CurrentPath $childPath -Depth ($Depth + 1) -MaxDepth $MaxDepth)) {
                $results.Add($childResult) | Out-Null
            }
            $index++
        }
        return @($results)
    }

    foreach ($property in @($InputObject.PSObject.Properties)) {
        $childPath = if ([string]::IsNullOrWhiteSpace($CurrentPath)) { [string]$property.Name } else { "$CurrentPath.$($property.Name)" }
        foreach ($childResult in @(Get-NotCollectedFieldPaths -InputObject $property.Value -CurrentPath $childPath -Depth ($Depth + 1) -MaxDepth $MaxDepth)) {
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
        SharePointCmdletsAvailable = [bool](Get-Module -ListAvailable -Name PnP.PowerShell)
        ExchangeCmdletsAvailable = [bool](Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)
        ComplianceCmdletsAvailable = [bool](Get-Command Connect-IPPSSession -ErrorAction SilentlyContinue)
    }

    $exchangeConnectionStatus = Get-ScriptStringParameter -Name 'ExchangeConnectionStatus'
    if ([string]::IsNullOrWhiteSpace($exchangeConnectionStatus)) { $exchangeConnectionStatus = 'not attempted' }
    $complianceConnectionStatus = Get-ScriptStringParameter -Name 'ComplianceConnectionStatus'
    if ([string]::IsNullOrWhiteSpace($complianceConnectionStatus)) { $complianceConnectionStatus = 'not attempted' }
    $exchangeConnectedVariable = Get-Variable -Name ExchangeConnected -Scope Script -ErrorAction SilentlyContinue
    $complianceConnectedVariable = Get-Variable -Name ComplianceConnected -Scope Script -ErrorAction SilentlyContinue
    $exchangeConnected = $exchangeConnectionStatus -eq 'connected' -or ($null -ne $exchangeConnectedVariable -and [bool]$exchangeConnectedVariable.Value)
    $complianceConnected = $complianceConnectionStatus -eq 'connected' -or ($null -ne $complianceConnectedVariable -and [bool]$complianceConnectedVariable.Value)
    if ($exchangeConnectionStatus -eq 'failed') { $exchangeConnected = $false }
    if ($complianceConnectionStatus -eq 'failed') { $complianceConnected = $false }

    $sessionAvailability = [ordered]@{
        GraphAuthenticated = -not [string]::IsNullOrWhiteSpace([string]$script:GraphAccessToken) -or $null -ne $script:MgContext
        TeamsConnected = [bool]$script:TeamsConnected
        SharePointConnected = [bool]$script:SharePointChildCollectionSucceeded
        ExchangeConnected = $exchangeConnected
        ComplianceConnected = $complianceConnected
        ExchangeConnectionStatus = $exchangeConnectionStatus
        ComplianceConnectionStatus = $complianceConnectionStatus
        ExchangeConnectionError = Get-ScriptStringParameter -Name 'ExchangeConnectionError'
        ComplianceConnectionError = Get-ScriptStringParameter -Name 'ComplianceConnectionError'
    }

    try {
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $connections = @((Get-ConnectionInformation -ErrorAction SilentlyContinue))
            # A recorded connection failure stays failed. A later probe must not
            # hide the assembly or sign-in error that left the session down.
            if ($exchangeConnectionStatus -ne 'failed') {
                $exchangeConnected = @($connections | Where-Object {
                    ($_.PSObject.Properties.Name -contains 'Name' -and [string]$_.Name -match 'ExchangeOnline') -or
                    ($_.PSObject.Properties.Name -contains 'ConnectionUri' -and [string]$_.ConnectionUri -match 'outlook\.office365\.com')
                }).Count -gt 0
                if ($exchangeConnected) {
                    $sessionAvailability.ExchangeConnected = $true
                    if ($exchangeConnectionStatus -eq 'not attempted') {
                        $sessionAvailability.ExchangeConnectionStatus = 'connected'
                    }
                }
            }
            if ($complianceConnectionStatus -ne 'failed') {
                $complianceConnected = @($connections | Where-Object {
                    ($_.PSObject.Properties.Name -contains 'Name' -and [string]$_.Name -match 'IPPSSession|SecurityCompliance') -or
                    ($_.PSObject.Properties.Name -contains 'ConnectionUri' -and [string]$_.ConnectionUri -match 'ps\.compliance\.protection\.outlook\.com')
                }).Count -gt 0
                if ($complianceConnected) {
                    $sessionAvailability.ComplianceConnected = $true
                    if ($complianceConnectionStatus -eq 'not attempted') {
                        $sessionAvailability.ComplianceConnectionStatus = 'connected'
                    }
                }
            }
        }
    }
    catch {
        Write-Debug "Connection inspection skipped: $($_.Exception.Message)"
    }

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

    $mfaEnabledMemberPercentValue = Get-ObjectPropertyValue -InputObject $flags -Name 'MfaRegisteredEnabledMemberPercent' -Default 'not collected'
    $mfaEnabledMemberPercent = if (Test-ValueCollected -Value $mfaEnabledMemberPercentValue) { Get-DoubleValue -Value $mfaEnabledMemberPercentValue -Default 0 } else { $null }
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

    if ($null -ne $mfaEnabledMemberPercent -and $mfaEnabledMemberPercent -lt 80) {
        $topRisks.Add('MFA registration is below 80% for enabled members; identity posture is insufficient for broad AI rollout.') | Out-Null
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
    elseif ($null -eq $anyoneLinkSites -and (Get-ObjectPropertyValue -InputObject $spo -Name 'AnyoneLinksPossible' -Default $false) -eq $true) {
        $topRisks.Add('Tenant sharing allows anyone links, and per-site anyone-link counts were not collected.') | Out-Null
        $quickWins.Add('Collect a SharePoint site inventory and restrict anyone-link sharing where the tenant still allows it.') | Out-Null
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
        [ordered]@{ Section = 'SharePointOneDrive'; Action = { Resolve-SharePointGraphSettingsFallback -Section (Get-SharePointOneDriveReadiness -Sample:$IncludeSampling -Max $SampleSize) } }
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

    # ReadinessFlags runs last - it summarises the sections above.
    Write-Phase 'Summary'
    Write-Step 'Generating readiness summary flags...' -Color Cyan
    # Run summary builders through Invoke-Collector so a summary failure never blocks the JSON export.
    $script:CollectorStepTotal = $script:CollectorStepIndex + 4
    Invoke-Collector -Section 'ReadinessFlags' -Action { Get-ReadinessFlags }
    Invoke-Collector -Section 'ReadinessEvidence' -Action { Get-ReadinessEvidence }
    Invoke-Collector -Section 'Recommendations' -Action { Get-Recommendations }
    Invoke-Collector -Section 'PrerequisitesAndGaps' -Action { Get-PrerequisitesAndGaps }
    Write-Progress -Id 1 -Activity 'Microsoft 365 AI Readiness Export' -Completed

    # ========================================================================
    # OUTPUT & SUMMARY
    # ========================================================================
    Write-Phase 'Output'
    $reportWritten = $false
    if ($PSCmdlet.ShouldProcess($resolvedOutputPath, 'Write AI readiness JSON report')) {
        $outputDirectory = Split-Path -Path $resolvedOutputPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory)) {
            [void][System.IO.Directory]::CreateDirectory($outputDirectory)
        }

        Write-Step "Writing JSON export to '$resolvedOutputPath'..." -Color Cyan
        $exportReport = [ordered]@{}
        foreach ($key in @($script:Report.Keys)) {
            if (-not $IncludeDiagnostics -and $key -eq 'CollectorTimings') {
                continue
            }

            $exportReport[$key] = $script:Report[$key]
        }

        $exportReport | ConvertTo-Json -Depth 15 | Out-File -LiteralPath $resolvedOutputPath -Encoding utf8NoBOM -ErrorAction Stop
        $reportWritten = $true
        Write-Step "AI readiness export complete: $resolvedOutputPath" -Color Green
    }
    else {
        Write-Step "Report file was not written: $resolvedOutputPath" -Color DarkYellow
    }

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
        Write-Information -MessageData '' -Tags @('AIReadiness') -InformationAction Continue
        Write-Step 'Top collector durations (ms):' -Color Cyan
        $timingTable = ($timingRows | Format-Table -AutoSize | Out-String).TrimEnd()
        Write-Information -MessageData $timingTable -Tags @('AIReadiness') -InformationAction Continue
    }

    $overallScoreSummary = Get-ObjectPropertyValue -InputObject $script:Report.ReadinessEvidence -Name 'OverallScore' -Default 'not collected'
    $overallReadinessSummary = Get-ObjectPropertyValue -InputObject $script:Report.Recommendations -Name 'OverallReadiness' -Default 'unknown'
    if ($script:Report.ReadinessEvidence) {
        Write-Information -MessageData '' -Tags @('AIReadiness') -InformationAction Continue
        Write-Step ("Overall AI readiness score: {0}" -f $overallScoreSummary) -Color Cyan
    }
    $reportTopRisks = @(Get-ObjectPropertyValue -InputObject $script:Report.Recommendations -Name 'TopRisks' -Default @())
    if (@($reportTopRisks).Count -gt 0) {
        Write-Step 'Top risk:' -Color Yellow
        Write-Step (" - {0}" -f $reportTopRisks[0]) -Color Yellow
    }
    $reportQuickWins = @(Get-ObjectPropertyValue -InputObject $script:Report.Recommendations -Name 'QuickWins' -Default @())
    if (@($reportQuickWins).Count -gt 0) {
        Write-Step 'Top quick win:' -Color Green
        Write-Step (" - {0}" -f $reportQuickWins[0]) -Color Green
    }

    $evaluatableSectionKeys = @($script:Report.Keys | Where-Object { $_ -notin @('CollectorTimings', 'Errors') })
    $processedSectionCount = 0
    $partialSectionCount = 0
    $skippedSectionCount = 0
    foreach ($sectionKey in $evaluatableSectionKeys) {
        $sectionValue = $script:Report[$sectionKey]
        if (Test-SectionCollected -SectionData $sectionValue) {
            $processedSectionCount++
            $sectionStatus = [string](Get-ObjectPropertyValue -InputObject $sectionValue -Name 'Status' -Default 'reported')
            if ($sectionStatus -eq 'partial') {
                $partialSectionCount++
            }
        }
        else {
            $skippedSectionCount++
        }
    }
    $failedCollectorCount = $script:Report.Errors.Count
    $script:ExitCode = Get-ExportExitCode -FailedCollectorCount $failedCollectorCount -ProcessedSectionCount $processedSectionCount -FailOnCollectorError:([bool]$FailOnCollectorError)

    Write-Step ("Summary - Processed: {0} (Partial: {1}) | Skipped/NotCollected: {2} | Errors: {3}" -f $processedSectionCount, $partialSectionCount, $skippedSectionCount, $failedCollectorCount) -Color Cyan
    Write-Step 'Run complete. Reminder: tenant settings were read-only throughout this export.' -Color Green

    Write-Output ([PSCustomObject][ordered]@{
        OutputPath        = $resolvedOutputPath
        ReportWritten     = $reportWritten
        OverallScore      = $overallScoreSummary
        OverallReadiness  = $overallReadinessSummary
        ProcessedSections = $processedSectionCount
        PartialSections   = $partialSectionCount
        SkippedSections   = $skippedSectionCount
        FailedCollectors  = $failedCollectorCount
        ExitCode          = $script:ExitCode
    })
}
catch {
    $script:ExitCode = 1
    Write-Progress -Id 1 -Activity 'Microsoft 365 AI Readiness Export' -Completed
    Write-Error "Fatal: $($_.Exception.Message)"
}
finally {
    # ========================================================================
    # CLEANUP
    # ========================================================================
    Disconnect-Services
}

$global:LASTEXITCODE = $script:ExitCode
if ($MyInvocation.InvocationName -ne '.') {
    exit $script:ExitCode
}