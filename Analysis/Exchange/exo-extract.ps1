#region Parameters
<#
USAGE:
1. Open a PowerShell terminal.
2. Run: .\exo-extract.ps1 -tenantdomain <yourdomain>
   The script will automatically connect to Exchange Online if needed.
   
Documentation - https://github.com/directorcia/Office365/wiki/Extract-Exchange-Online-information
PARAMETERS:
   -TenantDomain <string>       : (Mandatory) Tenant domain to connect to (e.g., contoso.onmicrosoft.com)
   -OutputFolder <string>       : (Optional) Output folder for exported files. Defaults to parent of script folder.
   -Compact                     : (Optional) If set, creates an additional ultra-compact summary (single-page) optimized for instant AI analysis.
   -Credential <PSCredential>   : (Optional) PSCredential object for non-interactive authentication. If not provided, interactive authentication will be used.
   -ConnectionRetries <int>     : (Optional) Number of times to retry Exchange Online connection (default: 3).
   -MaxRetries <int>            : (Optional) Maximum number of retries for data collection commands (default: 3).
   -SkipConnection              : (Optional) Skip connection attempt if already connected interactively in the same session.
   -JsonDepth <int>             : (Optional) JSON serialization depth to prevent truncation (default: 64).
EXAMPLES:
   # Auto-connect (interactive)
   .\exo-extract.ps1 -TenantDomain contoso.onmicrosoft.com
   
   # With credentials (non-interactive)
   $cred = Get-Credential
   .\exo-extract.ps1 -TenantDomain contoso.onmicrosoft.com -Credential $cred
   
   # Create additional ultra-compact summary
   .\exo-extract.ps1 -TenantDomain contoso.onmicrosoft.com -Compact
   
   # Skip connection if already connected
   .\exo-extract.ps1 -TenantDomain contoso.onmicrosoft.com -SkipConnection

OUTPUT FILES:
   - exo_summary_*.json          - Complete detailed configuration (all data)
    - exo_summary_*_compact.json  - AI-optimized compact summary (smaller models, ~256KB)
   - exo_summary_*_ultra-compact.json - Ultra-compact single-page summary (instant analysis, ~5KB)
#>
param(
    [Parameter(Mandatory)]
    [string]$TenantDomain,
    [string]$OutputFolder,
    [switch]$Compact,
    [int]$MaxRetries = 3,
    [switch]$SkipConnection,
    [PSCredential]$Credential,
    [int]$ConnectionRetries = 3,
    [ValidateRange(4, 200)]
    [int]$JsonDepth = 64
)
#endregion
#region Script Defaults & Globals
if (-not $MaxRetries) { $MaxRetries = 3 }
if (-not $global:ErrorCount) { $global:ErrorCount = 0 }
if (-not $global:WarningCount) { $global:WarningCount = 0 }
if (-not $global:ErrorList) { $global:ErrorList = @() }
if (-not $global:WarningList) { $global:WarningList = @() }
if (-not $global:CategoryTimings) { $global:CategoryTimings = @{} }
if (-not $global:ExportMetadata) { $global:ExportMetadata = @{} }
#endregion
#region Helper Functions
function Invoke-WithRetry {
    <#
    .SYNOPSIS
        Executes a script block with retry logic for transient errors.
    .PARAMETER Script
        The script block to execute.
    .PARAMETER MaxAttempts
        Maximum number of attempts.
    .PARAMETER Category
        Category name for debug/error output.
    .PARAMETER Critical
        If set, treat failures as critical.
    #>
    param(
        [Parameter(Mandatory)] [scriptblock]$Script,
        [int]$MaxAttempts = 3,
        [string]$Category = "",
        [switch]$Critical
    )
    $attempt = 0
    $delays = @(2,4,6,8,10)
    $lastError = $null
    while ($attempt -lt $MaxAttempts) {
        try {
            $attempt++
            Write-Debug "[Invoke-WithRetry] Attempt $attempt for $Category."
            return & $Script
        } catch {
            $lastError = $_.Exception.Message
            $retryable = ($lastError -match 'timeout|temporarily|rate limit|network|transient|unavailable|throttl|server busy|try again')
            if ($retryable -and $attempt -lt $MaxAttempts) {
                Write-Warn "${Category}: Attempt $attempt failed (retryable) - $lastError"
                Start-Sleep -Seconds $delays[$attempt-1]
            } else {
                Write-Err "${Category}: Attempt $attempt failed (non-retryable) - $lastError"
                break
            }
        }
    }
    if ($Critical) {
        Write-Err "CRITICAL: ${Category} collection failed after $MaxAttempts attempts."
    } else {
        Write-Warn "NON-CRITICAL: ${Category} collection failed after $MaxAttempts attempts."
    }
    return $null
}

function Write-Warn {
    <#
    .SYNOPSIS
        Writes a warning message and tracks global warning count.
    .PARAMETER Message
        The warning message to display.
    #>
    param([Parameter(Mandatory)][string]$Message)
    Write-Warning $Message
    $global:WarningCount++
    $global:WarningList += $Message
    Write-Debug "[Write-Warn] $Message"
}

function Write-Info {
    <#
    .SYNOPSIS
        Writes an informational message.
    .PARAMETER Message
        The info message to display.
    #>
    param([Parameter(Mandatory)][string]$Message)
    Write-Host $Message -ForegroundColor Cyan
    Write-Debug "[Write-Info] $Message"
}

function Write-Err {
    <#
    .SYNOPSIS
        Writes an error message and tracks global error count.
    .PARAMETER Message
        The error message to display.
    #>
    param([Parameter(Mandatory)][string]$Message)
    Write-Error $Message
    $global:ErrorCount++
    $global:ErrorList += $Message
    Write-Debug "[Write-Err] $Message"
}

function Write-Stat {
    <#
    .SYNOPSIS
        Writes a statistics message.
    .PARAMETER Message
        The statistics message to display.
    #>
    param([Parameter(Mandatory)][string]$Message)
    Write-Host $Message -ForegroundColor Magenta
    Write-Debug "[Write-Stat] $Message"
}
#endregion

#Requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
    Extracts Exchange Online configuration, security, and compliance settings to JSON files.
.DESCRIPTION
    Connects to Exchange Online, collects configuration, security, and compliance data, and exports to JSON files for analysis or archiving. 
    Generates both detailed and AI-optimized compact summaries by default. Includes error handling, retry logic, progress reporting, and metadata.
.PARAMETER TenantDomain
    The tenant domain to connect to (e.g., contoso.onmicrosoft.com).
.PARAMETER OutputFolder
    Optional. Output folder for exported files. Defaults to parent of script folder.
.PARAMETER Compact
    Optional. If set, creates an additional ultra-compact summary (single-page) optimized for instant AI analysis.
.PARAMETER Credential
    Optional. PSCredential object for non-interactive authentication. If not provided, interactive authentication will be used.
.PARAMETER ConnectionRetries
    Optional. Number of times to retry Exchange Online connection (default: 3).
.PARAMETER MaxRetries
    Optional. Maximum number of retries for data collection commands (default: 3).
.PARAMETER SkipConnection
    Optional. Skip connection attempt if already connected interactively in the same session.
.PARAMETER JsonDepth
    Optional. JSON serialization depth to prevent truncation (default: 64).
#>
#>

# Helper: Format time span for display
function Format-TimeSpan {
    param([TimeSpan]$ts)
    if ($ts.TotalSeconds -lt 60) { return "{0:N1}s" -f $ts.TotalSeconds }
    elseif ($ts.TotalMinutes -lt 60) { return "{0:N1}m" -f $ts.TotalMinutes }
    else { return "{0:N1}h" -f $ts.TotalHours }
}

#region Output Directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ($OutputFolder) {
    $OutputDir = $OutputFolder
} else {
    $OutputDir = Split-Path -Parent $scriptDir
}
if (-not (Test-Path $OutputDir)) {
    Write-Info "Creating output directory: $OutputDir"
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$sanitizedDomain = $TenantDomain -replace '[^a-zA-Z0-9]', '-'
Write-Info "Output directory: $OutputDir"
#endregion
# Helper: Get safe count for any object
function Get-SafeCount {
    param($obj)
    if ($null -eq $obj) { return 0 }
    elseif ($obj -is [System.Collections.IEnumerable] -and $obj.GetType().Name -ne 'String') { return @($obj).Count }
    else { return 1 }
}



function Save-Json {
    param(
        [Parameter(Mandatory)] $Data,
        [Parameter(Mandatory)] $FilePath,
        [string]$Category = "",
        [ValidateRange(4, 200)]
        [int]$JsonDepth = 64
    )
    $isEmpty = $false
    $dataType = $null
    if ($null -eq $Data -or $Data -eq $false) {
        $isEmpty = $true
        $dataType = if ($null -eq $Data) { "Null" } else { $Data.GetType().Name }
    } elseif ($Data -is [System.Collections.IEnumerable] -and $Data.GetType().Name -ne 'String') {
        $dataType = $Data.GetType().Name
        try {
            if ($Data.Count -eq 0) { $isEmpty = $true }
        } catch {
            if (-not ($Data | Measure-Object).Count) { $isEmpty = $true }
        }
    } else {
        $dataType = $Data.GetType().Name
    }
    Write-Debug "[Save-Json] DataType: $dataType, IsEmpty: $isEmpty, JsonDepth: $JsonDepth, Value: $Data"
    if ($isEmpty) {
        Write-Warn "No data to save for ${FilePath}. Skipping file."
        return
    }
    # Ensure .json extension
    $jsonFilePath = $FilePath
    if (-not $jsonFilePath.ToLower().EndsWith('.json')) {
        $jsonFilePath = "$jsonFilePath.json"
    }
    try {
        $Data | ConvertTo-Json -Depth $JsonDepth | Out-File -FilePath $jsonFilePath -Encoding UTF8
        $size = (Get-Item $jsonFilePath).Length
        $sizeStr = if ($size -gt 1MB) { "{0:N2} MB" -f ($size/1MB) } elseif ($size -gt 1KB) { "{0:N1} KB" -f ($size/1KB) } else { "$size bytes" }
        Write-Info "Saved: $jsonFilePath ($sizeStr)"
        if ($Category) { $global:ExportMetadata[$Category] = @{ File = $jsonFilePath; Size = $sizeStr } }
    } catch {
        Write-Err "Failed to save ${jsonFilePath}: $($_.Exception.Message)"
    }
}
## Output directory logic already handled above




#region Extraction

#region Connect to Exchange Online
# Ensure ExchangeOnlineManagement module is available
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Err "ExchangeOnlineManagement module is not installed. Please install it with 'Install-Module ExchangeOnlineManagement'."
    throw
}
Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

function Test-ExchangeOnlineConnection {
    <#
    .SYNOPSIS
        Tests if a valid Exchange Online connection exists.
    .DESCRIPTION
        Attempts to verify an active Exchange Online connection by running a simple command.
    .OUTPUTS
        [bool] $true if connected, $false otherwise
    #>
    try {
        Write-Debug "[Test-ExchangeOnlineConnection] Checking for active Exchange Online session..."
        $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue

        if ($null -eq $connInfo -or @($connInfo).Count -eq 0) {
            Write-Debug "[Test-ExchangeOnlineConnection] No active connection information found."
            return $false
        }

        # Additional validation: attempt a simple query (force stop on failure)
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Debug "[Test-ExchangeOnlineConnection] Connection is valid."
        return $true
    } catch {
        Write-Debug "[Test-ExchangeOnlineConnection] Connection test failed: $($_.Exception.Message)"
        return $false
    }
}

function Connect-ExchangeOnlineWithRetry {
    <#
    .SYNOPSIS
        Attempts to connect to Exchange Online with retry logic.
    .PARAMETER Credential
        Optional PSCredential object for authentication.
    .PARAMETER MaxAttempts
        Maximum number of connection attempts (default: 3).
    .OUTPUTS
        [bool] $true if connection successful, $false otherwise
    #>
    param(
        [PSCredential]$Credential,
        [int]$MaxAttempts = 3
    )
    
    $attempt = 0
    $delays = @(2, 5, 10)
    
    while ($attempt -lt $MaxAttempts) {
        try {
            $attempt++
            Write-Info "Attempting to connect to Exchange Online (Attempt $attempt of $MaxAttempts)..."
            
            $connectParams = @{
                ShowBanner = $false
                ErrorAction = 'Stop'
            }
            
            if ($Credential) {
                Write-Debug "[Connect-ExchangeOnlineWithRetry] Using provided credentials."
                $connectParams['Credential'] = $Credential
            } else {
                Write-Debug "[Connect-ExchangeOnlineWithRetry] Using interactive authentication."
            }
            
            Connect-ExchangeOnline @connectParams
            
            # Verify connection
            Start-Sleep -Seconds 2
            if (Test-ExchangeOnlineConnection) {
                Write-Info "Successfully connected to Exchange Online."
                return $true
            } else {
                throw "Connection verification failed."
            }
        } catch {
            $lastError = $_.Exception.Message
            if ($attempt -lt $MaxAttempts) {
                Write-Warn "Connection attempt $attempt failed: $lastError. Retrying in $($delays[$attempt-1]) seconds..."
                Start-Sleep -Seconds $delays[$attempt - 1]
            } else {
                Write-Err "Failed to connect to Exchange Online after $MaxAttempts attempts: $lastError"
            }
        }
    }
    
    return $false
}

# Connection Logic
$isConnected = Test-ExchangeOnlineConnection

if ($SkipConnection -and -not $isConnected) {
    Write-Err "No active Exchange Online session found and -SkipConnection was specified. Please run Connect-ExchangeOnline before running this script with -SkipConnection. Aborting extraction."
    exit 1
} elseif (-not $isConnected) {
    Write-Info "No active Exchange Online session detected. Establishing connection..."
    if (-not (Connect-ExchangeOnlineWithRetry -Credential $Credential -MaxAttempts $ConnectionRetries)) {
        Write-Err "Unable to establish Exchange Online connection. Aborting extraction."
        exit 1
    }
} else {
    Write-Info "Active Exchange Online session detected. Proceeding with extraction."
}
#endregion

#region Pre-Collection Summary
$psver = $PSVersionTable.PSVersion.ToString()
$user = $env:USERNAME
$modver = (Get-Module ExchangeOnlineManagement | Select-Object -First 1).Version
Write-Host "" 
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host " EXCHANGE ONLINE CONFIGURATION EXTRACTION " -BackgroundColor Cyan -ForegroundColor Black
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "" 
Write-Host "Tenant: $TenantDomain" -ForegroundColor Cyan
Write-Host "User: $user" -ForegroundColor Cyan
Write-Host "PowerShell: $psver" -ForegroundColor Cyan
Write-Host "ExchangeOnlineManagement: $modver" -ForegroundColor Cyan
Write-Host "Output Directory: $OutputDir" -ForegroundColor Cyan
Write-Host "MaxRetries: $MaxRetries" -ForegroundColor Cyan
Write-Host "" 
Write-Host "Categories to collect:"
foreach ($cat in @(
    "Organization config", "Mailboxes", "Mailbox permissions", "Transport rules", "Retention policies", "Retention policy tags", "Mobile device policies", "Inbound connectors", "Outbound connectors", "Accepted domains", "Remote domains", "Journaling rules", "Anti-spam policies", "Anti-malware policies", "Safe Links policies", "Safe Attachments policies", "Sharing policies", "Email address policies", "OWA policies", "Anti-phishing policies", "ATP policies", "Distribution groups", "Unified groups"
)) { Write-Host "  - $cat" -ForegroundColor Yellow }
Write-Host "" 
#endregion

$categories = @(
    @{ Name = "Organization config";      Cmd = { Get-OrganizationConfig }; Desc = "General organization settings."; Critical = $true },
    @{ Name = "Mailboxes";                Cmd = { Get-Mailbox -ResultSize Unlimited | Select-Object * }; Desc = "Full mailbox details."; Critical = $true },
    @{ Name = "Mailbox permissions";      Cmd = { Get-Mailbox -ResultSize Unlimited | ForEach-Object { Get-MailboxPermission -Identity $_.Identity } }; Desc = "Mailbox permissions."; Critical = $true },
    @{ Name = "Transport rules";          Cmd = { Get-TransportRule }; Desc = "Mail flow rules."; Critical = $true },
    @{ Name = "Retention policies";       Cmd = { Get-RetentionPolicy }; Desc = "Retention policies."; Critical = $false },
    @{ Name = "Retention policy tags";    Cmd = { Get-RetentionPolicyTag }; Desc = "Retention policy tags."; Critical = $false },
    @{ Name = "Mobile device policies";   Cmd = { Get-MobileDeviceMailboxPolicy }; Desc = "Mobile device mailbox policies."; Critical = $false },
    @{ Name = "Inbound connectors";       Cmd = { Get-InboundConnector }; Desc = "Inbound mail connectors."; Critical = $false },
    @{ Name = "Outbound connectors";      Cmd = { Get-OutboundConnector }; Desc = "Outbound mail connectors."; Critical = $false },
    @{ Name = "Accepted domains";         Cmd = { Get-AcceptedDomain }; Desc = "Accepted domains."; Critical = $true },
    @{ Name = "Remote domains";           Cmd = { Get-RemoteDomain }; Desc = "Remote domains."; Critical = $false },
    @{ Name = "Journaling rules";         Cmd = { Get-JournalRule }; Desc = "Journaling rules."; Critical = $false },
    @{ Name = "Anti-spam policies";       Cmd = { Get-HostedContentFilterPolicy }; Desc = "Anti-spam policies."; Critical = $true },
    @{ Name = "Anti-malware policies";    Cmd = { Get-MalwareFilterPolicy }; Desc = "Anti-malware policies."; Critical = $true },
    @{ Name = "Safe Links policies";      Cmd = { Get-SafeLinksPolicy }; Desc = "Defender for Office 365 Safe Links."; Critical = $false },
    @{ Name = "Safe Attachments policies";Cmd = { Get-SafeAttachmentPolicy }; Desc = "Defender for Office 365 Safe Attachments."; Critical = $false },
    @{ Name = "Sharing policies";         Cmd = { Get-SharingPolicy }; Desc = "Sharing policies."; Critical = $false },
    @{ Name = "Email address policies";   Cmd = { Get-EmailAddressPolicy }; Desc = "Email address policies."; Critical = $false },
    @{ Name = "OWA policies";             Cmd = { Get-OwaMailboxPolicy }; Desc = "OWA mailbox policies."; Critical = $false },
    @{ Name = "Anti-phishing policies";   Cmd = { Get-AntiPhishPolicy }; Desc = "Anti-phishing policies."; Critical = $true },
    @{ Name = "ATP policies";             Cmd = { Get-AtpPolicyForO365 }; Desc = "ATP (Advanced Threat Protection) policies for O365."; Critical = $true },
    @{ Name = "Distribution groups";      Cmd = { Get-DistributionGroup | ForEach-Object { $group = $_; [PSCustomObject]@{ Group = $group; Members = Get-DistributionGroupMember -Identity $group.Identity } } }; Desc = "Distribution groups with members."; Critical = $false },
    @{ Name = "Unified groups";           Cmd = { Get-UnifiedGroup | ForEach-Object { $group = $_; [PSCustomObject]@{ Group = $group; Members = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Members } } }; Desc = "Unified groups (Microsoft 365 Groups) with members."; Critical = $false }
)

$results = @{}
$total = $categories.Count
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()
$categoryStats = @{}
$results = @{}
$total = $categories.Count
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()
$categoryStats = @{}
for ($i = 0; $i -lt $total; $i++) {
    $cat = $categories[$i]
    $percent = [math]::Round((($i+1)/$total)*100,1)
    $msg = "[$($percent)%] Collecting $($cat.Name)... ($($i+1)/$total)"
    Write-Progress -Activity "Exchange Online Config Export" -Status $msg -PercentComplete $percent
    Write-Host "────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "→ $($cat.Name): $($cat.Desc)" -ForegroundColor Yellow
    $swCat = [System.Diagnostics.Stopwatch]::StartNew()
    $data = Invoke-WithRetry -Script $cat.Cmd -MaxAttempts $MaxRetries -Category $cat.Name -Critical:($cat.Critical)
    $swCat.Stop()
    $duration = $swCat.Elapsed
    $count = Get-SafeCount $data
    $success = ($null -ne $data)
    $results[$cat.Name] = $data
    $global:CategoryTimings[$cat.Name] = Format-TimeSpan $duration
    $categoryStats[$cat.Name] = @{ Count = $count; Duration = $duration.TotalSeconds; Success = $success; Critical = $cat.Critical }
    if ($success -and ($count -gt 0 -or $cat.Name -ne 'Transport rules')) {
        Write-Host ("  ✓ $($cat.Name): {0} item(s) collected in {1}" -f $count, (Format-TimeSpan $duration)) -ForegroundColor Green
    } elseif ($cat.Name -eq 'Transport rules' -and $count -eq 0) {
        Write-Info "  $($cat.Name): No transport rules found. Treated as informational only."
    } else {
        if ($cat.Critical) {
            Write-Host ("  ✗ $($cat.Name): FAILED (CRITICAL) in {0}" -f (Format-TimeSpan $duration)) -ForegroundColor Red
        } else {
            Write-Host ("  ✗ $($cat.Name): FAILED (non-critical) in {0}" -f (Format-TimeSpan $duration)) -ForegroundColor Yellow
        }
    }
}
$swTotal.Stop()
Write-Progress -Activity "Exchange Online Config Export" -Completed -Status "All categories collected."
Write-Host "" 
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  ✓ ALL CONFIGURATION CATEGORIES COLLECTED" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
#endregion
#endregion


#region Assign Results
#region Assign Results
$orgConfig          = $results["Organization config"]
$mailboxes          = $results["Mailboxes"]
$mailboxPermissions = $results["Mailbox permissions"]
$transportRules     = $results["Transport rules"]
$retentionPolicies  = $results["Retention policies"]
$retentionTags      = $results["Retention policy tags"]
$mobilePolicies     = $results["Mobile device policies"]
$inboundConnectors  = $results["Inbound connectors"]
$outboundConnectors = $results["Outbound connectors"]
$acceptedDomains    = $results["Accepted domains"]
$remoteDomains      = $results["Remote domains"]
$journalRules       = $results["Journaling rules"]
$antiSpam           = $results["Anti-spam policies"]
$antiMalware        = $results["Anti-malware policies"]
$safeLinks          = $results["Safe Links policies"]
$safeAttachments    = $results["Safe Attachments policies"]
$sharingPolicies    = $results["Sharing policies"]
$emailAddrPolicies  = $results["Email address policies"]
$owaPolicies        = $results["OWA policies"]
#endregion


#region Summary Output
Write-Info "Bundling all configuration data into a single summary JSON file..."
$summary = [pscustomobject]@{
    OrganizationConfig = $orgConfig
    Mailboxes = $mailboxes # All properties, all items
    MailboxPermissions = $mailboxPermissions # All properties, all items
    TransportRules = $transportRules # All properties, all items
    RetentionPolicies = $retentionPolicies # All properties, all items
    RetentionPolicyTags = $retentionTags # All properties, all items
    MobileDevicePolicies = $mobilePolicies # All properties, all items
    InboundConnectors = $inboundConnectors # All properties, all items
    OutboundConnectors = $outboundConnectors # All properties, all items
    AcceptedDomains = $acceptedDomains # All properties, all items
    RemoteDomains = $remoteDomains # All properties, all items
    JournalRules = $journalRules # All properties, all items
    AntiSpamPolicies = $antiSpam # All properties, all items
    AntiMalwarePolicies = $antiMalware # All properties, all items
    SafeLinksPolicies = $safeLinks # All properties, all items
    SafeAttachmentsPolicies = $safeAttachments # All properties, all items
    SharingPolicies = $sharingPolicies # All properties, all items
    EmailAddressPolicies = $emailAddrPolicies # All properties, all items
    OWAPolicies = $owaPolicies # All properties, all items
    AntiPhishingPolicies = $results['Anti-phishing policies'] # All properties, all items
    ATPPolicies = $results['ATP policies'] # All properties, all items
    DistributionGroups = $results['Distribution groups'] # All properties, all items
    UnifiedGroups = $results['Unified groups'] # All properties, all items
    ErrorCount = $global:ErrorCount
    WarningCount = $global:WarningCount
    Timings = $global:CategoryTimings
    ExportMetadata = $global:ExportMetadata
    CollectionDate = (Get-Date).ToString('o')
    CollectedBy = $TenantDomain
    ScriptVersion = "1.1"
}
$summaryFile = Join-Path $OutputDir ("exo_summary_{0}_{1}.json" -f $sanitizedDomain, $timestamp)
Save-Json -Data $summary -FilePath $summaryFile -Category "Summary" -JsonDepth $JsonDepth
#endregion


#region Compact Output (AI-Optimized)
Write-Info "Creating compact/AI-optimized summary JSON file..."
$compactSummary = [pscustomobject]@{
    Metadata = [pscustomobject]@{
        TenantDomain = $TenantDomain
        CollectionDate = (Get-Date).ToString('o')
        ScriptVersion = "1.2"
        ErrorCount = $global:ErrorCount
        WarningCount = $global:WarningCount
        TotalCategoriesCollected = $total
        TotalExecutionTime = Format-TimeSpan $swTotal.Elapsed
    }
    OrganizationConfig = $orgConfig | Select-Object Name,DomainName,IsDehydrated,IsMultiGeo,DefaultPublicFolderMailbox,Languages,PublicFoldersEnabled,AddressBookPolicyRoutingEnabled
    MailboxesSummary = [pscustomobject]@{
        TotalCount = (Get-SafeCount $mailboxes)
        Sample = ($mailboxes | Select-Object -First 200 DisplayName,PrimarySmtpAddress,RecipientTypeDetails,ForwardingAddress,AuditEnabled,UserPrincipalName,DeliverToMailboxAndForward,RetentionPolicy,RetentionHoldEnabled,MailboxPlan,WhenCreated,WhenChanged)
    }
    MailboxPermissionsSummary = [pscustomobject]@{
        TotalCount = (Get-SafeCount $mailboxPermissions)
        Sample = ($mailboxPermissions | Select-Object -First 400 Identity,User,AccessRights,IsInherited)
    }
    TransportRules = $transportRules | Select-Object -First 100 Name,Enabled,Priority,Mode,Comments,From,SentTo,SubjectContainsWords,Actions
    RetentionPolicies = $retentionPolicies | Select-Object -First 50 Name,IsDefault,Comment,RetentionId,RetentionPolicyTagLinks
    RetentionPolicyTags = $retentionTags | Select-Object -First 50 Name,Type,RetentionEnabled,RetentionAction,AgeLimitForRetention,IsVisible
    MobileDevicePolicies = $mobilePolicies | Select-Object -First 25 Name,AllowNonProvisionableDevices,PasswordEnabled,DeviceEncryptionEnabled,AllowSimplePassword,MinPasswordLength
    InboundConnectors = $inboundConnectors | Select-Object -First 25 Name,ConnectorType,Enabled,SenderDomains,RequireTls,CloudServicesMailEnabled
    OutboundConnectors = $outboundConnectors | Select-Object -First 25 Name,ConnectorType,Enabled,RecipientDomains,SmartHost,TlsSettings
    AcceptedDomains = $acceptedDomains | Select-Object -First 25 Name,DomainType,Default,DomainName,MatchSubdomains
    RemoteDomains = $remoteDomains | Select-Object -First 25 Name,AllowedOOFType,AutoReplyEnabled,ContentType,TrustedMailInboundEnabled
    JournalRules = $journalRules | Select-Object -First 25 Name,Enabled,JournalEmailAddress,Recipient,Scope
    AntiSpamPolicies = $antiSpam | Select-Object -First 25 Name,IsEnabled,SpamAction,QuarantineTag,BypassInboundMessages,BypassOutboundMessages
    AntiMalwarePolicies = $antiMalware | Select-Object -First 25 Name,IsEnabled,Action,QuarantineTag,BypassInboundMessages,BypassOutboundMessages
    SafeLinksPolicies = $safeLinks | Select-Object -First 25 Name,IsEnabled,ScanUrls,TrackClicks,IsEnabledForInternalSenders
    SafeAttachmentsPolicies = $safeAttachments | Select-Object -First 25 Name,IsEnabled,Action,RedirectUrl,IsEnabledForInternalSenders
    SharingPolicies = $sharingPolicies | Select-Object -First 25 Name,Domains,Enabled,DefaultSharingPolicy
    EmailAddressPolicies = $emailAddrPolicies | Select-Object -First 25 Name,Enabled,RecipientFilter,Priority,EnabledEmailAddressTemplates
    OWAPolicies = $owaPolicies | Select-Object -First 25 Name,IsDefault,InstantMessagingType,DefaultTheme,LogonPagePublicPrivateSelectionEnabled
    AntiPhishingPolicies = $results['Anti-phishing policies'] | Select-Object -First 25 Name,Enabled,Action,IsDefault,AuthenticationMethods
    ATPPolicies = $results['ATP policies'] | Select-Object -First 25 Name,IsEnabled,Action,RedirectUrl,IsDefault
    DistributionGroupsSummary = [pscustomobject]@{
        TotalCount = (Get-SafeCount $results['Distribution groups'])
        Sample = ($results['Distribution groups'] | Select-Object -First 75 @{Name='GroupName';Expression={$_.Group.DisplayName}},@{Name='MemberCount';Expression={($_.Members | Measure-Object).Count}},@{Name='Members';Expression={($_.Members | Select-Object -First 10 DisplayName,PrimarySmtpAddress)}})
    }
    UnifiedGroupsSummary = [pscustomobject]@{
        TotalCount = (Get-SafeCount $results['Unified groups'])
        Sample = ($results['Unified groups'] | Select-Object -First 75 @{Name='GroupName';Expression={$_.Group.DisplayName}},@{Name='MemberCount';Expression={($_.Members | Measure-Object).Count}},@{Name='Members';Expression={($_.Members | Select-Object -First 10 DisplayName,PrimarySmtpAddress)}})
    }
    Timings = $global:CategoryTimings
}

$compactFile = Join-Path $OutputDir ("exo_summary_{0}_{1}_compact.json" -f $sanitizedDomain, $timestamp)
Save-Json -Data $compactSummary -FilePath $compactFile -Category "CompactSummary" -JsonDepth $JsonDepth
Write-Info "Compact/AI-optimized summary saved: $compactFile (optimized for smaller AI models)"

if ($Compact) {
    Write-Info "Additional compact mode requested (-Compact flag detected). Creating extra condensed summary..."
    $ultraCompactSummary = [pscustomobject]@{
        TenantDomain = $TenantDomain
        CollectionDate = (Get-Date).ToString('o')
        ScriptVersion = "1.2"
        Status = @{
            ErrorCount = $global:ErrorCount
            WarningCount = $global:WarningCount
            SuccessfulCategories = @($categoryStats.Keys | Where-Object { $categoryStats[$_].Success }).Count
        }
        QuickStats = @{
            TotalMailboxes = (Get-SafeCount $mailboxes)
            TotalTransportRules = (Get-SafeCount $transportRules)
            TotalAcceptedDomains = (Get-SafeCount $acceptedDomains)
            TotalDistributionGroups = (Get-SafeCount $results['Distribution groups'])
            TotalUnifiedGroups = (Get-SafeCount $results['Unified groups'])
        }
        SecurityPolicies = @{
            AntiSpam = (Get-SafeCount $antiSpam)
            AntiMalware = (Get-SafeCount $antiMalware)
            SafeLinks = (Get-SafeCount $safeLinks)
            SafeAttachments = (Get-SafeCount $safeAttachments)
            AntiPhishing = (Get-SafeCount $results['Anti-phishing policies'])
        }
        CriticalCategories = @{
            OrganizationConfigured = ($null -ne $orgConfig)
            MailboxesCollected = ($null -ne $mailboxes)
            DomainsConfigured = (Get-SafeCount $acceptedDomains)
            SecurityPoliciesActive = @($antiSpam, $antiMalware, $results['Anti-phishing policies'], $results['ATP policies'] | Where-Object { $null -ne $_ }).Count
        }
    }
    $ultraCompactFile = Join-Path $OutputDir ("exo_summary_{0}_{1}_ultra-compact.json" -f $sanitizedDomain, $timestamp)
    Save-Json -Data $ultraCompactSummary -FilePath $ultraCompactFile -Category "UltraCompactSummary" -JsonDepth $JsonDepth
    Write-Info "Ultra-compact summary saved: $ultraCompactFile (single-page summary for instant AI analysis)"
}
#endregion



#region Final Report
Write-Host "" 
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  ✓ EXCHANGE ONLINE CONFIG EXTRACTION COMPLETE" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  Source Tenant:   $TenantDomain" -ForegroundColor Cyan
Write-Host "  Output Directory: $OutputDir" -ForegroundColor Cyan
Write-Host "  Collected:       $((Get-Date).ToString('o'))" -ForegroundColor Cyan
Write-Host ""
Write-Host "GENERATED FILES:" -ForegroundColor Yellow
Write-Host "  ✓ Full Summary:          exo_summary_${sanitizedDomain}_${timestamp}.json" -ForegroundColor Green
Write-Host "    (Complete detailed data - use for full analysis/archiving)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  ✓ Compact Summary:       exo_summary_${sanitizedDomain}_${timestamp}_compact.json" -ForegroundColor Green
Write-Host "    (AI-optimized for smaller models - ~256KB, richer context)" -ForegroundColor DarkGray
if ($Compact) {
    Write-Host ""
    Write-Host "  ✓ Ultra-Compact:         exo_summary_${sanitizedDomain}_${timestamp}_ultra-compact.json" -ForegroundColor Green
    Write-Host "    (Single-page summary - ~5KB, instant analysis)" -ForegroundColor DarkGray
}
Write-Host ""
Write-Stat "  Total Categories Collected: $total"

# ERRORS & WARNINGS FILTERING (define before use)
$uniqueErrors = $global:ErrorList | Select-Object -Unique
$actionableErrors = $uniqueErrors | Where-Object {
    $_ -notmatch 'CRITICAL: .+ collection failed after \d+ attempts\.' -and
    $_ -notmatch 'NON-CRITICAL: .+ collection failed after \d+ attempts\.' -and
    $_ -ne 'Errors encountered during extraction:'
}
$genericErrors = $uniqueErrors | Where-Object {
    $_ -match 'CRITICAL: .+ collection failed after \d+ attempts\.' -or
    $_ -match 'NON-CRITICAL: .+ collection failed after \d+ attempts\.'
}
$uniqueWarnings = $global:WarningList | Select-Object -Unique
$actionableWarnings = $uniqueWarnings | Where-Object {
    $_ -notmatch 'NON-CRITICAL: .+ collection failed after \d+ attempts\.' -and
    $_ -ne 'Warnings encountered during extraction:'
}
$genericWarnings = $uniqueWarnings | Where-Object {
    $_ -match 'NON-CRITICAL: .+ collection failed after \d+ attempts\.'
}

# ERRORS SECTION
Write-Host ""  # Spacer
Write-Host "ERRORS (Actionable: $($actionableErrors.Count))" -ForegroundColor Red
Write-Host "---------------------------" -ForegroundColor DarkGray
if ($actionableErrors.Count -gt 0) {
    foreach ($err in $actionableErrors) {
        Write-Host "  $err" -ForegroundColor Red
    }
} else {
    Write-Host "  (No actionable errors. All failures were generic collection errors.)" -ForegroundColor DarkGray
}
Write-Debug "[ErrorList] Count: $($global:ErrorList.Count), Unique: $($uniqueErrors.Count), Actionable: $($actionableErrors.Count), Generic: $($genericErrors.Count)"
Write-Debug "[ErrorList] Contents: $($actionableErrors -join '; ')"
Write-Debug "[GenericErrorList] Contents: $($genericErrors -join '; ')"

# WARNINGS SECTION
Write-Host ""  # Spacer
Write-Host "WARNINGS (Actionable: $($actionableWarnings.Count))" -ForegroundColor Yellow
Write-Host "------------------------------" -ForegroundColor DarkGray
if ($actionableWarnings.Count -gt 0) {
    foreach ($warn in $actionableWarnings) {
        Write-Host "  $warn" -ForegroundColor Yellow
    }
} else {
    Write-Host "  (No actionable warnings. All failures were generic collection warnings.)" -ForegroundColor DarkGray
}
Write-Debug "[WarningList] Count: $($global:WarningList.Count), Unique: $($uniqueWarnings.Count), Actionable: $($actionableWarnings.Count), Generic: $($genericWarnings.Count)"
Write-Debug "[WarningList] Contents: $($actionableWarnings -join '; ')"
Write-Debug "[GenericWarningList] Contents: $($genericWarnings -join '; ')"

# TOTAL TIME
Write-Host ""  # Spacer
Write-Stat "  Total Time: $(Format-TimeSpan $swTotal.Elapsed)"
Write-Host ""
#endregion
#endregion