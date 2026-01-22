# ============================================================================
# o365-secure-score-extract.ps1
#
# Extracts Microsoft 365 Secure Score, control profiles, Conditional Access policies,
# Security Defaults, and MFA registration summary for a given tenant. Outputs to JSON.
#
# Detailed comments and debug output added for clarity and troubleshooting.
# ============================================================================
#
# Full documentation - https://github.com/directorcia/Office365/wiki/Extract-Microsoft-365-Secure-Score-information

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantDomain, # The primary domain of the target Microsoft 365 tenant
    [string]$DataFile = "", # Optional: Output file path for JSON data
    [switch]$Compact # Optional: Also output a compact summary file for AI/analysis
)

# Script extracts Microsoft 365 security posture data (Secure Score, controls, CA policies, MFA adoption)
# and saves to a JSON file for analysis or integration with other tools

$ErrorActionPreference = "Stop" # Stop on all errors
$ProgressPreference = "SilentlyContinue" # Suppress progress bars for cleaner output


# Utility functions for consistent output
function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Write-Err($msg) { Write-Host "[ERROR] $msg" -ForegroundColor Red }


# Helper to fetch all pages from a Graph API endpoint, handling paging and errors
function Invoke-GraphCollection {
    param(
        [string]$Uri, # The Graph API endpoint to call
        [switch]$UseMgGraph # Use Microsoft Graph PowerShell SDK
    )

    $results = @()
    $next = $Uri
    $pageCount = 0

    Write-Debug "[Invoke-GraphCollection] Starting collection for URI: $Uri"
    Write-Host "  Calling Graph API: $Uri" -ForegroundColor Gray

    while ($next) {
        try {
            $pageCount++
            $pageStart = Get-Date
            Write-Debug "[Invoke-GraphCollection] Fetching page $pageCount from $next"
            if ($UseMgGraph) {
                $resp = Invoke-MgGraphRequest -Uri $next -Method GET -ErrorAction Stop
            } else {
                Write-Err "Graph request requires Connect-MgGraph connection"
                return @()
            }
            $pageElapsed = ((Get-Date) - $pageStart).TotalSeconds
            $itemCount = if ($resp.value) { $resp.value.Count } else { 0 }
            if ($resp.value) { $results += $resp.value }
            Write-Host "  ✓ Page $pageCount completed: $itemCount items ($([math]::Round($pageElapsed, 1))s) | Total: $($results.Count) items" -ForegroundColor Green
            $next = $resp.'@odata.nextLink'
            if ($next) {
                Write-Host "  → Fetching page $($pageCount + 1)..." -ForegroundColor Cyan
            }
        } catch {
            Write-Err "Graph call failed: $($_.Exception.Message)"
            Write-Debug "[Invoke-GraphCollection] Exception: $($_.Exception | Out-String)"
            Write-Host ""
            Write-Host "=== Graph API Error Details ===" -ForegroundColor Red
            
            # Try to parse error details from the error object
            if ($_.ErrorDetails.Message) {
                try {
                    $errorObj = $_.ErrorDetails.Message | ConvertFrom-Json
                    if ($errorObj.error) {
                        Write-Err "Error Code: $($errorObj.error.code)"
                        Write-Err "Error Message: $($errorObj.error.message)"
                        
                        # Specific guidance based on error code
                        if ($errorObj.error.code -eq "Authorization_RequestDenied" -or $errorObj.error.code -eq "Forbidden") {
                            Write-Host ""
                            Write-Host "PERMISSION ISSUE DETECTED:" -ForegroundColor Yellow
                            Write-Host "  Run: Connect-MgGraph -Scopes SecurityEvents.Read.All,Policy.Read.All,Reports.Read.All,Directory.Read.All" -ForegroundColor Cyan
                            Write-Host ""
                        }
                    }
                } catch {
                    Write-Err "Error Details: $($_.ErrorDetails.Message)"
                }
            }
            
            # Method 2: Try to read response stream if error details are not present
            if (-not $_.ErrorDetails.Message) {
                try {
                    $result = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($result)
                    $responseBody = $reader.ReadToEnd()
                    if ($responseBody) {
                        try {
                            $errorObj = $responseBody | ConvertFrom-Json
                            if ($errorObj.error) {
                                Write-Err "Error Code: $($errorObj.error.code)"
                                Write-Err "Error Message: $($errorObj.error.message)"
                            }
                        } catch {
                            Write-Err "Response: $responseBody"
                        }
                    }
                } catch {
                    Write-Warn "Could not read error response body"
                }
            }
            
            Write-Host "================================" -ForegroundColor Red
            Write-Host ""
            
            # Log the URI that failed for troubleshooting
            Write-Err "Failed URI: $next"
            break
        }
    }

    return $results
}

function Get-SecureScoreData {
    # Retrieves the latest Secure Score and history for the tenant
    Write-Debug "[Get-SecureScoreData] Collecting Secure Score data"
    $scores = Invoke-GraphCollection -Uri "https://graph.microsoft.com/beta/security/secureScores?`$top=5&`$orderby=createdDateTime%20desc" -UseMgGraph
    $latest = $scores | Sort-Object createdDateTime -Descending | Select-Object -First 1
    return [pscustomobject]@{
        Latest = $latest
        History = $scores
    }
}

function Get-SecureScoreControls {
    # Retrieves all Secure Score controls and highlights open/important ones
    Write-Debug "[Get-SecureScoreControls] Collecting Secure Score control profiles"
    $controls = Invoke-GraphCollection -Uri "https://graph.microsoft.com/beta/security/secureScoreControlProfiles?`$top=200" -UseMgGraph
    # Flag likely gaps so they can be prioritized
    $openControls = $controls | Where-Object { $_.controlStateUpdates.state -ne "completed" -and $_.tier -ne "informational" }
    return [pscustomobject]@{
        All = $controls
        Open = $openControls
        TopOpen = $openControls | Sort-Object rank | Select-Object -First 25
    }
}


function Get-ConditionalAccessPolicies {
    # Retrieves Conditional Access policies and strips verbose fields
    Write-Debug "[Get-ConditionalAccessPolicies] Collecting Conditional Access policies"
    $policies = Invoke-GraphCollection -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?`$top=100" -UseMgGraph
    $sanitized = $policies | ForEach-Object {
        [pscustomobject]@{
            DisplayName   = $_.displayName
            State         = $_.state
            CreatedDate   = $_.createdDateTime
            ModifiedDate  = $_.modifiedDateTime
            Conditions    = $_.conditions
            GrantControls = $_.grantControls
            SessionControls = $_.sessionControls
            Id            = $_.id
        }
    }
    return $sanitized
}


function Get-SecurityDefaultsStatus {
    # Checks if Security Defaults are enabled for the tenant
    Write-Debug "[Get-SecurityDefaultsStatus] Checking security defaults status"
    try {
        $policy = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy" -Method GET -ErrorAction Stop
        return $policy
    } catch {
        Write-Warn "Could not retrieve security defaults status: $($_.Exception.Message)"
        return $null
    }
}


function Get-MfaRegistrationSummary {
    # Retrieves MFA registration summary for the last 30 days
    Write-Debug "[Get-MfaRegistrationSummary] Collecting MFA registration summary"
    try {
        # API requires a period argument, otherwise returns 400
        $summary = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationSummary(period='D30')" -Method GET -ErrorAction Stop
        return $summary.value
     } catch {
        Write-Warn "Could not retrieve MFA registration summary: $($_.Exception.Message)"
        return @()
    }
}


function Get-TenantInfoFromGraph {
    # (Unused) Helper to get tenant info from Graph
    param([string]$GraphToken)
    if ([string]::IsNullOrWhiteSpace($GraphToken)) {
        # Using MgGraph context instead
        try {
            $org = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/organization?`$top=1" -Method GET -ErrorAction SilentlyContinue
            if ($org.value -and $org.value.Count -gt 0) {
                return [pscustomobject]@{
                    OrgDisplayName = $org.value[0].displayName
                    TenantId = $org.value[0].id
                }
            }
        } catch {}
        return $null
    }
    return $null
}


function Convert-SecureScoreDataSummary {
    param([object]$SecurityData)
    Write-Debug "[Convert-SecureScoreDataSummary] Summarizing security data for compact output"
    # Summarize to reduce payload size from MB to KB for AI processing
    $summarized = [pscustomobject]@{
        Tenant = $SecurityData.Tenant
        TenantId = $SecurityData.TenantId
        CollectionDate = $SecurityData.CollectionDate
        CollectedBy = $SecurityData.CollectedBy
        # Keep only latest secure score
        SecureScore = @{
            Latest = $SecurityData.SecureScore.Latest
            HistorySummary = @{
                Count = @($SecurityData.SecureScore.History).Count
                OldestDate = ($SecurityData.SecureScore.History | Sort-Object createdDateTime | Select-Object -First 1).createdDateTime
                NewestDate = ($SecurityData.SecureScore.History | Sort-Object createdDateTime -Descending | Select-Object -First 1).createdDateTime
            }
        }
        # Summarize controls - keep only open/important ones
        SecureScoreControls = @{
            TotalCount = @($SecurityData.SecureScoreControls.All).Count
            OpenCount = @($SecurityData.SecureScoreControls.Open).Count
            TopOpen = $SecurityData.SecureScoreControls.TopOpen | Select-Object -Property id, title, rank, implementationCost, userImpact, threats, tier, remediation, controlCategory, actionUrl, maxScore, score -First 25
        }
        # Keep CA policies but remove verbose internal fields
        ConditionalAccess = $SecurityData.ConditionalAccess | Select-Object -Property DisplayName, State, CreatedDate, ModifiedDate, Conditions, GrantControls, SessionControls -First 50
        # Keep security defaults
        SecurityDefaults = $SecurityData.SecurityDefaults
        # Keep MFA summary
        MfaRegistrationSummary = $SecurityData.MfaRegistrationSummary
    }
    return $summarized
}


# ===================== MAIN EXECUTION =====================
# Banner
Write-Host "" 
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  M365 Secure Score Data Extraction" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""


# Validate required parameter
if (-not $TenantDomain) { 
    Write-Err "Parameter -TenantDomain is required"
    exit 1 
}


# Auto-generate data file path if not specified
if (-not $DataFile) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $sanitizedDomain = $TenantDomain -replace '[^a-zA-Z0-9]', '-'
    $parentDir = Split-Path -Parent (Get-Location).Path
    if (-not $parentDir) { $parentDir = Split-Path -Parent $PSScriptRoot }
    $DataFile = Join-Path $parentDir "${sanitizedDomain}_ss_${timestamp}.json"
    Write-Debug "[Main] Auto-generated data file path: $DataFile"
}


Write-Info "Target Tenant: $TenantDomain"
Write-Info "Output File: $DataFile"
Write-Host ""


# Check if Microsoft.Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Err "Microsoft.Graph.Authentication module not found"
    Write-Info "Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    exit 1
}


# Connect to Microsoft Graph with required permissions
try {
    $requiredScopes = @(
        "SecurityEvents.Read.All",
        "Policy.Read.All",
        "Reports.Read.All",
        "Directory.Read.All"
    )
    # Check if already connected in this session
    $context = Get-MgContext -ErrorAction SilentlyContinue
    $needsConnection = $true
    if ($context -and $context.TenantId) {
        Write-Info "Microsoft Graph session detected in current PowerShell session"
        Write-Host "  Current Account: $($context.Account)" -ForegroundColor Gray
        # Try to get tenant domain to compare
        try {
            $org = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/organization" -Method GET -ErrorAction Stop
            $currentDomain = $org.value[0].verifiedDomains | Where-Object { $_.isDefault -eq $true } | Select-Object -ExpandProperty name
            if ($currentDomain) {
                Write-Host "  Current Domain: $currentDomain" -ForegroundColor Gray
                # Compare with requested tenant
                if ($currentDomain -eq $TenantDomain) {
                    Write-Host "  ✓ Already connected to target tenant" -ForegroundColor Green
                    $needsConnection = $false
                } else {
                    Write-Warn "Currently connected to different tenant: $currentDomain"
                    Write-Info "Reconnecting to: $TenantDomain"
                    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                }
            }
        } catch {
            Write-Host "  (Could not verify tenant domain - will reconnect)" -ForegroundColor DarkGray
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    } else {
        Write-Host "  No active Graph session in this PowerShell session" -ForegroundColor Gray
    }
    if ($needsConnection) {
        Write-Info "Connecting to Microsoft Graph..."
        Write-Debug "[Main] Connecting to Microsoft Graph with scopes: $($requiredScopes -join ', ')"
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
        $context = Get-MgContext
    }
    Write-Host ""
    Write-Host "✓ CONNECTED TO TENANT" -ForegroundColor Green -BackgroundColor DarkGreen
    Write-Host "  Account: $($context.Account)" -ForegroundColor White
    Write-Host ""
} catch {
    Write-Err "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    Write-Debug "[Main] Exception during Connect-MgGraph: $($_.Exception | Out-String)"
    exit 1
}


# Test Graph connection health before starting
Write-Host "" 
Write-Host "Testing Microsoft Graph connectivity..." -ForegroundColor Cyan
try {
    $testStart = Get-Date
    $null = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/organization" -Method GET -ErrorAction Stop
    $testElapsed = ((Get-Date) - $testStart).TotalMilliseconds
    Write-Host "  ✓ Graph API responding in $([math]::Round($testElapsed, 0))ms" -ForegroundColor Green
    if ($testElapsed -gt 5000) {
        Write-Host "  ⚠️  Slow response time detected - data collection may take longer" -ForegroundColor Yellow
    }
} catch {
    Write-Err "Graph API connectivity test failed: $($_.Exception.Message)"
    Write-Debug "[Main] Exception during Graph connectivity test: $($_.Exception | Out-String)"
    throw
}


# ===================== DATA COLLECTION STEPS =====================

# 1. Secure Score
Write-Host "" 
Write-Host "[1/5] Collecting Secure Score data..." -ForegroundColor Cyan
Write-Host "      This may take 10-30 seconds depending on tenant size" -ForegroundColor Gray
$startTime = Get-Date
Write-Debug "[Main] Starting Secure Score data collection"
$secureScoreData = Get-SecureScoreData
$elapsed = ((Get-Date) - $startTime).TotalSeconds
Write-Host "      ✓ Completed in $([math]::Round($elapsed, 1)) seconds" -ForegroundColor Green

# 2. Secure Score Controls
Write-Host "" 
Write-Host "[2/5] Collecting Secure Score control profiles..." -ForegroundColor Cyan
Write-Host "      Retrieving up to 200 security controls" -ForegroundColor Gray
$startTime = Get-Date
Write-Debug "[Main] Starting Secure Score control profile collection"
$controlProfiles = Get-SecureScoreControls
$elapsed = ((Get-Date) - $startTime).TotalSeconds
Write-Host "      ✓ Retrieved $(@($controlProfiles.All).Count) controls in $([math]::Round($elapsed, 1)) seconds" -ForegroundColor Green

# 3. Conditional Access Policies
Write-Host "" 
Write-Host "[3/5] Collecting Conditional Access policies..." -ForegroundColor Cyan
$startTime = Get-Date
Write-Debug "[Main] Starting Conditional Access policy collection"
$caPolicies = Get-ConditionalAccessPolicies
$elapsed = ((Get-Date) - $startTime).TotalSeconds
$policyCount = if ($caPolicies) { @($caPolicies).Count } else { 0 }
Write-Host "      ✓ Retrieved $policyCount policies in $([math]::Round($elapsed, 1)) seconds" -ForegroundColor Green

# 4. Security Defaults
Write-Host "" 
Write-Host "[4/5] Checking security defaults status..." -ForegroundColor Cyan
$startTime = Get-Date
Write-Debug "[Main] Starting Security Defaults status check"
$securityDefaults = Get-SecurityDefaultsStatus
$elapsed = ((Get-Date) - $startTime).TotalSeconds
Write-Host "      ✓ Completed in $([math]::Round($elapsed, 1)) seconds" -ForegroundColor Green

# 5. MFA Registration Summary
Write-Host "" 
Write-Host "[5/5] Collecting MFA registration summary..." -ForegroundColor Cyan
$startTime = Get-Date
Write-Debug "[Main] Starting MFA registration summary collection"
$mfaSummary = Get-MfaRegistrationSummary
$elapsed = ((Get-Date) - $startTime).TotalSeconds
Write-Host "      ✓ Completed in $([math]::Round($elapsed, 1)) seconds" -ForegroundColor Green
Write-Host ""


# ===================== OUTPUT & FINALIZATION =====================

Write-Host ""
Write-Info "Building security data object..."
$securityData = [pscustomobject]@{
    Tenant                 = $TenantDomain
    CollectionDate         = (Get-Date).ToString('o')
    TenantId               = $context.TenantId
    CollectedBy            = $context.Account
    SecureScore            = $secureScoreData
    SecureScoreControls    = $controlProfiles
    ConditionalAccess      = $caPolicies
    SecurityDefaults       = $securityDefaults
    MfaRegistrationSummary = $mfaSummary
}
Write-Debug "[Main] Security data object built"

# Save full data file
Write-Info "Saving full data to: $DataFile"
$securityData | ConvertTo-Json -Depth 10 | Out-File -FilePath $DataFile -Encoding UTF8
$fileSize = (Get-Item $DataFile).Length
$fileSizeKB = [math]::Round($fileSize / 1KB, 2)
Write-Debug "[Main] Full data file saved: $DataFile ($fileSizeKB KB)"

Write-Host ""
Write-Host "✓ Full data file saved" -ForegroundColor Green
Write-Host "  File:        $DataFile" -ForegroundColor Gray
Write-Host "  Size:        $fileSizeKB KB" -ForegroundColor Gray
Write-Host ""

# Save compact version if requested
$compactFilePath = ""
if ($Compact) {
    Write-Info "Creating compact/summarized data for AI processing..."
    try {
        $summarizedData = Summarize-SecureScoreData -SecurityData $securityData
        # Generate compact file path
        $compactFile = $DataFile -replace '\.json$', '_compact.json'
        Write-Info "Saving compact data to: $compactFile"
        $summarizedData | ConvertTo-Json -Depth 10 | Out-File -FilePath $compactFile -Encoding UTF8 -Force
        $compactSize = (Get-Item $compactFile).Length
        $compactSizeKB = [math]::Round($compactSize / 1KB, 2)
        $reductionPercent = [math]::Round((1 - ($compactSize / $fileSize)) * 100, 1)
        Write-Debug "[Main] Compact data file saved: $compactFile ($compactSizeKB KB, $reductionPercent% smaller)"
        Write-Host ""
        Write-Host "✓ Compact data file saved" -ForegroundColor Green
        Write-Host "  File:        $compactFile" -ForegroundColor Gray
        Write-Host "  Size:        $compactSizeKB KB" -ForegroundColor Gray
        Write-Host "  Reduction:   $reductionPercent% smaller" -ForegroundColor Cyan
        Write-Host ""
        $compactFilePath = $compactFile
    } catch {
        Write-Err "Failed to create compact file: $($_.Exception.Message)"
        Write-Debug "[Main] Exception during compact file creation: $($_.Exception | Out-String)"
        Write-Host "  Full data file is still available" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  ✓ DATA EXTRACTION COMPLETE" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  Source Tenant:   $($securityData.Tenant)" -ForegroundColor Cyan
Write-Host "  Tenant ID:       $($securityData.TenantId)" -ForegroundColor Cyan
Write-Host "  Collected:       $($securityData.CollectionDate)" -ForegroundColor Cyan
Write-Host ""
Write-Host "  FULL DATA FILE (all details):" -ForegroundColor Yellow
Write-Host "    File: $DataFile" -ForegroundColor Gray
Write-Host "    Size: $fileSizeKB KB" -ForegroundColor Gray
if ($compactFilePath) {
    Write-Host ""
    Write-Host "  COMPACT DATA FILE (AI-optimized):" -ForegroundColor Yellow
    Write-Host "    File: $compactFilePath" -ForegroundColor Gray
    Write-Host "    Size: $compactSizeKB KB" -ForegroundColor Gray
    Write-Host "    Use this for AI systems with file size limits" -ForegroundColor Green
}
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""

Write-Info "Data extraction completed successfully!"
Write-Host "  Use the full file for detailed analysis and integration" -ForegroundColor Gray
if ($compactFilePath) {
    Write-Host "  Use the compact file for AI analysis and systems with upload limits" -ForegroundColor Gray
}
Write-Host ""
