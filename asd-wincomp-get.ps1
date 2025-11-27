<#
.SYNOPSIS
    Check Intune Windows Compliance Policy settings against ASD Blueprint requirements

.DESCRIPTION
    This script checks Intune Windows 10/11 Compliance Policies against the ASD Blueprint baseline
    defined in JSON (default sourced from GitHub). It compares each policy's settings and
    reports PASS/FAIL per setting. Generates an HTML report and optional CSV export.

    Baseline example (windows-compliance.json):
    {
      "@odata.type": "#microsoft.graph.windows10CompliancePolicy",
      "passwordRequired": true,
      "passwordBlockSimple": true,
      "passwordMinimumLength": 15,
      ...
    }

.EXAMPLE
    .\asd-wincomp-get.ps1

    Connects to Microsoft Graph, downloads the latest Windows compliance baseline from GitHub, 
    checks settings, and generates an HTML report in the parent directory.

.EXAMPLE
    .\asd-wincomp-get.ps1 -ExportToCSV

    Also exports results to CSV in the parent directory.

.EXAMPLE
    .\asd-wincomp-get.ps1 -BaselinePath "C:\Baselines\windows-compliance.json"

    Uses a custom baseline JSON file.

.NOTES
    Author: CIAOPS
    Date: 11-18-2025
    Version: 1.0

    Requirements:
    - Microsoft.Graph.DeviceManagement PowerShell module
    - Permissions: DeviceManagementConfiguration.Read.All or Global Reader
    - Internet connection when using GitHub baseline

    Default Baseline:
    https://raw.githubusercontent.com/directorcia/bp/main/Intune/Policies/ASD/windows-compliance.json

.LINK
    https://github.com/directorcia/office365
    https://github.com/directorcia/office365/wiki/Windows-Compliance-Policy-Check - Documentation
    https://blueprint.asd.gov.au/configuration/intune/device-compliance/
#>

[CmdletBinding()]
param(
    [switch]$ExportToCSV,
    [string]$CSVPath,
    [Parameter(HelpMessage = "Path or URL to baseline JSON file. Defaults to GitHub URL for latest ASD Blueprint settings")]
    [string]$BaselinePath,
    [Parameter(HelpMessage = "Enable detailed logging to file")]
    [switch]$DetailedLogging,
    [Parameter(HelpMessage = "Path to log file. Defaults to parent directory with timestamp")]
    [string]$LogPath,
    [Parameter(HelpMessage = "Custom output path for HTML compliance report. Defaults to timestamped file in parent directory.")]
    [string]$HTMLPath,
    [Parameter(HelpMessage = "Target a specific compliance policy by display name. If not specified, all Windows compliance policies are checked.")]
    [string]$PolicyName
)

# Paths
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$parentPath = Split-Path -Parent $scriptPath

if (-not $CSVPath) { $CSVPath = Join-Path $parentPath "asd-wincomp-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv" }
if ($DetailedLogging -and -not $LogPath) { $LogPath = Join-Path $parentPath "asd-wincomp-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').log" }

# Default GitHub URL for baseline settings
$defaultGitHubURL = "https://raw.githubusercontent.com/directorcia/bp/main/Intune/Policies/ASD/windows-compliance.json"
if (-not $BaselinePath) { $BaselinePath = $defaultGitHubURL }

# Script-scope state
$script:BaselinePath = $BaselinePath
$script:baselineLoaded = $false
$script:HTMLPath = if ($HTMLPath) {
    # If relative path provided, resolve to parent directory
    if ([IO.Path]::IsPathRooted($HTMLPath)) { $HTMLPath } else { Join-Path $parentPath $HTMLPath }
} else { Join-Path $parentPath "asd-wincomp-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').html" }
$script:LogPath = $LogPath
$script:DetailedLogging = $DetailedLogging

$scriptVersion = "1.0"
$scriptName = "ASD Windows Compliance Policy Check"

# Logging
function Write-Log {
    param([string]$Message,[string]$Level = "INFO")
    if ($script:DetailedLogging -and $script:LogPath) {
        try { $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"; Add-Content -Path $script:LogPath -Value "[$ts] [$Level] $Message" -ErrorAction Stop } catch { }
    }
}

# Console color output
function Write-ColorOutput {
    param([string]$Message,[string]$Type = "Info")
    $level = switch($Type){"Success"{"INFO"}"Warning"{"WARN"}"Error"{"ERROR"}default{"INFO"}}
    Write-Log -Message $Message -Level $level
    switch($Type){
        "Success" { Write-Host $Message -ForegroundColor Green }
        "Warning" { Write-Host $Message -ForegroundColor Yellow }
        "Error"   { Write-Host $Message -ForegroundColor Red }
        default    { Write-Host $Message -ForegroundColor Cyan }
    }
}

# Baseline loader
function Test-BaselineSchema {
    param([object]$Baseline)
    # Expect: Root object with @odata.type and compliance settings properties
    if ($null -eq $Baseline) { return $false }
    if (-not $Baseline.'@odata.type') { return $false }
    if ($Baseline.'@odata.type' -notlike '*windows10CompliancePolicy*') { return $false }
    return $true
}

function Get-BaselineSettings {
    param([string]$Path)
    Write-Log "Loading baseline from: $Path"
    $json = $null
    $isUrl = $Path -match '^https?://'
    try {
        if ($isUrl) {
            Write-ColorOutput "Downloading baseline from GitHub..." -Type Info
            $content = (Invoke-WebRequest -Uri $Path -UseBasicParsing -ErrorAction Stop).Content
            $json = $content | ConvertFrom-Json -ErrorAction Stop
        }
        elseif (Test-Path $Path) {
            Write-ColorOutput "Loading baseline from local file..." -Type Info
            $json = Get-Content -Path $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        }
        else {
            Write-ColorOutput "Baseline not found at path: $Path" -Type Warning
            return $null
        }
    }
    catch {
        Write-ColorOutput "Failed to load/parse baseline JSON: $($_.Exception.Message)" -Type Error
        return $null
    }

    if (-not (Test-BaselineSchema -Baseline $json)) {
        Write-ColorOutput "Baseline JSON schema validation failed. Expecting Windows 10 Compliance Policy JSON." -Type Error
        return $null
    }

    $script:baselineLoaded = $true
    return $json
}

# Graph module & connection
function Test-GraphModule {
    Write-ColorOutput "Checking for Microsoft.Graph modules..." -Type Info
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop | Out-Null
        Write-ColorOutput "Microsoft.Graph.Authentication module loaded." -Type Success
        return $true
    } 
    catch {
        Write-ColorOutput "Failed to load Microsoft.Graph.Authentication module: $($_.Exception.Message)" -Type Error
        Write-ColorOutput "Install with: Install-Module Microsoft.Graph -Scope CurrentUser" -Type Warning
        return $false
    }
}

function Connect-MSGraph {
    Write-ColorOutput "`nChecking Microsoft Graph connection..." -Type Info
    try {
        # Check if already connected
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context) {
            Write-ColorOutput "Already connected to Microsoft Graph." -Type Success
            Write-ColorOutput "Tenant: $($context.TenantId)" -Type Info
            return $true
        }
        
        Write-ColorOutput "Connecting to Microsoft Graph..." -Type Info
        try {
            # Try interactive browser authentication first
            Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All" -NoWelcome -ErrorAction Stop
            Write-ColorOutput "Connected to Microsoft Graph." -Type Success
            return $true
        }
        catch {
            # If interactive fails (localhost binding issue), try device code flow
            if ($_.Exception.Message -like "*HttpListenerException*" -or $_.Exception.Message -like "*localhost*" -or $_.Exception.Message -like "*unable to listen*") {
                Write-ColorOutput "`nInteractive browser authentication failed (localhost binding issue)." -Type Warning
                Write-ColorOutput "Switching to device code authentication..." -Type Info
                Write-ColorOutput "`n============================================================" -Type Info
                Write-ColorOutput "  DEVICE CODE AUTHENTICATION REQUIRED" -Type Info
                Write-ColorOutput "============================================================" -Type Info
                Write-Host ""
                Write-Host "Opening browser to: " -NoNewline
                Write-Host "https://microsoft.com/devicelogin" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "The device code will appear below - copy it and paste into the browser." -ForegroundColor Cyan
                Write-Host "============================================================`n" -ForegroundColor Cyan
                
                # Open the device login page in default browser
                try {
                    Start-Process "https://microsoft.com/devicelogin"
                    Start-Sleep -Seconds 2  # Give browser time to open
                } catch {
                    Write-ColorOutput "Could not open browser automatically." -Type Warning
                }
                
                # Connect with device code - output goes directly to console
                Write-Host ""
                Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All" -UseDeviceAuthentication -ErrorAction Stop | Out-Default
                Write-Host ""
                Write-ColorOutput "Connected to Microsoft Graph via device code." -Type Success
                return $true
            }
            else {
                throw
            }
        }
    }
    catch { 
        Write-ColorOutput "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Type Error
        Write-ColorOutput "`n╔════════════════════════════════════════════════════════════════╗" -Type Warning
        Write-ColorOutput "║                    POSSIBLE SOLUTIONS                          ║" -Type Warning
        Write-ColorOutput "╠════════════════════════════════════════════════════════════════╣" -Type Warning
        Write-ColorOutput "║ 1. Run PowerShell as Administrator (easiest solution)          ║" -Type Info
        Write-ColorOutput "║                                                                ║" -Type Info
        Write-ColorOutput "║ 2. Run this command as Admin, then retry:                      ║" -Type Info
        Write-ColorOutput "║    netsh http add iplisten 127.0.0.1                           ║" -Type Info
        Write-ColorOutput "║                                                                ║" -Type Info
        Write-ColorOutput "║ 3. Pre-connect manually with device code:                      ║" -Type Info
        Write-ColorOutput "║    Connect-MgGraph -Scopes 'DeviceManagementConfiguration.Read.All' -UseDeviceAuthentication" -Type Info
        Write-ColorOutput "║    Then run this script again                                  ║" -Type Info
        Write-ColorOutput "╚════════════════════════════════════════════════════════════════╝" -Type Warning
        return $false 
    }
}

function Test-GraphPermissions {
    Write-ColorOutput "\nValidating Microsoft Graph permissions..." -Type Info
    try { 
        $url = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?`$top=1"
        $null = Invoke-MgGraphRequest -Method GET -Uri $url -ErrorAction Stop
        Write-ColorOutput "Permission validation passed." -Type Success
        return $true 
    }
    catch { 
        Write-ColorOutput "Permission validation failed: $($_.Exception.Message)" -Type Error
        Write-ColorOutput "Required permission: DeviceManagementConfiguration.Read.All" -Type Warning
        
        # Check if this is a permission issue (BadRequest often means insufficient permissions)
        if ($_.Exception.Message -like "*BadRequest*" -or $_.Exception.Message -like "*Forbidden*" -or $_.Exception.Message -like "*Unauthorized*") {
            Write-ColorOutput "\nInsufficient permissions detected. Attempting to reconnect with correct scope..." -Type Warning
            
            # Try to reconnect with the required permission
            $reconnected = Connect-MSGraph -ForceReconnect
            if ($reconnected) {
                Write-ColorOutput "\nRetrying permission validation..." -Type Info
                try {
                    $null = Invoke-MgGraphRequest -Method GET -Uri $url -ErrorAction Stop
                    Write-ColorOutput "Permission validation passed after reconnection." -Type Success
                    return $true
                }
                catch {
                    Write-ColorOutput "Permission validation still failed: $($_.Exception.Message)" -Type Error
                    Write-ColorOutput "\nYou may need to grant admin consent for the app in Azure AD." -Type Warning
                    Write-ColorOutput "Required permission: DeviceManagementConfiguration.Read.All" -Type Warning
                    return $false
                }
            }
        }
        return $false 
    }
}

# Comparison helpers
function Normalize-Value {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return [bool]$Value }
    # Handle arrays (like scheduledActionsForRule, validOperatingSystemBuildRanges)
    if ($Value -is [array]) {
        if ($Value.Count -eq 0) { return "[]" }
        return ($Value | ConvertTo-Json -Compress)
    }
    # Trim and normalize strings; treat "True"/"False" as booleans when possible
    $s = $Value.ToString().Trim()
    if ($s -match '^(?i:true|false)$') { return [System.Convert]::ToBoolean($s) }
    return $s
}

function Compare-Values {
    param([object]$Current,[object]$Required)
    $c = Normalize-Value $Current
    $r = Normalize-Value $Required
    if ($null -eq $r -and $null -eq $c) { return $true }
    if ($null -eq $r) { return $true }
    
    # Special handling for arrays
    if ($r -is [string] -and $r.StartsWith('[') -and $r.EndsWith(']')) {
        if ($c -is [string] -and $c.StartsWith('[') -and $c.EndsWith(']')) {
            return ($c -eq $r)
        }
        return $false
    }
    
    if ($c -is [bool] -and $r -is [bool]) { return ($c -eq $r) }
    # case-insensitive for strings
    return ("$c" -ieq "$r")
}

function Test-Setting {
    param(
        [string]$PolicyName,
        [string]$SettingName,
        [object]$CurrentValue,
        [object]$RequiredValue
    )
    $cur = if ($null -eq $CurrentValue) { "Not set" } else { $CurrentValue.ToString() }
    $req = if ($null -eq $RequiredValue) { "Not set" } else { $RequiredValue.ToString() }
    $ok = Compare-Values -Current $CurrentValue -Required $RequiredValue
    Write-Log "Check [$PolicyName] $SettingName - Current: $cur, Required: $req, Status: $(if($ok){'PASS'}else{'FAIL'})" -Level $(if($ok){'INFO'}else{'WARN'})
    [pscustomobject]@{
        Policy       = $PolicyName
        Setting      = $SettingName
        CurrentValue = $cur
        RequiredValue= $req
        Compliant    = $ok
        Status       = if ($ok) { 'PASS' } else { 'FAIL' }
    }
}

# HTML report
function New-HTMLReport {
    param([array]$CheckResults,[string]$OutputPath)
    $total = $CheckResults.Count
    $passed = ($CheckResults | Where-Object { $_.Compliant }).Count
    $failed = $total - $passed
    $pct = if ($total -gt 0) { [math]::Round(($passed/$total)*100,2) } else { 0 }
    $overall = if ($pct -eq 100) { 'COMPLIANT' } else { 'NON-COMPLIANT' }
    $statusColor = if ($pct -eq 100) { '#28a745' } else { '#dc3545' }
    $reportDate = Get-Date -Format "dd MMMM yyyy - HH:mm:ss"

        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ASD Windows Compliance Policy Report</title>
<style>
body{font-family:'Segoe UI',Tahoma,Verdana,sans-serif;background:#f0f2f5;padding:20px}
.container{max-width:1200px;margin:0 auto;background:#fff;border-radius:10px;box-shadow:0 10px 30px rgba(0,0,0,.15);overflow:hidden}
.header{background:linear-gradient(135deg,#1e3c72,#2a5298);color:#fff;padding:30px;text-align:center}
.summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:20px;padding:20px;background:#f8f9fa}
.card{background:#fff;padding:18px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.08);text-align:center;transition:transform .3s ease}
.card:hover{transform:translateY(-4px);box-shadow:0 6px 18px rgba(0,0,0,.12)}
.card .value{font-size:2.2em;font-weight:700}
.card.total .value{color:#007bff}
.card.passed .value{color:#28a745}
.card.failed .value{color:#dc3545}
.card.compliance .value{color:$statusColor}
.results{padding:25px}
table{width:100%;border-collapse:collapse}
thead{background:linear-gradient(135deg,#1e3c72,#2a5298);color:#fff}
th,td{padding:12px 14px;border-bottom:1px solid #e9ecef;text-align:left}
tbody tr:nth-child(even){background:#f8f9fa}
.badge{display:inline-block;padding:4px 10px;border-radius:14px;font-weight:600}
.pass{background:#d4edda;color:#155724;border:1px solid #c3e6cb}
.fail{background:#f8d7da;color:#721c24;border:1px solid #f5c6cb}
.overall{background:$statusColor;color:#fff;text-align:center;padding:24px}
.footer{padding:20px;text-align:center;background:#f8f9fa;color:#6c757d;font-size:.9em;border-top:1px solid #e9ecef}
.footer a{color:#2a5298;text-decoration:none;font-weight:700}
.footer a:hover{text-decoration:underline}
</style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>🛡️ ASD Windows Compliance Policy Report</h1>
      <p>Generated: $reportDate</p>
    </div>
    <div class="summary">
            <div class="card total"><div>Total Checks</div><div class="value">$total</div></div>
            <div class="card passed"><div>Passed</div><div class="value">$passed</div></div>
            <div class="card failed"><div>Failed</div><div class="value">$failed</div></div>
            <div class="card compliance"><div>Compliance</div><div class="value">$pct%</div></div>
    </div>
    <div class="results">
      <table>
        <thead><tr><th>Status</th><th>Policy</th><th>Setting</th><th>Current</th><th>Required</th></tr></thead>
        <tbody>
"@

    foreach ($r in $CheckResults) {
        $cls = if ($r.Compliant) { 'pass' } else { 'fail' }
        $txt = if ($r.Compliant) { 'PASS' } else { 'FAIL' }
        $html += @"
          <tr>
            <td><span class="badge $cls">$txt</span></td>
            <td><strong>$($r.Policy)</strong></td>
            <td>$($r.Setting)</td>
            <td>$($r.CurrentValue)</td>
            <td>$($r.RequiredValue)</td>
          </tr>
"@
    }

        $html += @"
        </tbody>
      </table>
    </div>
        <div class="overall"><h2>Overall Status: $overall</h2>
            <p style="font-size:1.1em;margin-top:8px;">$passed of $total checks passed</p>
        </div>

            <div class="footer">
                <p><strong>Reference:</strong> <a href="https://blueprint.asd.gov.au/configuration/intune/devices/compliance-policies/policies/windows-10-11-compliance-policy/" target="_blank">ASD's Blueprint for Secure Cloud - Windows Device Compliance</a></p>
                <p style="margin-top:10px;"><strong>Security Controls:</strong> <a href="https://github.com/directorcia/bp/wiki/Windows-Compliance-Policy-Settings-%E2%80%90-Security-Rationale" target="_blank">Why These Recommendations Matter</a></p>
            </div>
  </div>
</body>
</html>
"@

    try { $html | Out-File -FilePath $OutputPath -Encoding UTF8 -Force; return $true }
    catch { Write-ColorOutput "Failed to generate HTML report: $($_.Exception.Message)" -Type Error; return $false }
}

# Main check
function Invoke-CompliancePolicyCheck {
    param([psobject]$Requirements, [string]$TargetPolicyName)

    Write-ColorOutput "`n========================================" -Type Info
    Write-ColorOutput "  $scriptName v$scriptVersion" -Type Info
    Write-ColorOutput "  ASD Blueprint Compliance Check" -Type Info
    Write-ColorOutput "========================================`n" -Type Info

    Write-ColorOutput "Retrieving Windows compliance policies from Intune..." -Type Info
    try {
        # Use Graph API directly to avoid module loading conflicts
        $url = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"
        $response = Invoke-MgGraphRequest -Method GET -Uri $url -ErrorAction Stop
        $allPolicies = $response.value
        
        # Handle pagination if needed
        while ($response.'@odata.nextLink') {
            $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink' -ErrorAction Stop
            $allPolicies += $response.value
        }
        
        # Filter for Windows 10 compliance policies
        $policies = $allPolicies | Where-Object { 
            $_.'@odata.type' -eq '#microsoft.graph.windows10CompliancePolicy' 
        }
        
        # Further filter by policy name if specified
        if ($TargetPolicyName) {
            $policies = $policies | Where-Object { $_.displayName -eq $TargetPolicyName }
            if (-not $policies) { 
                Write-ColorOutput "No Windows compliance policy found with name: $TargetPolicyName" -Type Warning 
                return $null
            }
        }
        
        if (-not $policies) { 
            Write-ColorOutput "No Windows 10/11 compliance policies found." -Type Warning 
            return $null
        }
        
        Write-ColorOutput "Found $($policies.Count) Windows compliance $(if($policies.Count -eq 1){'policy'}else{'policies'}) to check." -Type Success
    }
    catch {
        Write-ColorOutput "Failed to retrieve compliance policies: $($_.Exception.Message)" -Type Error
        return $null
    }

    $results = @()

    # Get baseline settings (excluding metadata fields)
    $baselineSettings = ($Requirements | Get-Member -MemberType NoteProperty | 
        Where-Object { $_.Name -notin @('@odata.type', 'displayName', 'description', 'version', 'scheduledActionsForRule', 'validOperatingSystemBuildRanges', 'roleScopeTagIds') } | 
        Select-Object -ExpandProperty Name)

    # Handle conflicting password settings - passwordComplexity and passwordRequiredType are mutually exclusive
    # If both are present in baseline, only check passwordRequiredType (which is what the set script uses)
    if (('passwordRequiredType' -in $baselineSettings) -and ('passwordComplexity' -in $baselineSettings)) {
        Write-ColorOutput "`nNote: Both 'passwordRequiredType' and 'passwordComplexity' detected in baseline." -Type Warning
        Write-ColorOutput "These settings are mutually exclusive. Skipping 'passwordComplexity' check (per Microsoft Graph API limitation)." -Type Warning
        $baselineSettings = $baselineSettings | Where-Object { $_ -ne 'passwordComplexity' }
    }

    foreach ($policy in $policies) {
        $policyName = $policy.displayName
        Write-ColorOutput "`nChecking policy: $policyName" -Type Info
        
        # Get full policy details using Graph API
        try {
            $detailUrl = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/$($policy.id)"
            $policyDetails = Invoke-MgGraphRequest -Method GET -Uri $detailUrl -ErrorAction Stop
        }
        catch {
            Write-ColorOutput "Failed to retrieve details for policy: $policyName" -Type Error
            continue
        }

        # Compare each setting in baseline
        foreach ($setting in $baselineSettings) {
            $required = $Requirements.$setting
            $current = $null
            
            # Try to read property directly
            try { 
                $current = $policyDetails.$setting 
            } catch { 
                $current = $null 
            }
            
            # Handle additional properties that might be in AdditionalProperties
            if ($null -eq $current -and $policyDetails.AdditionalProperties) {
                try {
                    $current = $policyDetails.AdditionalProperties[$setting]
                } catch {
                    $current = $null
                }
            }
            
            $results += Test-Setting -PolicyName $policyName -SettingName $setting -CurrentValue $current -RequiredValue $required
        }
    }

    # Output to console
    Write-ColorOutput "`n========================================" -Type Info
    Write-ColorOutput "  CHECK RESULTS" -Type Info
    Write-ColorOutput "========================================`n" -Type Info
    foreach ($r in $results) {
        $type = if ($r.Compliant) { 'Success' } else { 'Error' }
        $sym = if ($r.Compliant) { '[✓]' } else { '[✗]' }
        Write-ColorOutput "$sym [$($r.Policy)] $($r.Setting)" -Type $type
        Write-Host "    Current : $($r.CurrentValue)"
        Write-Host "    Required: $($r.RequiredValue)"
        Write-Host "    Status  : $($r.Status)"
    }

    $total = $results.Count
    $passed = ($results | Where-Object { $_.Compliant }).Count
    $failed = $total - $passed
    $pct = if ($total -gt 0) { [math]::Round(($passed/$total)*100,2) } else { 0 }
    Write-ColorOutput "========================================" -Type Info
    Write-ColorOutput "  SUMMARY" -Type Info
    Write-ColorOutput "========================================" -Type Info
    Write-Host "Total Checks : $total"
    Write-ColorOutput "Passed       : $passed" -Type Success
    if ($failed -gt 0) { Write-ColorOutput "Failed       : $failed" -Type Error } else { Write-ColorOutput "Failed       : $failed" -Type Success }
    Write-Host "Compliance   : $pct%"
    if ($pct -eq 100) { Write-ColorOutput "`nStatus       : COMPLIANT ✓" -Type Success } else { Write-ColorOutput "`nStatus       : NON-COMPLIANT ✗" -Type Error }
    Write-ColorOutput "========================================`n" -Type Info

    # CSV export
    if ($ExportToCSV) {
        try {
            $results | Select-Object Policy,Setting,CurrentValue,RequiredValue,Status | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8
            Write-ColorOutput "Results exported to: $CSVPath" -Type Success
        } catch { Write-ColorOutput "Failed to export CSV: $($_.Exception.Message)" -Type Error }
    }

    # HTML report
    Write-ColorOutput "Generating HTML report..." -Type Info
    if (New-HTMLReport -CheckResults $results -OutputPath $script:HTMLPath) {
        Write-ColorOutput "HTML report generated: $script:HTMLPath" -Type Success
        try { Start-Process $script:HTMLPath } catch { Write-ColorOutput "Could not open report in browser: $($_.Exception.Message)" -Type Warning }
    }

    return $results
}

# Main
try {
    if ($script:DetailedLogging) {
        Write-Log "=== ASD Windows Compliance Policy Check Started ==="
        Write-Log "Script Version: $scriptVersion"
        Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)"
        Write-Log "Log Path: $script:LogPath"
    }

    Write-ColorOutput "`n========================================" -Type Info
    Write-ColorOutput "  ASD Windows Compliance Policy Check" -Type Info
    Write-ColorOutput "========================================" -Type Info
    $isUrl = $script:BaselinePath -match '^https?://'
    if ($isUrl) { Write-ColorOutput "Baseline: GitHub (latest)" -Type Info } elseif (Test-Path $script:BaselinePath) { Write-ColorOutput "Baseline: Local File (found)" -Type Success } else { Write-ColorOutput "Baseline: Local File (not found)" -Type Warning }
    Write-ColorOutput "Location: $script:BaselinePath" -Type Info
    Write-ColorOutput "Output:   $parentPath`n" -Type Info

    $baseline = Get-BaselineSettings -Path $BaselinePath
    if (-not $baseline) { Write-ColorOutput "Failed to load baseline settings. Cannot proceed." -Type Error; exit 1 }

    if (-not (Test-GraphModule)) { Write-ColorOutput "`nMicrosoft.Graph.Authentication module is required." -Type Error; exit 1 }
    if (-not (Connect-MSGraph)) { Write-ColorOutput "`nFailed to connect to Microsoft Graph." -Type Error; exit 1 }
    if (-not (Test-GraphPermissions)) { Write-ColorOutput "`nInsufficient permissions to read compliance policies." -Type Error; exit 1 }

    $null = Invoke-CompliancePolicyCheck -Requirements $baseline -TargetPolicyName $PolicyName
    Write-ColorOutput "`nScript completed." -Type Success
}
catch {
    Write-Log "SCRIPT EXECUTION FAILED: $($_.Exception.Message)"
    Write-ColorOutput "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -Type Error
    Write-ColorOutput "❌ SCRIPT EXECUTION FAILED" -Type Error
    Write-ColorOutput "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -Type Error
    Write-Host ""
    Write-ColorOutput "Error Message:" -Type Error
    Write-Host "  $($_.Exception.Message)"
    Write-Host ""
    Write-ColorOutput "Error Location:" -Type Warning
    Write-Host "  Line: $($_.InvocationInfo.ScriptLineNumber)"
    Write-Host "  Command: $($_.InvocationInfo.Line.Trim())"
    Write-Host ""
    if ($script:DetailedLogging) { Write-Host ""; Write-ColorOutput "Detailed error log saved to: $script:LogPath" -Type Info }
    exit 1
}
