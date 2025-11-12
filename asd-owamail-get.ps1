<#
.SYNOPSIS
    Check Exchange Online OWA Mailbox Policy settings against ASD Blueprint requirements

.DESCRIPTION
    This script checks Exchange Online OWA Mailbox Policies against the ASD Blueprint baseline
    defined in JSON (default sourced from GitHub). It compares each policy's settings and
    reports PASS/FAIL per setting. Generates an HTML report and optional CSV export.

    Baseline example (owamail.json):
    {
      "OwaMailboxPolicy-Default": {
        "InstantMessagingEnabled": true,
        "TextMessagingEnabled": true,
        "ActiveSyncIntegrationEnabled": false,
        ...
      }
    }

.EXAMPLE
    .\asd-owamail-get.ps1

    Connects to Exchange Online, downloads the latest OWA baseline from GitHub, checks settings,
    and generates an HTML report in the parent directory.

.EXAMPLE
    .\asd-owamail-get.ps1 -ExportToCSV

    Also exports results to CSV in the parent directory.

.EXAMPLE
    .\asd-owamail-get.ps1 -BaselinePath "C:\Baselines\owamail.json"

    Uses a custom baseline JSON file.

.NOTES
    Author: CIAOPS
    Date: 11-12-2025
    Version: 1.0

    Requirements:
    - ExchangeOnlineManagement PowerShell module
    - Permissions: Global Reader or Exchange read permissions (View-Only Organization Management)
    - Internet connection when using GitHub baseline

    Default Baseline:
    https://raw.githubusercontent.com/directorcia/bp/main/ASD/Exchange-Online/Roles/owamail.json

.LINK
    https://github.com/directorcia/office365
    https://github.com/directorcia/Office365/wiki/ASD-OWA-Mailbox-Configuration-Check - Documentation
    https://github.com/directorcia/bp/wiki/Exchange-Online-OWA-Mailbox-Security-Controls - Exchange Online OWA Mailbox Security Controls
    https://blueprint.asd.gov.au/configuration/exchange-online/roles/outlook-web-app-policies/
#>

#Requires -Modules ExchangeOnlineManagement

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
    [string]$HTMLPath
)

# Paths
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$parentPath = Split-Path -Parent $scriptPath

if (-not $CSVPath) { $CSVPath = Join-Path $parentPath "asd-owamail-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv" }
if ($DetailedLogging -and -not $LogPath) { $LogPath = Join-Path $parentPath "asd-owamail-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').log" }

# Default GitHub URL for baseline settings
$defaultGitHubURL = "https://raw.githubusercontent.com/directorcia/bp/main/ASD/Exchange-Online/Roles/owamail.json"
if (-not $BaselinePath) { $BaselinePath = $defaultGitHubURL }

# Script-scope state
$script:BaselinePath = $BaselinePath
$script:baselineLoaded = $false
$script:HTMLPath = if ($HTMLPath) {
    # If relative path provided, resolve to parent directory
    if ([IO.Path]::IsPathRooted($HTMLPath)) { $HTMLPath } else { Join-Path $parentPath $HTMLPath }
} else { Join-Path $parentPath "asd-owamail-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').html" }
$script:LogPath = $LogPath
$script:DetailedLogging = $DetailedLogging

$scriptVersion = "1.0"
$scriptName = "ASD OWA Mailbox Policy Settings Check"

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
    # Expect: Root object where each property name is an OWA policy name and value is an object of settings
    if ($null -eq $Baseline) { return $false }
    $props = $Baseline | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue
    if (-not $props -or $props.Count -eq 0) { return $false }
    foreach ($p in $props) {
        $val = $Baseline.$p
        if (-not ($val -is [psobject])) { return $false }
    }
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
        Write-ColorOutput "Baseline JSON schema validation failed. Expecting: { 'PolicyName': { 'Setting': value, ... } }" -Type Error
        return $null
    }

    $script:baselineLoaded = $true
    return $json
}

# EXO module & connection
function Test-ExchangeModule {
    Write-ColorOutput "Checking for ExchangeOnlineManagement module..." -Type Info
    if (Get-Module -Name ExchangeOnlineManagement) { Write-ColorOutput "Module already loaded." -Type Success; return $true }
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-ColorOutput "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -Type Error
        return $false
    }
    try { Import-Module ExchangeOnlineManagement -ErrorAction Stop; Write-ColorOutput "Module loaded." -Type Success; return $true } catch { Write-ColorOutput "Failed to load module: $($_.Exception.Message)" -Type Error; return $false }
}

function Connect-EXO {
    Write-ColorOutput "`nChecking Exchange Online connection..." -Type Info
    try {
        try { $null = Get-OrganizationConfig -ErrorAction Stop; Write-ColorOutput "Already connected to Exchange Online." -Type Success; return $true } catch {
            Write-ColorOutput "Connecting to Exchange Online..." -Type Info
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            Write-ColorOutput "Connected to Exchange Online." -Type Success
            return $true
        }
    }
    catch { Write-ColorOutput "Failed to connect to Exchange Online: $($_.Exception.Message)" -Type Error; return $false }
}

function Test-ExchangePermissions {
    Write-ColorOutput "`nValidating Exchange Online permissions..." -Type Info
    try { $null = Get-OrganizationConfig -ErrorAction Stop; Write-ColorOutput "Permission validation passed." -Type Success; return $true }
    catch { Write-ColorOutput "Permission validation failed: $($_.Exception.Message)" -Type Error; return $false }
}

# Comparison helpers
function Normalize-Value {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return [bool]$Value }
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
    param([array]$CheckResults,[object]$OrgConfig,[string]$OutputPath)
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
<title>ASD OWA Mailbox Policy Compliance Report</title>
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
/* Centered footer styling for info links */
.footer{padding:20px;text-align:center;background:#f8f9fa;color:#6c757d;font-size:.9em;border-top:1px solid #e9ecef}
.footer a{color:#2a5298;text-decoration:none;font-weight:700}
.footer a:hover{text-decoration:underline}
</style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>🛡️ ASD OWA Mailbox Policy Compliance Report</h1>
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
                <p><strong>Reference:</strong> <a href="https://blueprint.asd.gov.au/configuration/exchange-online/roles/outlook-web-app-policies/" target="_blank">ASD's Blueprint for Secure Cloud - OWA Mailbox Policies</a></p>
                <p style="margin-top:10px;"><strong>Security Controls Explanation:</strong> <a href="https://github.com/directorcia/bp/wiki/Exchange-Online-OWA-Mailbox-Security-Controls" target="_blank">Why These Recommendations Matter</a></p>
            </div>
  </div>
</body>
</html>
"@

    try { $html | Out-File -FilePath $OutputPath -Encoding UTF8 -Force; return $true }
    catch { Write-ColorOutput "Failed to generate HTML report: $($_.Exception.Message)" -Type Error; return $false }
}

# Main check
function Invoke-OwaPolicyCheck {
    param([psobject]$Requirements)

    Write-ColorOutput "`n========================================" -Type Info
    Write-ColorOutput "  $scriptName v$scriptVersion" -Type Info
    Write-ColorOutput "  ASD Blueprint Compliance Check" -Type Info
    Write-ColorOutput "========================================`n" -Type Info

    Write-ColorOutput "Retrieving organization and OWA mailbox policy configuration..." -Type Info
    try {
        $orgConfig = Get-OrganizationConfig -ErrorAction Stop
    } catch { Write-ColorOutput "Failed to get organization config: $($_.Exception.Message)" -Type Warning; $orgConfig = $null }

    try {
        $policies = Get-OwaMailboxPolicy -ErrorAction Stop
        if (-not $policies) { Write-ColorOutput "No OWA mailbox policies found." -Type Warning }
    }
    catch {
        Write-ColorOutput "Failed to retrieve OWA mailbox policies: $($_.Exception.Message)" -Type Error
        return $null
    }

    $results = @()

    # Iterate baseline policies
    $baselinePolicyNames = ($Requirements | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)
    foreach ($policyName in $baselinePolicyNames) {
        $baselinePolicy = $Requirements.$policyName
        $tenantPolicy = $policies | Where-Object { $_.Name -eq $policyName }

        if (-not $tenantPolicy) {
            Write-ColorOutput "Policy not found in tenant: $policyName" -Type Error
            # Create FAIL entries for each expected setting under missing policy
            $settingNames = ($baselinePolicy | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)
            foreach ($s in $settingNames) { $results += Test-Setting -PolicyName $policyName -SettingName $s -CurrentValue $null -RequiredValue $baselinePolicy.$s }
            continue
        }

        # Compare each setting in baseline
        $baselineSettings = ($baselinePolicy | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)
        foreach ($setting in $baselineSettings) {
            $required = $baselinePolicy.$setting
            # Try to read property directly; if not present, mark as null
            $current = $null
            try { $current = $tenantPolicy.$setting } catch { $current = $null }
            # Known alias fallback for Offline access naming variations
            if ($null -eq $current -and $setting -match 'OfflineAccessEnabled') {
                # Try common alternatives
                foreach ($alt in @('OfflineEnabled','AllowOfflineOn','OWAforDevicesOfflineEnabled')) {
                    try { $current = $tenantPolicy.$alt } catch { $current = $null }
                    if ($null -ne $current) { break }
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
    if (New-HTMLReport -CheckResults $results -OrgConfig $orgConfig -OutputPath $script:HTMLPath) {
        Write-ColorOutput "HTML report generated: $script:HTMLPath" -Type Success
        try { Start-Process $script:HTMLPath } catch { Write-ColorOutput "Could not open report in browser: $($_.Exception.Message)" -Type Warning }
    }

    return $results
}

# Main
try {
    if ($script:DetailedLogging) {
        Write-Log "=== ASD OWA Mailbox Policy Settings Check Started ==="
        Write-Log "Script Version: $scriptVersion"
        Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)"
        Write-Log "Log Path: $script:LogPath"
    }

    Write-ColorOutput "`n========================================" -Type Info
    Write-ColorOutput "  ASD OWA Mailbox Policy Settings Check" -Type Info
    Write-ColorOutput "========================================" -Type Info
    $isUrl = $script:BaselinePath -match '^https?://'
    if ($isUrl) { Write-ColorOutput "Baseline: GitHub (latest)" -Type Info } elseif (Test-Path $script:BaselinePath) { Write-ColorOutput "Baseline: Local File (found)" -Type Success } else { Write-ColorOutput "Baseline: Local File (not found)" -Type Warning }
    Write-ColorOutput "Location: $script:BaselinePath" -Type Info
    Write-ColorOutput "Output:   $parentPath`n" -Type Info

    $baseline = Get-BaselineSettings -Path $BaselinePath
    if (-not $baseline) { Write-ColorOutput "Failed to load baseline settings. Cannot proceed." -Type Error; exit 1 }

    if (-not (Test-ExchangeModule)) { Write-ColorOutput "`nExchangeOnlineManagement module is required." -Type Error; exit 1 }
    if (-not (Connect-EXO)) { Write-ColorOutput "`nFailed to connect to Exchange Online." -Type Error; exit 1 }
    if (-not (Test-ExchangePermissions)) { Write-ColorOutput "`nInsufficient permissions to read settings." -Type Error; exit 1 }

    $null = Invoke-OwaPolicyCheck -Requirements $baseline
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
