<#
.SYNOPSIS
    Check Exchange Online Remote Domains settings against ASD Blueprint requirements

.DESCRIPTION
    This script checks the Exchange Online Remote Domains configuration against 
    ASD's Blueprint for Secure Cloud requirements.
    
    Reference: https://blueprint.asd.gov.au/configuration/exchange-online/mail-flow/remote-domains/

.EXAMPLE
    .\asd-remotedomain-get.ps1
    
    Connects to Exchange Online, checks remote domain settings, and automatically generates 
    an HTML report in the script directory that opens in the default browser

.EXAMPLE
    .\asd-remotedomain-get.ps1 -ExportToCSV
    
    Runs the check with both HTML report (automatic) and CSV export.
    All files created in parent directory.

.EXAMPLE
    .\asd-remotedomain-get.ps1 -BaselinePath "C:\Baselines\prod-remote-domains.json"
    
    Uses a custom baseline JSON file for the compliance check

.EXAMPLE
    .\asd-remotedomain-get.ps1 -BaselinePath ".\baselines\dev-environment.json" -ExportToCSV
    
    Uses a development environment baseline and exports results to CSV

.EXAMPLE
    .\asd-remotedomain-get.ps1 -CSVPath "C:\Reports\custom-report.csv" -ExportToCSV
    
    Exports CSV to a custom location instead of the default parent directory

.NOTES
    Author: CIAOPS
    Date: 2025-11-04
    Version: 1.0
    
    Requirements:
    - ExchangeOnlineManagement PowerShell module
    - Exchange Online Permissions: View-Only Organization Management role (minimum) or Exchange Administrator
    - Baseline File (optional): Place remote-domains.json in parent directory of script
      If not found, script will fall back to built-in ASD Blueprint defaults
    
    File Locations (Default):
    - Baseline: {parent-directory}\remote-domains.json
    - HTML Report: {parent-directory}\asd-remotedomain-get-{timestamp}.html
    - CSV Export: {parent-directory}\asd-remotedomain-get-{timestamp}.csv

.LINK
    https://github.com/directorcia/office365
    https://github.com/directorcia/Office365/wiki/ASD-Remote-Domain-Configuration-Check - Documentation
    https://blueprint.asd.gov.au/configuration/exchange-online/mail-flow/remote-domains/
#>

#Requires -Modules ExchangeOnlineManagement

[CmdletBinding()]
param(
    [switch]$ExportToCSV,
    [string]$CSVPath,
    [Parameter(HelpMessage = "Path to baseline JSON file. Defaults to remote-domains.json in parent directory")]
    [string]$BaselinePath
)

# Get script and parent directory paths
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$parentPath = Split-Path -Parent $scriptPath

# Set default paths for all files in parent directory
if (-not $CSVPath) {
    $CSVPath = Join-Path $parentPath "asd-remotedomain-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
}

$HTMLPath = Join-Path $parentPath "asd-remotedomain-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"

# Set default baseline path if not provided (in parent directory)
if (-not $BaselinePath) {
    $BaselinePath = Join-Path $parentPath 'remote-domains.json'
}

# Make baseline path available at script scope for HTML report
$script:BaselinePath = $BaselinePath
$script:baselineLoaded = $false

# Script variables
$scriptVersion = "1.0"
$scriptName = "ASD Remote Domains Check"

# Color output functions
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    
    switch ($Type) {
        "Success" { Write-Host $Message -ForegroundColor Green }
        "Warning" { Write-Host $Message -ForegroundColor Yellow }
        "Error"   { Write-Host $Message -ForegroundColor Red }
        "Info"    { Write-Host $Message -ForegroundColor Cyan }
        default   { Write-Host $Message }
    }
}

# Helper to safely read JSON values
function Get-BaselineValue {
    param(
        [object]$Parent,
        [string]$Property,
        [object]$Default
    )

    if ($null -eq $Parent) { return $Default }
    try {
        $val = $Parent.$Property
        if ($null -eq $val) { return $Default }
        return $val
    }
    catch {
        return $Default
    }
}

# Validate baseline JSON schema
function Test-BaselineSchema {
    param(
        [object]$Baseline
    )
    
    $requiredFields = @(
        @{Path = 'RemoteDomain'; Type = 'Object'; Description = 'Root RemoteDomain object'},
        @{Path = 'RemoteDomain.Name'; Type = 'String'; Description = 'Remote domain name'},
        @{Path = 'RemoteDomain.DomainName'; Type = 'String'; Description = 'Domain name pattern'},
        @{Path = 'RemoteDomain.EmailReplyTypes'; Type = 'Object'; Description = 'Email reply types configuration'},
        @{Path = 'RemoteDomain.MessageReporting'; Type = 'Object'; Description = 'Message reporting configuration'},
        @{Path = 'RemoteDomain.TextAndCharacterSet'; Type = 'Object'; Description = 'Text and character set configuration'}
    )
    
    $missingFields = @()
    
    foreach ($field in $requiredFields) {
        $pathParts = $field.Path -split '\.'
        $current = $Baseline
        $found = $true
        
        foreach ($part in $pathParts) {
            if ($null -eq $current) {
                $found = $false
                break
            }
            
            try {
                $current = $current.$part
                if ($null -eq $current) {
                    $found = $false
                    break
                }
            }
            catch {
                $found = $false
                break
            }
        }
        
        if (-not $found) {
            $missingFields += @{
                Path = $field.Path
                Description = $field.Description
            }
        }
    }
    
    if ($missingFields.Count -gt 0) {
        Write-ColorOutput "`nBaseline JSON schema validation FAILED!" -Type Error
        Write-ColorOutput "Missing required fields:" -Type Error
        foreach ($missing in $missingFields) {
            Write-ColorOutput "  - $($missing.Path): $($missing.Description)" -Type Error
        }
        Write-Host ""
        return $false
    }
    
    Write-ColorOutput "Baseline JSON schema validation passed." -Type Success
    return $true
}

# Load baseline from JSON file
function Get-BaselineSettings {
    param(
        [string]$Path
    )
    
    $remoteDomainBaseline = $null
    $script:baselineLoaded = $false

    if (Test-Path $Path) {
        try {
            Write-ColorOutput "Loading baseline settings from: $Path" -Type Info
            $json = Get-Content -Path $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            
            # Validate schema
            if (Test-BaselineSchema -Baseline $json) {
                $remoteDomainBaseline = $json.RemoteDomain
                $script:baselineLoaded = $true
                Write-ColorOutput "Baseline loaded successfully from JSON file.`n" -Type Success
            }
            else {
                Write-ColorOutput "Baseline JSON file has invalid schema - falling back to built-in defaults`n" -Type Warning
                $remoteDomainBaseline = $null
            }
        }
        catch {
            Write-ColorOutput "Failed to parse baseline JSON: $($_.Exception.Message)" -Type Error
            Write-ColorOutput "Error at line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())" -Type Error
            Write-ColorOutput "Falling back to built-in defaults`n" -Type Warning
            $remoteDomainBaseline = $null
        }
    }
    else {
        Write-ColorOutput "Baseline file not found at: $Path" -Type Warning
        Write-ColorOutput "Falling back to built-in defaults`n" -Type Warning
    }

    # Build and return ASD Blueprint requirements (from baseline if available, otherwise defaults)
    return @{
        Name = (Get-BaselineValue -Parent $remoteDomainBaseline -Property 'Name' -Default 'Default')
        DomainName = (Get-BaselineValue -Parent $remoteDomainBaseline -Property 'DomainName' -Default '*')
        # Email reply types
        AllowedOOFType = (Get-BaselineValue -Parent $remoteDomainBaseline.EmailReplyTypes -Property 'AllowedOOFType' -Default 'External')
        AutoReplyEnabled = (Get-BaselineValue -Parent $remoteDomainBaseline.EmailReplyTypes -Property 'AutoReplyEnabled' -Default $false)
        AutoForwardEnabled = (Get-BaselineValue -Parent $remoteDomainBaseline.EmailReplyTypes -Property 'AutoForwardEnabled' -Default $false)
        # Message reporting
        DeliveryReportEnabled = (Get-BaselineValue -Parent $remoteDomainBaseline.MessageReporting -Property 'DeliveryReportEnabled' -Default $false)
        NDREnabled = (Get-BaselineValue -Parent $remoteDomainBaseline.MessageReporting -Property 'NDREnabled' -Default $false)
        MeetingForwardNotificationEnabled = (Get-BaselineValue -Parent $remoteDomainBaseline.MessageReporting -Property 'MeetingForwardNotificationEnabled' -Default $false)
        # Text and character set
        TNEFEnabled = (Get-BaselineValue -Parent $remoteDomainBaseline.TextAndCharacterSet -Property 'TNEFEnabled' -Default $null)
        CharacterSet = (Get-BaselineValue -Parent $remoteDomainBaseline.TextAndCharacterSet -Property 'CharacterSet' -Default $null)
        NonMimeCharacterSet = (Get-BaselineValue -Parent $remoteDomainBaseline.TextAndCharacterSet -Property 'NonMimeCharacterSet' -Default $null)
    }
}

# Check if ExchangeOnlineManagement module is installed and load it
function Test-ExchangeModule {
    Write-ColorOutput "Checking for ExchangeOnlineManagement module..." -Type Info
    
    # Check if module is already loaded
    if (Get-Module -Name ExchangeOnlineManagement) {
        Write-ColorOutput "ExchangeOnlineManagement module already loaded." -Type Success
        return $true
    }
    
    # Check if module is available
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-ColorOutput "ExchangeOnlineManagement module not found!" -Type Error
        Write-ColorOutput "Install it with: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser" -Type Warning
        return $false
    }
    
    # Load the module
    Write-ColorOutput "Loading ExchangeOnlineManagement module..." -Type Info
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Write-ColorOutput "ExchangeOnlineManagement module loaded successfully." -Type Success
        return $true
    }
    catch {
        Write-ColorOutput "Failed to load ExchangeOnlineManagement module: $($_.Exception.Message)" -Type Error
        return $false
    }
}

# Connect to Exchange Online
function Connect-EXO {
    Write-ColorOutput "`nChecking Exchange Online connection..." -Type Info
    
    try {
        # Try to run a simple command to test if already connected
        try {
            $null = Get-OrganizationConfig -ErrorAction Stop
            Write-ColorOutput "Already connected to Exchange Online." -Type Success
            return $true
        }
        catch {
            # Not connected or connection expired, need to authenticate
            Write-ColorOutput "Not connected. Connecting to Exchange Online..." -Type Info
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            Write-ColorOutput "Successfully connected to Exchange Online." -Type Success
            return $true
        }
    }
    catch {
        Write-ColorOutput "Failed to connect to Exchange Online: $($_.Exception.Message)" -Type Error
        return $false
    }
}

# Check a single setting
function Test-Setting {
    param(
        [string]$SettingName,
        [object]$CurrentValue,
        [object]$RequiredValue,
        [string]$Description
    )
    
    $result = @{
        Setting = $SettingName
        Description = $Description
        CurrentValue = if ($null -eq $CurrentValue) { "Not set" } else { $CurrentValue.ToString() }
        RequiredValue = if ($null -eq $RequiredValue) { "Not set" } else { $RequiredValue.ToString() }
        Compliant = $false
        Status = ""
    }
    
    # Special handling for null values (meaning "not set" or "follow defaults")
    if ($null -eq $RequiredValue -and $null -eq $CurrentValue) {
        $result.Compliant = $true
        $result.Status = "PASS"
    }
    elseif ($null -eq $RequiredValue) {
        # If required is null, we accept any value
        $result.Compliant = $true
        $result.Status = "PASS"
    }
    elseif ($CurrentValue -eq $RequiredValue) {
        $result.Compliant = $true
        $result.Status = "PASS"
    }
    else {
        $result.Compliant = $false
        $result.Status = "FAIL"
    }
    
    return $result
}

# Generate HTML Report
function New-HTMLReport {
    param(
        [array]$CheckResults,
        [object]$RemoteDomain,
        [string]$OutputPath
    )
    
    $totalChecks = $CheckResults.Count
    $passedChecks = ($CheckResults | Where-Object { $_.Compliant }).Count
    $failedChecks = $totalChecks - $passedChecks
    $compliancePercentage = [math]::Round(($passedChecks / $totalChecks) * 100, 2)
    $overallStatus = if ($compliancePercentage -eq 100) { "COMPLIANT" } else { "NON-COMPLIANT" }
    $statusColor = if ($compliancePercentage -eq 100) { "#28a745" } else { "#dc3545" }
    
    $reportDate = Get-Date -Format "MMMM dd, yyyy - HH:mm:ss"
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ASD Remote Domains Compliance Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            line-height: 1.6;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        
        .header p {
            font-size: 1.1em;
            opacity: 0.9;
        }
        
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #f8f9fa;
        }
        
        .summary-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            text-align: center;
            transition: transform 0.3s ease;
        }
        
        .summary-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 5px 20px rgba(0,0,0,0.15);
        }
        
        .summary-card h3 {
            color: #6c757d;
            font-size: 0.9em;
            text-transform: uppercase;
            margin-bottom: 10px;
            letter-spacing: 1px;
        }
        
        .summary-card .value {
            font-size: 2.5em;
            font-weight: bold;
            margin: 10px 0;
        }
        
        .summary-card.total .value { color: #007bff; }
        .summary-card.passed .value { color: #28a745; }
        .summary-card.failed .value { color: #dc3545; }
        .summary-card.compliance .value { color: $statusColor; }
        
        .info-section {
            padding: 30px;
            background: white;
            border-bottom: 1px solid #e9ecef;
        }
        
        .info-section h2 {
            color: #2a5298;
            margin-bottom: 15px;
            font-size: 1.5em;
            border-bottom: 2px solid #2a5298;
            padding-bottom: 10px;
        }
        
        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }
        
        .info-item {
            background: #f8f9fa;
            padding: 12px;
            border-radius: 5px;
            border-left: 4px solid #2a5298;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }
        
        .info-item strong {
            color: #495057;
            display: inline-block;
            min-width: 120px;
        }
        
        .results-section {
            padding: 30px;
        }
        
        .results-section h2 {
            color: #2a5298;
            margin-bottom: 20px;
            font-size: 1.5em;
            border-bottom: 2px solid #2a5298;
            padding-bottom: 10px;
        }
        
        .result-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        .result-table thead {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
        }
        
        .result-table th {
            padding: 15px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.85em;
            letter-spacing: 0.5px;
        }
        
        .result-table td {
            padding: 12px 15px;
            border-bottom: 1px solid #e9ecef;
        }
        
        .result-table tbody tr {
            background: white;
            transition: background-color 0.2s ease;
        }
        
        .result-table tbody tr:hover {
            background: #f8f9fa;
        }
        
        .result-table tbody tr:nth-child(even) {
            background: #f8f9fa;
        }
        
        .result-table tbody tr:nth-child(even):hover {
            background: #e9ecef;
        }
        
        .status-badge {
            display: inline-block;
            padding: 5px 12px;
            border-radius: 20px;
            font-weight: bold;
            font-size: 0.85em;
            text-transform: uppercase;
        }
        
        .status-pass {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        
        .status-fail {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        
        .status-icon {
            font-size: 1.2em;
            margin-right: 5px;
        }
        
        .overall-status {
            text-align: center;
            padding: 30px;
            background: $statusColor;
            color: white;
        }
        
        .overall-status h2 {
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        
        .footer {
            padding: 20px;
            text-align: center;
            background: #f8f9fa;
            color: #6c757d;
            font-size: 0.9em;
        }
        
        .footer a {
            color: #2a5298;
            text-decoration: none;
            font-weight: bold;
        }
        
        .footer a:hover {
            text-decoration: underline;
        }
        
        .timestamp {
            color: #ffffff;
            font-size: 1em;
            margin-top: 10px;
            opacity: 0.95;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            
            .container {
                box-shadow: none;
            }
            
            .summary-card:hover {
                transform: none;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🛡️ ASD Remote Domains Compliance Report</h1>
            <p>Exchange Online Remote Domains Configuration Check</p>
            <p class="timestamp">Generated: $reportDate</p>
        </div>
        
        <div class="summary">
            <div class="summary-card total">
                <h3>Total Checks</h3>
                <div class="value">$totalChecks</div>
            </div>
            <div class="summary-card passed">
                <h3>Passed</h3>
                <div class="value">$passedChecks</div>
            </div>
            <div class="summary-card failed">
                <h3>Failed</h3>
                <div class="value">$failedChecks</div>
            </div>
            <div class="summary-card compliance">
                <h3>Compliance</h3>
                <div class="value">$compliancePercentage%</div>
            </div>
        </div>
        
        <div class="info-section">
            <h2>📋 Domain Information</h2>
            <div class="info-grid">
                <div class="info-item">
                    <strong>Identity:</strong> $($RemoteDomain.Identity)
                </div>
                <div class="info-item">
                    <strong>Domain Name:</strong> $($RemoteDomain.DomainName)
                </div>
                <div class="info-item">
                    <strong>Distinguished Name:</strong> $($RemoteDomain.DistinguishedName)
                </div>
                <div class="info-item">
                    <strong>Baseline Source:</strong> $(if ($script:baselineLoaded) { 
                        $fileName = Split-Path -Leaf $script:BaselinePath
                        $location = if ($script:BaselinePath -match '^https?://') { "Online" } else { "Local" }
                        "$fileName ($location)"
                    } else { 
                        "Built-in Defaults" 
                    })
                </div>
            </div>
        </div>
        
        <div class="results-section">
            <h2>🔍 Detailed Check Results</h2>
            <table class="result-table">
                <thead>
                    <tr>
                        <th>Status</th>
                        <th>Setting</th>
                        <th>Description</th>
                        <th>Current Value</th>
                        <th>Required Value</th>
                    </tr>
                </thead>
                <tbody>
"@

    foreach ($result in $CheckResults) {
        $statusClass = if ($result.Compliant) { "status-pass" } else { "status-fail" }
        $statusIcon = if ($result.Compliant) { "✓" } else { "✗" }
        $statusText = if ($result.Compliant) { "PASS" } else { "FAIL" }
        
        $html += @"
                    <tr>
                        <td>
                            <span class="status-badge $statusClass">
                                <span class="status-icon">$statusIcon</span>$statusText
                            </span>
                        </td>
                        <td><strong>$($result.Setting)</strong></td>
                        <td>$($result.Description)</td>
                        <td>$($result.CurrentValue)</td>
                        <td>$($result.RequiredValue)</td>
                    </tr>
"@
    }

    $html += @"
                </tbody>
            </table>
        </div>
        
        <div class="overall-status">
            <h2>Overall Status: $overallStatus</h2>
            <p style="font-size: 1.2em; margin-top: 10px;">
                $($passedChecks) out of $($totalChecks) checks passed
            </p>
        </div>
        
        <div class="footer">
            <p><strong>Reference:</strong> <a href="https://blueprint.asd.gov.au/configuration/exchange-online/mail-flow/remote-domains/" target="_blank">ASD's Blueprint for Secure Cloud - Remote Domains</a></p>
        </div>
    </div>
</body>
</html>
"@

    try {
        $html | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
        return $true
    }
    catch {
        Write-ColorOutput "Failed to generate HTML report: $($_.Exception.Message)" -Type Error
        return $false
    }
}

# Main check function
function Invoke-RemoteDomainCheck {
    param(
        [hashtable]$Requirements
    )
    
    Write-ColorOutput "`n========================================" -Type Info
    Write-ColorOutput "  $scriptName v$scriptVersion" -Type Info
    Write-ColorOutput "  ASD Blueprint Compliance Check" -Type Info
    Write-ColorOutput "========================================`n" -Type Info
    
    # Get the Default remote domain
    Write-ColorOutput "Retrieving Default remote domain configuration..." -Type Info
    
    try {
        $remoteDomain = Get-RemoteDomain -Identity "Default" -ErrorAction Stop
        
        if (-not $remoteDomain) {
            Write-ColorOutput "Default remote domain not found!" -Type Error
            return
        }
        
        Write-ColorOutput "Default remote domain found: $($remoteDomain.DomainName)`n" -Type Success
        
    }
    catch {
        Write-ColorOutput "Failed to retrieve remote domain: $($_.Exception.Message)" -Type Error
        return
    }
    
    # Array to store all check results
    $checkResults = @()
    
    # Check each setting
    Write-ColorOutput "Checking settings against ASD Blueprint requirements...`n" -Type Info
    
    # Domain Name
    $checkResults += Test-Setting -SettingName "DomainName" `
        -CurrentValue $remoteDomain.DomainName `
        -RequiredValue $Requirements.DomainName `
        -Description "Remote Domain (should be *)"
    
    # Out of Office Type
    $checkResults += Test-Setting -SettingName "AllowedOOFType" `
        -CurrentValue $remoteDomain.AllowedOOFType `
        -RequiredValue $Requirements.AllowedOOFType `
        -Description "Out of Office automatic reply types"
    
    # Auto Reply
    $checkResults += Test-Setting -SettingName "AutoReplyEnabled" `
        -CurrentValue $remoteDomain.AutoReplyEnabled `
        -RequiredValue $Requirements.AutoReplyEnabled `
        -Description "Allow automatic replies"
    
    # Auto Forward
    $checkResults += Test-Setting -SettingName "AutoForwardEnabled" `
        -CurrentValue $remoteDomain.AutoForwardEnabled `
        -RequiredValue $Requirements.AutoForwardEnabled `
        -Description "Allow automatic forwarding"
    
    # Delivery Reports
    $checkResults += Test-Setting -SettingName "DeliveryReportEnabled" `
        -CurrentValue $remoteDomain.DeliveryReportEnabled `
        -RequiredValue $Requirements.DeliveryReportEnabled `
        -Description "Allow delivery reports"
    
    # NDR (Non-Delivery Reports)
    $checkResults += Test-Setting -SettingName "NDREnabled" `
        -CurrentValue $remoteDomain.NDREnabled `
        -RequiredValue $Requirements.NDREnabled `
        -Description "Allow non-delivery reports"
    
    # Meeting Forward Notifications
    $checkResults += Test-Setting -SettingName "MeetingForwardNotificationEnabled" `
        -CurrentValue $remoteDomain.MeetingForwardNotificationEnabled `
        -RequiredValue $Requirements.MeetingForwardNotificationEnabled `
        -Description "Allow meeting forward notifications"
    
    # TNEF (Rich Text Format)
    $checkResults += Test-Setting -SettingName "TNEFEnabled" `
        -CurrentValue $remoteDomain.TNEFEnabled `
        -RequiredValue $Requirements.TNEFEnabled `
        -Description "Use rich-text format (null = Follow user settings)"
    
    # Character Set
    $checkResults += Test-Setting -SettingName "CharacterSet" `
        -CurrentValue $remoteDomain.CharacterSet `
        -RequiredValue $Requirements.CharacterSet `
        -Description "MIME character set"
    
    # Non-MIME Character Set
    $checkResults += Test-Setting -SettingName "NonMimeCharacterSet" `
        -CurrentValue $remoteDomain.NonMimeCharacterSet `
        -RequiredValue $Requirements.NonMimeCharacterSet `
        -Description "Non-MIME character set"
    
    # Display results
    Write-ColorOutput "`n========================================" -Type Info
    Write-ColorOutput "  CHECK RESULTS" -Type Info
    Write-ColorOutput "========================================`n" -Type Info
    
    foreach ($result in $checkResults) {
        $statusColor = if ($result.Compliant) { "Success" } else { "Error" }
        $statusSymbol = if ($result.Compliant) { "[✓]" } else { "[✗]" }
        
        Write-ColorOutput "$statusSymbol $($result.Setting)" -Type $statusColor
        Write-Host "    Description : $($result.Description)"
        Write-Host "    Current     : $($result.CurrentValue)"
        Write-Host "    Required    : $($result.RequiredValue)"
        Write-Host "    Status      : $($result.Status)"
        Write-Host ""
    }
    
    # Summary
    $totalChecks = $checkResults.Count
    $passedChecks = ($checkResults | Where-Object { $_.Compliant }).Count
    $failedChecks = $totalChecks - $passedChecks
    $compliancePercentage = [math]::Round(($passedChecks / $totalChecks) * 100, 2)
    
    Write-ColorOutput "========================================" -Type Info
    Write-ColorOutput "  SUMMARY" -Type Info
    Write-ColorOutput "========================================" -Type Info
    Write-Host "Total Checks    : $totalChecks"
    Write-ColorOutput "Passed          : $passedChecks" -Type Success
    
    if ($failedChecks -gt 0) {
        Write-ColorOutput "Failed          : $failedChecks" -Type Error
    } else {
        Write-ColorOutput "Failed          : $failedChecks" -Type Success
    }
    
    Write-Host "Compliance      : $compliancePercentage%"
    
    if ($compliancePercentage -eq 100) {
        Write-ColorOutput "`nStatus          : COMPLIANT ✓" -Type Success
    } else {
        Write-ColorOutput "`nStatus          : NON-COMPLIANT ✗" -Type Error
    }
    
    Write-ColorOutput "========================================`n" -Type Info
    
    # Export to CSV if requested
    if ($ExportToCSV) {
        try {
            $checkResults | Select-Object Setting, Description, CurrentValue, RequiredValue, Status | 
                Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8
            Write-ColorOutput "Results exported to: $CSVPath" -Type Success
        }
        catch {
            Write-ColorOutput "Failed to export results: $($_.Exception.Message)" -Type Error
        }
    }
    
    # Generate HTML Report (always)
    Write-ColorOutput "`nGenerating HTML report..." -Type Info
    if (New-HTMLReport -CheckResults $checkResults -RemoteDomain $remoteDomain -OutputPath $script:HTMLPath) {
        Write-ColorOutput "HTML report generated: $script:HTMLPath" -Type Success
        Write-ColorOutput "Opening report in default browser..." -Type Info
        try {
            Start-Process $script:HTMLPath -ErrorAction Stop
        }
        catch {
            Write-ColorOutput "Could not automatically open browser: $($_.Exception.Message)" -Type Warning
            Write-ColorOutput "Please open the report manually: $script:HTMLPath" -Type Warning
        }
    }
    else {
        Write-ColorOutput "Failed to generate HTML report." -Type Error
    }
    
    return $checkResults
}

# Main execution
try {
    Write-ColorOutput "`n========================================" -Type Info
    Write-ColorOutput "  ASD Remote Domain Configuration Check" -Type Info
    Write-ColorOutput "========================================" -Type Info
    Write-ColorOutput "Baseline: $(if (Test-Path $script:BaselinePath) { 'Found' } else { 'Not Found (will use defaults)' })" -Type $(if (Test-Path $script:BaselinePath) { 'Success' } else { 'Warning' })
    Write-ColorOutput "Location: $script:BaselinePath" -Type Info
    Write-ColorOutput "Output:   $parentPath`n" -Type Info
    
    # Load baseline settings
    $asdRequirements = Get-BaselineSettings -Path $BaselinePath
    
    # Check module
    if (-not (Test-ExchangeModule)) {
        exit 1
    }
    
    # Connect to Exchange Online
    if (-not (Connect-EXO)) {
        exit 1
    }
    
    # Run the check
    Invoke-RemoteDomainCheck -Requirements $asdRequirements | Out-Null
    
    Write-ColorOutput "`nScript completed successfully." -Type Success
}
catch {
    Write-ColorOutput "`nScript failed with error: $($_.Exception.Message)" -Type Error
    Write-ColorOutput $_.ScriptStackTrace -Type Error
    exit 1
}
