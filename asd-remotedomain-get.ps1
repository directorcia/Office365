<#
.SYNOPSIS
    Check Exchange Online Remote Domain settings against ASD Blueprint requirements

.DESCRIPTION
    This script checks all Exchange Online Remote Domain configurations against 
    ASD's Blueprint for Secure Cloud requirements. It validates all remote domains
    in the tenant, including the 'Default' domain and any custom domains (such as 
    domains configured for email forwarding).
    
    Reference: https://blueprint.asd.gov.au/configuration/exchange-online/mail-flow/remote-domains/

.EXAMPLE
    .\asd-remotedomain-get.ps1
    
    Connects to Exchange Online, downloads latest baseline from GitHub, checks remote domain 
    settings, and automatically generates an HTML report in the script directory that opens 
    in the default browser

.EXAMPLE
    .\asd-remotedomain-get.ps1 -ExportToCSV
    
    Runs the check with both HTML report (automatic) and CSV export.
    Downloads latest baseline from GitHub. All files created in parent directory.

.EXAMPLE
    .\asd-remotedomain-get.ps1 -BaselinePath "C:\Baselines\prod-remote-domains.json"
    
    Uses a custom baseline JSON file for the compliance check

.EXAMPLE
    .\asd-remotedomain-get.ps1 -BaselinePath ".\baselines\dev-environment.json" -ExportToCSV
    
    Uses a development environment baseline and exports results to CSV

.EXAMPLE
    .\asd-remotedomain-get.ps1 -CSVPath "C:\Reports\custom-report.csv" -ExportToCSV
    
    Exports CSV to a custom location instead of the default parent directory

.EXAMPLE
    .\asd-remotedomain-get.ps1 -DetailedLogging
    
    Runs the check with detailed logging enabled. Log file created in parent directory

.EXAMPLE
    .\asd-remotedomain-get.ps1 -DetailedLogging -LogPath "C:\Logs\custom-log.log"
    
    Runs the check with detailed logging to a custom log file location

.NOTES
    Author: CIAOPS
    Date: 11-04-2025
    Version: 1.0
    
    Requirements:
    - ExchangeOnlineManagement PowerShell module
    - Exchange Online Permissions (one of the following roles):
      * Exchange Administrator
      * Global Administrator
      * Global Reader
      * View-Only Organization Management
      * Compliance Administrator
    - Internet connection (to download baseline from GitHub by default)
    
    Baseline Sources (in order of precedence):
    1. Custom path specified via -BaselinePath parameter
    2. GitHub (default): https://raw.githubusercontent.com/directorcia/bp/main/ASD/Exchange-Online/Mail-flow/remote-domains.json
    3. Built-in ASD Blueprint defaults (fallback if GitHub unavailable)
    
    File Locations (Default):
    - HTML Report: {parent-directory}\asd-remotedomain-get-{timestamp}.html
    - CSV Export: {parent-directory}\asd-remotedomain-get-{timestamp}.csv
    - Log File (if enabled): {parent-directory}\asd-remotedomain-get-{timestamp}.log

.LINK
    https://github.com/directorcia/office365
    https://github.com/directorcia/Office365/wiki/ASD-Remote-Domain-Configuration-Check - Documentation
    https://github.com/directorcia/bp/wiki/Exchange-Online-Remote-Domain-Security-Controls - Exchange Online Remote Domain Security Controls
    
    https://blueprint.asd.gov.au/configuration/exchange-online/mail-flow/remote-domains/
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
    [string]$LogPath
)

# Get script and parent directory paths
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$parentPath = Split-Path -Parent $scriptPath

# Set default paths for all files in parent directory
if (-not $CSVPath) {
    $CSVPath = Join-Path $parentPath "asd-remotedomain-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
}

# Set default log path if detailed logging is enabled
if ($DetailedLogging -and -not $LogPath) {
    $LogPath = Join-Path $parentPath "asd-remotedomain-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
}

# Default GitHub URL for baseline settings
$defaultGitHubURL = "https://raw.githubusercontent.com/directorcia/bp/main/ASD/Exchange-Online/Mail-flow/remote-domains.json"

# Set default baseline path if not provided (in parent directory)
if (-not $BaselinePath) {
    $BaselinePath = $defaultGitHubURL
}

# Script-scope variables for tracking state
$script:BaselinePath = $BaselinePath
$script:baselineLoaded = $false
$script:HTMLPath = Join-Path $parentPath "asd-remotedomain-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
$script:LogPath = $LogPath
$script:DetailedLogging = $DetailedLogging

# Script variables
$scriptVersion = "1.0"
$scriptName = "ASD Remote Domains Check"

# Logging function
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    if ($script:DetailedLogging -and $script:LogPath) {
        try {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logEntry = "[$timestamp] [$Level] $Message"
            Add-Content -Path $script:LogPath -Value $logEntry -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to write to log file: $($_.Exception.Message)"
        }
    }
}

# Color output functions
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    
    # Map Type to log Level
    $logLevel = switch ($Type) {
        "Success" { "INFO" }
        "Warning" { "WARN" }
        "Error"   { "ERROR" }
        "Info"    { "INFO" }
        default   { "INFO" }
    }
    
    # Write to log
    Write-Log -Message $Message -Level $logLevel
    
    # Write to console
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
    # Validate it's an array with at least one element
    if ($null -eq $Baseline) { return $false }
    if ($Baseline -is [string]) { return $false }
    if (-not ($Baseline -is [System.Array] -or $Baseline -is [System.Collections.ArrayList])) { return $false }
    if ($Baseline.Count -lt 1) { return $false }
    
    $requiredFields = @(
        @{Path = 'Identity'; Description = 'Remote domain identity (Default)'},
        @{Path = 'DomainName'; Description = 'Domain name pattern'},
        @{Path = 'AllowedOOFType'; Description = 'Out of Office type'},
        @{Path = 'AutoReplyEnabled'; Description = 'Automatic replies allowed'},
        @{Path = 'AutoForwardEnabled'; Description = 'Automatic forwarding allowed'},
        @{Path = 'DeliveryReportEnabled'; Description = 'Delivery reports allowed'},
        @{Path = 'NDRRequired'; Description = 'NDR required'},
        @{Path = 'MeetingForwardNotificationEnabled'; Description = 'Meeting forward notifications'},
        @{Path = 'TNEFEnabled'; Description = 'TNEF (Rich Text)'},
        @{Path = 'CharacterSet'; Description = 'MIME character set'},
        @{Path = 'NonMIMECharacterSet'; Description = 'Non-MIME character set'}
    )
    $BaselineToCheck = $Baseline[0]
    
    $missingFields = @()
    
    foreach ($field in $requiredFields) {
        $pathParts = $field.Path -split '\.'
        $current = $BaselineToCheck
        $found = $true
        
        foreach ($part in $pathParts) {
            if ($null -eq $current) {
                $found = $false
                break
            }
            
            try {
                # Check if property exists (not if it's null - null values are valid)
                $properties = $current.PSObject.Properties.Name
                if ($properties -notcontains $part) {
                    $found = $false
                    break
                }
                $current = $current.$part
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
        Write-ColorOutput "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -Type Error
        Write-ColorOutput "❌ BASELINE JSON SCHEMA VALIDATION FAILED" -Type Error
        Write-ColorOutput "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -Type Error
        Write-Host ""
        Write-ColorOutput "Missing required fields:" -Type Error
        foreach ($missing in $missingFields) {
            Write-Host "  • $($missing.Path)"
            Write-Host "    └─ $($missing.Description)"
        }
        Write-Host ""
        Write-ColorOutput "The baseline JSON file does not conform to the expected schema." -Type Warning
        Write-ColorOutput "Please check the file format or use the default GitHub baseline." -Type Warning
        Write-Host ""
        Write-ColorOutput "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -Type Error
        return $false
    }
    
    return $true
}

# Load baseline from JSON file or URL
function Get-BaselineSettings {
    param([string]$Path)
    Write-Log "Starting baseline settings load from: $Path" -Level "INFO"
    $script:baselineLoaded = $false
    $isUrl = $Path -match '^https?://'
    $record = $null
    try {
        $raw = if ($isUrl) {
            Write-Progress -Activity "Loading Baseline" -Status "Downloading..." -PercentComplete 20
            (Invoke-WebRequest -Uri $Path -UseBasicParsing -ErrorAction Stop).Content
        } else {
            Write-Progress -Activity "Loading Baseline" -Status "Reading file..." -PercentComplete 20
            Get-Content -Path $Path -Raw -ErrorAction Stop
        }
        Write-Progress -Activity "Loading Baseline" -Status "Parsing JSON..." -PercentComplete 45
        $json = $raw | ConvertFrom-Json -ErrorAction Stop
        if (-not (Test-BaselineSchema -Baseline $json)) { throw "Baseline JSON schema invalid (array format expected)." }
        $record = ($json | Where-Object { $_.Identity -eq 'Default' } | Select-Object -First 1)
        if (-not $record) { throw "No 'Default' identity record found in baseline." }
        $script:baselineLoaded = $true
        Write-Progress -Activity "Loading Baseline" -Completed
        Write-ColorOutput "✓ Baseline loaded successfully (array schema)" -Type Success
    }
    catch {
        Write-Progress -Activity "Loading Baseline" -Completed
        Write-Log "Baseline load failed: $($_.Exception.Message)" -Level "WARN"
        Write-ColorOutput "Failed to load baseline: $($_.Exception.Message)" -Type Warning
        Write-ColorOutput "Using built-in ASD Blueprint defaults instead." -Type Warning
        # Built-in defaults (ASD recommended posture)
        $record = [pscustomobject]@{
            Identity = 'Default'
            DomainName = '*'
            AllowedOOFType = 'External'
            AutoReplyEnabled = $true
            AutoForwardEnabled = $false
            DeliveryReportEnabled = $true
            NDRRequired = $true
            MeetingForwardNotificationEnabled = $true
            TNEFEnabled = $null   # Follow user settings
            CharacterSet = $null  # Use automatic (most flexible)
            NonMIMECharacterSet = $null  # Use automatic (most flexible)
        }
        $script:baselineLoaded = $false
    }

    return @{
        Name = $record.Identity
        DomainName = $record.DomainName
        AllowedOOFType = $record.AllowedOOFType
        AutoReplyEnabled = $record.AutoReplyEnabled
        AutoForwardEnabled = $record.AutoForwardEnabled
        DeliveryReportEnabled = $record.DeliveryReportEnabled
        NDRRequired = $record.NDRRequired
        MeetingForwardNotificationEnabled = $record.MeetingForwardNotificationEnabled
        TNEFEnabled = $record.TNEFEnabled
        CharacterSet = $record.CharacterSet
        NonMIMECharacterSet = $record.NonMIMECharacterSet
    }
}

# Check if ExchangeOnlineManagement module is installed and load it
function Test-ExchangeModule {
    Write-Log "Checking for ExchangeOnlineManagement module" -Level "INFO"
    Write-ColorOutput "Checking for ExchangeOnlineManagement module..." -Type Info
    
    # Check if module is already loaded
    if (Get-Module -Name ExchangeOnlineManagement) {
        Write-Log "ExchangeOnlineManagement module already loaded" -Level "INFO"
        Write-ColorOutput "ExchangeOnlineManagement module already loaded." -Type Success
        return $true
    }
    
    # Check if module is available
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log "ExchangeOnlineManagement module not found" -Level "ERROR"
        Write-ColorOutput "ExchangeOnlineManagement module not found!" -Type Error
        Write-ColorOutput "Install it with: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser" -Type Warning
        return $false
    }
    
    # Load the module
    Write-Log "Loading ExchangeOnlineManagement module" -Level "INFO"
    Write-ColorOutput "Loading ExchangeOnlineManagement module..." -Type Info
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Write-Log "ExchangeOnlineManagement module loaded successfully" -Level "INFO"
        Write-ColorOutput "ExchangeOnlineManagement module loaded successfully." -Type Success
        return $true
    }
    catch {
        Write-Log "Failed to load module: $($_.Exception.Message)" -Level "ERROR"
        Write-ColorOutput "Failed to load ExchangeOnlineManagement module: $($_.Exception.Message)" -Type Error
        return $false
    }
}

# Connect to Exchange Online
function Connect-EXO {
    Write-Log "Checking Exchange Online connection status" -Level "INFO"
    Write-ColorOutput "`nChecking Exchange Online connection..." -Type Info
    
    try {
        # Try to run a simple command to test if already connected
        try {
            $null = Get-OrganizationConfig -ErrorAction Stop
            Write-Log "Already connected to Exchange Online" -Level "INFO"
            Write-ColorOutput "Already connected to Exchange Online." -Type Success
            return $true
        }
        catch {
            # Not connected or connection expired, need to authenticate
            Write-Log "Not connected - initiating Exchange Online connection" -Level "INFO"
            Write-ColorOutput "Not connected. Connecting to Exchange Online..." -Type Info
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            Write-Log "Successfully connected to Exchange Online" -Level "INFO"
            Write-ColorOutput "Successfully connected to Exchange Online." -Type Success
            return $true
        }
    }
    catch {
        Write-Log "Failed to connect to Exchange Online: $($_.Exception.Message)" -Level "ERROR"
        Write-ColorOutput "Failed to connect to Exchange Online: $($_.Exception.Message)" -Type Error
        return $false
    }
}

# Check Exchange Online permissions
function Test-ExchangePermissions {
    Write-ColorOutput "`nValidating Exchange Online permissions..." -Type Info
    
    try {
        # Try to get organization config (requires View-Only Organization Management minimum)
        $null = Get-OrganizationConfig -ErrorAction Stop
        
        # Try to get a remote domain (specific permission test)
        $null = Get-RemoteDomain -Identity "Default" -ErrorAction Stop
        
        Write-ColorOutput "Permission validation passed." -Type Success
        return $true
    }
    catch {
        $errorMessage = $_.Exception.Message
        
        # Check if it's a permission-related error
        if ($errorMessage -match "Access.*Denied|not have permission|Insufficient|Unauthorized|Authorization|forbidden") {
            Write-ColorOutput "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -Type Error
            Write-ColorOutput "❌ INSUFFICIENT PERMISSIONS" -Type Error
            Write-ColorOutput "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -Type Error
            Write-Host ""
            Write-ColorOutput "This script requires Exchange Online read permissions." -Type Warning
            Write-Host ""
            Write-ColorOutput "Required Roles (one of the following):" -Type Info
            Write-Host "  • Exchange Administrator"
            Write-Host "  • Global Administrator" 
            Write-Host "  • Global Reader"
            Write-Host "  • View-Only Organization Management"
            Write-Host "  • Compliance Administrator"
            Write-Host ""
            Write-ColorOutput "Error Details:" -Type Warning
            Write-Host "  $errorMessage"
            Write-Host ""
            Write-ColorOutput "Action Required:" -Type Info
            Write-Host "  1. Contact your Exchange Online administrator"
            Write-Host "  2. Request one of the roles listed above"
            Write-Host "  3. Wait for role assignment to propagate (may take a few minutes)"
            Write-Host "  4. Re-run this script after role assignment"
            Write-Host ""
            Write-ColorOutput "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -Type Error
            return $false
        }
        else {
            # Some other error occurred
            Write-ColorOutput "Permission check failed: $errorMessage" -Type Error
            Write-ColorOutput "Please verify your Exchange Online connection and try again." -Type Warning
            return $false
        }
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
    
    # Perform comparison BEFORE converting to strings (to handle enums properly)
    $isCompliant = $false
    
    # Special handling for null values (meaning "not set" or "follow defaults")
    if ($null -eq $RequiredValue -and $null -eq $CurrentValue) {
        $isCompliant = $true
    }
    elseif ($null -eq $RequiredValue) {
        # If required is null, we accept any value
        $isCompliant = $true
    }
    elseif ($null -eq $CurrentValue) {
        # Current is null but required is not
        $isCompliant = $false
    }
    else {
        # Compare as strings to handle enums and other types
        $currentStr = $CurrentValue.ToString()
        $requiredStr = $RequiredValue.ToString()
        $isCompliant = ($currentStr -eq $requiredStr)
    }
    
    # Format display values with better context
    $currentDisplay = if ($null -eq $CurrentValue) { 
        "Not set (null)" 
    } elseif ($CurrentValue -is [bool]) {
        $CurrentValue.ToString()
    } else { 
        $CurrentValue.ToString() 
    }
    
    $requiredDisplay = if ($null -eq $RequiredValue) { 
        "Not set (null)" 
    } elseif ($RequiredValue -is [bool]) {
        $RequiredValue.ToString()
    } else { 
        $RequiredValue.ToString() 
    }
    
    $result = @{
        Setting = $SettingName
        Description = $Description
        CurrentValue = $currentDisplay
        RequiredValue = $requiredDisplay
        Compliant = $isCompliant
        Status = if ($isCompliant) { "PASS" } else { "FAIL" }
    }
    
    # Log the check result
    $logMsg = "Check: $SettingName - Current: $($result.CurrentValue), Required: $($result.RequiredValue), Status: $($result.Status)"
    Write-Log -Message $logMsg -Level $(if ($result.Compliant) { "INFO" } else { "WARN" })
    
    return $result
}

# Generate HTML Report
function New-HTMLReport {
    param(
        [array]$AllDomainResults,
        [string]$OutputPath
    )
    
    # Calculate overall statistics
    $totalDomains = $AllDomainResults.Count
    $totalAllChecks = ($AllDomainResults | ForEach-Object { $_.TotalChecks } | Measure-Object -Sum).Sum
    $totalAllPassed = ($AllDomainResults | ForEach-Object { $_.PassedChecks } | Measure-Object -Sum).Sum
    $totalAllFailed = $totalAllChecks - $totalAllPassed
    $overallCompliance = if ($totalAllChecks -gt 0) { [math]::Round(($totalAllPassed / $totalAllChecks) * 100, 2) } else { 0 }
    $overallStatus = if ($overallCompliance -eq 100) { "COMPLIANT" } else { "NON-COMPLIANT" }
    $statusColor = if ($overallCompliance -eq 100) { "#28a745" } else { "#dc3545" }
    
    # Get organization/domain name
    $domainName = $null
    try {
        $orgConfig = Get-OrganizationConfig -ErrorAction Stop
        if ($orgConfig.OrganizationalUnitRoot) {
            $domainName = $orgConfig.OrganizationalUnitRoot
        }
        elseif ($orgConfig.OrganizationId) {
            $domainName = ($orgConfig.OrganizationId -split '/')[1]
        }
    }
    catch {
        $domainName = $null
    }
    
    # Create domain HTML if available
    $domainHtml = if ($domainName) { 
        "<p style='margin-top:6px;font-size:1.05em;font-weight:600'>$domainName</p>" 
    } else { 
        '' 
    }
    
    $reportDate = Get-Date -Format "dd MMMM yyyy - HH:mm:ss"
    
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
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #f8f9fa;
        }
        
        .domain-filter {
            padding: 20px 30px;
            background: #e9ecef;
            border-bottom: 1px solid #dee2e6;
        }
        
        .domain-filter h3 {
            margin-bottom: 10px;
            color: #2a5298;
        }
        
        .domain-tabs {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
        }
        
        .domain-tab {
            padding: 8px 16px;
            background: white;
            border: 2px solid #2a5298;
            border-radius: 5px;
            cursor: pointer;
            transition: all 0.3s ease;
            font-weight: 500;
        }
        
        .domain-tab:hover {
            background: #2a5298;
            color: white;
        }
        
        .domain-tab.active {
            background: #2a5298;
            color: white;
        }
        
        .domain-section {
            display: none;
        }
        
        .domain-section.active {
            display: block;
        }
        
        .domain-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 30px;
            margin: 20px 30px;
            border-radius: 8px;
        }
        
        .domain-header h3 {
            margin: 0;
            font-size: 1.3em;
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
            $domainHtml
            <p class="timestamp">Generated: $reportDate</p>
        </div>
        
        <div class="summary">
            <div class="summary-card total">
                <h3>Total Domains</h3>
                <div class="value">$($AllDomainResults.Count)</div>
            </div>
            <div class="summary-card total">
                <h3>Total Checks</h3>
                <div class="value">$totalAllChecks</div>
            </div>
            <div class="summary-card passed">
                <h3>Passed</h3>
                <div class="value">$totalAllPassed</div>
            </div>
            <div class="summary-card failed">
                <h3>Failed</h3>
                <div class="value">$totalAllFailed</div>
            </div>
            <div class="summary-card compliance">
                <h3>Compliance</h3>
                <div class="value">$overallCompliance%</div>
            </div>
        </div>
        
        <div class="info-section">
            <h2>📋 Report Information</h2>
            <div class="info-grid">
                <div class="info-item">
                    <strong>Total Domains:</strong> $($AllDomainResults.Count)
                </div>
                <div class="info-item">
                    <strong>Domains Checked:</strong> $($AllDomainResults.Domain.Identity -join ', ')
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
"@

    # Generate domain sections
    foreach ($domainResult in $AllDomainResults) {
        $domain = $domainResult.Domain
        $checkResults = $domainResult.CheckResults
        $domainCompliance = $domainResult.CompliancePercentage
        $domainStatusColor = if ($domainCompliance -eq 100) { "#28a745" } else { "#dc3545" }
        
        $html += @"
        <div class="results-section">
            <div style="background: $domainStatusColor; color: white; padding: 15px; margin: 0 0 20px 0; border-radius: 5px;">
                <h2 style="margin: 0; border: none; padding: 0;">🌐 Domain: $($domain.Identity)</h2>
                <p style="margin: 5px 0 0 0; font-size: 0.9em;">Domain Name: $($domain.DomainName) | Compliance: $domainCompliance%</p>
            </div>
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

        foreach ($result in $checkResults) {
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
"@
    }

    $html += @"
        
        <div class="overall-status">
            <h2>Overall Status: $overallStatus</h2>
            <p style="font-size: 1.2em; margin-top: 10px;">
                $totalAllPassed out of $totalAllChecks checks passed across $($AllDomainResults.Count) domain(s)
            </p>
        </div>
        
        <div class="footer">
            <p><strong>Reference:</strong> <a href="https://blueprint.asd.gov.au/configuration/exchange-online/mail-flow/remote-domains/" target="_blank">ASD's Blueprint for Secure Cloud - Remote Domains</a></p>
            <p style="margin-top: 10px;"><strong>Security Controls Explanation:</strong> <a href="https://github.com/directorcia/bp/wiki/Exchange-Online-Remote-Domain-Security-Controls" target="_blank">Why These Recommendations Matter</a></p>
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
    
    # Initialize progress
    Write-Progress -Activity "ASD Remote Domain Check" -Status "Initializing..." -PercentComplete 0
    
    # Get all remote domains
    Write-Progress -Activity "ASD Remote Domain Check" -Status "Retrieving all remote domain configurations..." -PercentComplete 10
    Write-ColorOutput "Retrieving all remote domain configurations..." -Type Info
    
    try {
        $remoteDomains = @(Get-RemoteDomain -ErrorAction Stop)
        
        if (-not $remoteDomains -or $remoteDomains.Count -eq 0) {
            Write-ColorOutput "No remote domains found!" -Type Error
            return
        }
        
        Write-ColorOutput "Found $($remoteDomains.Count) remote domain(s): $($remoteDomains.Identity -join ', ')`n" -Type Success
        
    }
    catch {
        Write-Progress -Activity "ASD Remote Domain Check" -Completed
        Write-ColorOutput "Failed to retrieve remote domains: $($_.Exception.Message)" -Type Error
        return
    }
    
    # Array to store all results for all domains
    $allDomainResults = @()
    
    # Process each remote domain
    $domainIndex = 0
    foreach ($remoteDomain in $remoteDomains) {
        $domainIndex++
        
        Write-ColorOutput "`n========================================" -Type Info
        Write-ColorOutput "Checking Domain $domainIndex of $($remoteDomains.Count): $($remoteDomain.Identity)" -Type Info
        Write-ColorOutput "Domain Name Pattern: $($remoteDomain.DomainName)" -Type Info
        Write-ColorOutput "========================================`n" -Type Info
        
        # Array to store check results for this domain
        $checkResults = @()
        
        # Check each setting
        Write-Progress -Activity "ASD Remote Domain Check" -Status "Checking domain $domainIndex of $($remoteDomains.Count): $($remoteDomain.Identity)" -PercentComplete (10 + ($domainIndex / $remoteDomains.Count * 50))
        
        # Define total checks for progress calculation
        $totalChecks = 10
        $currentCheck = 0
        
        # Domain Name (special handling: only Default must be '*')
        $currentCheck++
        if ($remoteDomain.Identity -eq 'Default') {
            $checkResults += Test-Setting -SettingName "DomainName" `
                -CurrentValue $remoteDomain.DomainName `
                -RequiredValue $Requirements.DomainName `
                -Description "Default remote domain wildcard (should be *)"
        }
        else {
            # For non-default remote domains, the baseline does not mandate '*'; treat as informational PASS
            $checkResults += [pscustomobject]@{
                Setting      = 'DomainName'
                Description  = 'Custom remote domain name (not required to be *)'
                CurrentValue = $remoteDomain.DomainName
                RequiredValue= 'Not enforced'
                Compliant    = $true
                Status       = 'PASS'
            }
        }
        
        # Out of Office Type - Normalize baseline value (ExternalOnly -> External)
        $currentCheck++
        $requiredOOFType = $Requirements.AllowedOOFType
        if ($requiredOOFType -eq "ExternalOnly") {
            $requiredOOFType = "External"
        }
        $checkResults += Test-Setting -SettingName "AllowedOOFType" `
            -CurrentValue $remoteDomain.AllowedOOFType `
            -RequiredValue $requiredOOFType `
            -Description "Out of Office automatic reply types"
        
        # Auto Reply
        $currentCheck++
        $checkResults += Test-Setting -SettingName "AutoReplyEnabled" `
            -CurrentValue $remoteDomain.AutoReplyEnabled `
            -RequiredValue $Requirements.AutoReplyEnabled `
            -Description "Allow automatic replies"
        
        # Auto Forward
        $currentCheck++
        $checkResults += Test-Setting -SettingName "AutoForwardEnabled" `
            -CurrentValue $remoteDomain.AutoForwardEnabled `
            -RequiredValue $Requirements.AutoForwardEnabled `
            -Description "Allow automatic forwarding"
        
        # Delivery Reports
        $currentCheck++
        $checkResults += Test-Setting -SettingName "DeliveryReportEnabled" `
            -CurrentValue $remoteDomain.DeliveryReportEnabled `
            -RequiredValue $Requirements.DeliveryReportEnabled `
            -Description "Allow delivery reports"
        
        # NDR Required (Note: Exchange uses NDREnabled property)
        $currentCheck++
        $checkResults += Test-Setting -SettingName "NDREnabled" `
            -CurrentValue $remoteDomain.NDREnabled `
            -RequiredValue $Requirements.NDRRequired `
            -Description "Send non-delivery reports"
        
        # Meeting Forward Notifications
        $currentCheck++
        $checkResults += Test-Setting -SettingName "MeetingForwardNotificationEnabled" `
            -CurrentValue $remoteDomain.MeetingForwardNotificationEnabled `
            -RequiredValue $Requirements.MeetingForwardNotificationEnabled `
            -Description "Allow meeting forward notifications"
        
        # TNEF (Rich Text Format)
        $currentCheck++
        $checkResults += Test-Setting -SettingName "TNEFEnabled" `
            -CurrentValue $remoteDomain.TNEFEnabled `
            -RequiredValue $Requirements.TNEFEnabled `
            -Description "Use rich-text format (null = Follow user settings)"
        
        # Character Set
        $currentCheck++
        $checkResults += Test-Setting -SettingName "CharacterSet" `
            -CurrentValue $remoteDomain.CharacterSet `
            -RequiredValue $Requirements.CharacterSet `
            -Description "MIME character set"
        
        # Non-MIME Character Set
        $currentCheck++
        $checkResults += Test-Setting -SettingName "NonMIMECharacterSet" `
            -CurrentValue $remoteDomain.NonMIMECharacterSet `
            -RequiredValue $Requirements.NonMIMECharacterSet `
            -Description "Non-MIME character set"
        
        # Display results for this domain
        Write-ColorOutput "`n  CHECK RESULTS - $($remoteDomain.Identity)" -Type Info
        Write-ColorOutput "  ========================================`n" -Type Info
        
        foreach ($result in $checkResults) {
            $statusColor = if ($result.Compliant) { "Success" } else { "Error" }
            $statusSymbol = if ($result.Compliant) { "[✓]" } else { "[✗]" }
            
            Write-ColorOutput "  $statusSymbol $($result.Setting)" -Type $statusColor
            Write-Host "      Description : $($result.Description)"
            Write-Host "      Current     : $($result.CurrentValue)"
            Write-Host "      Required    : $($result.RequiredValue)"
            Write-Host "      Status      : $($result.Status)"
            Write-Host ""
        }
        
        # Summary for this domain
        $totalChecks = $checkResults.Count
        $passedChecks = ($checkResults | Where-Object { $_.Compliant }).Count
        $failedChecks = $totalChecks - $passedChecks
        $compliancePercentage = [math]::Round(($passedChecks / $totalChecks) * 100, 2)
        
        Write-ColorOutput "  ========================================" -Type Info
        Write-ColorOutput "  SUMMARY - $($remoteDomain.Identity)" -Type Info
        Write-ColorOutput "  ========================================" -Type Info
        Write-Host "  Total Checks    : $totalChecks"
        Write-ColorOutput "  Passed          : $passedChecks" -Type Success
        
        if ($failedChecks -gt 0) {
            Write-ColorOutput "  Failed          : $failedChecks" -Type Error
        } else {
            Write-ColorOutput "  Failed          : $failedChecks" -Type Success
        }
        
        Write-Host "  Compliance      : $compliancePercentage%"
        
        if ($compliancePercentage -eq 100) {
            Write-ColorOutput "`n  Status          : COMPLIANT ✓" -Type Success
        } else {
            Write-ColorOutput "`n  Status          : NON-COMPLIANT ✗" -Type Error
        }
        
        Write-ColorOutput "  ========================================`n" -Type Info
        
        # Store domain results
        $allDomainResults += @{
            Domain = $remoteDomain
            CheckResults = $checkResults
            TotalChecks = $totalChecks
            PassedChecks = $passedChecks
            FailedChecks = $failedChecks
            CompliancePercentage = $compliancePercentage
        }
    }
    
    # Overall summary
    Write-Progress -Activity "ASD Remote Domain Check" -Status "Generating overall summary..." -PercentComplete 60
    Write-ColorOutput "`n========================================" -Type Info
    Write-ColorOutput "  OVERALL SUMMARY (ALL DOMAINS)" -Type Info
    Write-ColorOutput "========================================" -Type Info
    Write-Host "Total Domains   : $($remoteDomains.Count)"
    
    $totalAllChecks = ($allDomainResults | ForEach-Object { $_.TotalChecks } | Measure-Object -Sum).Sum
    $totalAllPassed = ($allDomainResults | ForEach-Object { $_.PassedChecks } | Measure-Object -Sum).Sum
    $totalAllFailed = $totalAllChecks - $totalAllPassed
    $overallCompliance = if ($totalAllChecks -gt 0) { [math]::Round(($totalAllPassed / $totalAllChecks) * 100, 2) } else { 0 }
    
    Write-Host "Total Checks    : $totalAllChecks"
    Write-ColorOutput "Passed          : $totalAllPassed" -Type Success
    
    if ($totalAllFailed -gt 0) {
        Write-ColorOutput "Failed          : $totalAllFailed" -Type Error
    } else {
        Write-ColorOutput "Failed          : $totalAllFailed" -Type Success
    }
    
    Write-Host "Compliance      : $overallCompliance%"
    
    if ($overallCompliance -eq 100) {
        Write-ColorOutput "`nOverall Status  : COMPLIANT ✓" -Type Success
    } else {
        Write-ColorOutput "`nOverall Status  : NON-COMPLIANT ✗" -Type Error
    }
    
    Write-ColorOutput "========================================`n" -Type Info
    
    # Export to CSV if requested
    if ($ExportToCSV) {
        Write-Progress -Activity "ASD Remote Domain Check" -Status "Exporting results to CSV..." -PercentComplete 70
        try {
            $csvData = @()
            foreach ($domainResult in $allDomainResults) {
                foreach ($check in $domainResult.CheckResults) {
                    $csvData += [PSCustomObject]@{
                        Domain = $domainResult.Domain.Identity
                        DomainName = $domainResult.Domain.DomainName
                        Setting = $check.Setting
                        Description = $check.Description
                        CurrentValue = $check.CurrentValue
                        RequiredValue = $check.RequiredValue
                        Status = $check.Status
                    }
                }
            }
            $csvData | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8
            Write-ColorOutput "Results exported to: $CSVPath" -Type Success
        }
        catch {
            Write-ColorOutput "Failed to export results: $($_.Exception.Message)" -Type Error
        }
    }
    
    # Generate HTML Report (always) - now with all domains
    Write-Progress -Activity "ASD Remote Domain Check" -Status "Generating HTML report..." -PercentComplete 80
    Write-ColorOutput "`nGenerating HTML report..." -Type Info
    if (New-HTMLReport -AllDomainResults $allDomainResults -OutputPath $script:HTMLPath) {
        Write-ColorOutput "HTML report generated: $script:HTMLPath" -Type Success
        Write-Progress -Activity "ASD Remote Domain Check" -Status "Opening report in browser..." -PercentComplete 90
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
    
    # Complete progress
    Write-Progress -Activity "ASD Remote Domain Check" -Status "Completed" -PercentComplete 100
    Start-Sleep -Milliseconds 500
    Write-Progress -Activity "ASD Remote Domain Check" -Completed
    
    return $allDomainResults
}

# Main execution
try {
    # Initialize logging if enabled
    if ($script:DetailedLogging) {
        Write-Log "=== ASD Remote Domain Configuration Check Started ===" -Level "INFO"
        Write-Log "Script Version: $scriptVersion" -Level "INFO"
        Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)" -Level "INFO"
        Write-Log "Detailed Logging: Enabled" -Level "INFO"
        Write-Log "Log Path: $script:LogPath" -Level "INFO"
    }
    
    Write-ColorOutput "`n========================================" -Type Info
    Write-ColorOutput "  ASD Remote Domain Configuration Check" -Type Info
    Write-ColorOutput "========================================" -Type Info
    $isUrl = $script:BaselinePath -match '^https?://'
    if ($isUrl) {
        Write-ColorOutput "Baseline: GitHub (latest)" -Type Info
    }
    elseif (Test-Path $script:BaselinePath) {
        Write-ColorOutput "Baseline: Local File (found)" -Type Success
    }
    else {
        Write-ColorOutput "Baseline: Local File (not found - will use defaults)" -Type Warning
    }
    Write-ColorOutput "Location: $script:BaselinePath" -Type Info
    Write-ColorOutput "Output:   $parentPath`n" -Type Info
    
    if ($script:DetailedLogging) {
        Write-ColorOutput "Logging:  $script:LogPath" -Type Info
    }
    
    # Load baseline settings
    $asdRequirements = Get-BaselineSettings -Path $BaselinePath
    
    if (-not $asdRequirements) {
        Write-ColorOutput "`nFailed to load baseline settings. Cannot proceed." -Type Error
        exit 1
    }
    
    # Check module
    if (-not (Test-ExchangeModule)) {
        Write-ColorOutput "`nExchangeOnlineManagement module is required. Please install it first." -Type Error
        exit 1
    }
    
    # Connect to Exchange Online
    if (-not (Connect-EXO)) {
        Write-ColorOutput "`nFailed to connect to Exchange Online. Cannot proceed." -Type Error
        exit 1
    }
    
    # Validate permissions before proceeding
    if (-not (Test-ExchangePermissions)) {
        Write-ColorOutput "`nScript cannot continue without proper permissions." -Type Error
        exit 1
    }
    
    # Run the check
    Write-Log "Starting remote domain compliance check" -Level "INFO"
    $results = Invoke-RemoteDomainCheck -Requirements $asdRequirements
    
    if ($results) {
        Write-Log "Script completed successfully" -Level "INFO"
        Write-ColorOutput "`nScript completed successfully." -Type Success
        
        if ($script:DetailedLogging) {
            Write-ColorOutput "Detailed log saved to: $script:LogPath" -Type Info
        }
    }
    else {
        Write-Log "Script completed with warnings" -Level "WARN"
        Write-ColorOutput "`nScript completed with warnings. Please review the results." -Type Warning
    }
    
    # Log final summary
    if ($script:DetailedLogging) {
        Write-Log "=== ASD Remote Domain Configuration Check Completed ===" -Level "INFO"
        Write-Log "HTML Report: $script:HTMLPath" -Level "INFO"
        if ($ExportToCSV) {
            Write-Log "CSV Export: $CSVPath" -Level "INFO"
        }
    }
}
catch {
    Write-Log "SCRIPT EXECUTION FAILED: $($_.Exception.Message)" -Level "ERROR"
    Write-Log "Error Location: Line $($_.InvocationInfo.ScriptLineNumber) - $($_.InvocationInfo.Line.Trim())" -Level "ERROR"
    if ($_.ScriptStackTrace) {
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
    }
    
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
    if ($_.ScriptStackTrace) {
        Write-ColorOutput "Stack Trace:" -Type Warning
        Write-Host "  $($_.ScriptStackTrace)"
    }
    Write-Host ""
    Write-ColorOutput "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -Type Error
    
    if ($script:DetailedLogging) {
        Write-Host ""
        Write-ColorOutput "Detailed error log saved to: $script:LogPath" -Type Info
    }
    
    exit 1
}
