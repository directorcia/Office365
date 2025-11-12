<#
.SYNOPSIS
    Check Exchange Online Mail Flow settings against ASD Blueprint requirements

.DESCRIPTION
    This script checks the Exchange Online Mail Flow configuration against 
    ASD's Blueprint for Secure Cloud requirements. It validates the organization 
    config settings including plus addressing, alias sending, SMTP authentication,
    legacy TLS clients, and message recall.
    
    It also validates transport config settings including reply all storm protection.
    
    Reference: https://blueprint.asd.gov.au/configuration/exchange-online/settings/mail-flow/

.EXAMPLE
    .\asd-mailflow-get.ps1
    
    Connects to Exchange Online, downloads latest baseline from GitHub, checks mail flow 
    settings, and automatically generates an HTML report in the script directory that opens 
    in the default browser

.EXAMPLE
    .\asd-mailflow-get.ps1 -ExportToCSV
    
    Runs the check with both HTML report (automatic) and CSV export.
    Downloads latest baseline from GitHub. All files created in parent directory.

.EXAMPLE
    .\asd-mailflow-get.ps1 -BaselinePath "C:\Baselines\prod-mailflow.json"
    
    Uses a custom baseline JSON file for the compliance check

.EXAMPLE
    .\asd-mailflow-get.ps1 -BaselinePath ".\baselines\dev-environment.json" -ExportToCSV
    
    Uses a development environment baseline and exports results to CSV

.EXAMPLE
    .\asd-mailflow-get.ps1 -CSVPath "C:\Reports\custom-report.csv" -ExportToCSV
    
    Exports CSV to a custom location instead of the default parent directory

.EXAMPLE
    .\asd-mailflow-get.ps1 -DetailedLogging
    
    Runs the check with detailed logging enabled. Log file created in parent directory

.EXAMPLE
    .\asd-mailflow-get.ps1 -DetailedLogging -LogPath "C:\Logs\custom-log.log"
    
    Runs the check with detailed logging to a custom log file location

.NOTES
    Author: CIAOPS
    Date: 11-12-2025
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
    2. GitHub (default): https://raw.githubusercontent.com/directorcia/bp/main/ASD/Exchange-Online/Settings/mailflow.json
    3. Built-in ASD Blueprint defaults (fallback if GitHub unavailable)
    
    File Locations (Default):
    - HTML Report: {parent-directory}\asd-mailflow-get-{timestamp}.html
    - CSV Export: {parent-directory}\asd-mailflow-get-{timestamp}.csv
    - Log File (if enabled): {parent-directory}\asd-mailflow-get-{timestamp}.log

.LINK
    https://github.com/directorcia/office365
    https://github.com/directorcia/Office365/wiki/ASD-Mail-Flow-Configuration-Check - Documentation
    https://github.com/directorcia/bp/wiki/Exchange-Online-Mail-Flow-Security-Controls - Exchange Online Mail Flow Security Controls
    
    https://blueprint.asd.gov.au/configuration/exchange-online/settings/mail-flow/
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
    $CSVPath = Join-Path $parentPath "asd-mailflow-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
}

# Set default log path if detailed logging is enabled
if ($DetailedLogging -and -not $LogPath) {
    $LogPath = Join-Path $parentPath "asd-mailflow-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
}

# Default GitHub URL for baseline settings
$defaultGitHubURL = "https://raw.githubusercontent.com/directorcia/bp/main/ASD/Exchange-Online/Settings/mailflow.json"

# Set default baseline path if not provided (in parent directory)
if (-not $BaselinePath) {
    $BaselinePath = $defaultGitHubURL
}

# Script-scope variables for tracking state
$script:BaselinePath = $BaselinePath
$script:baselineLoaded = $false
$script:HTMLPath = Join-Path $parentPath "asd-mailflow-get-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
$script:LogPath = $LogPath
$script:DetailedLogging = $DetailedLogging

# Script variables
$scriptVersion = "1.0"
$scriptName = "ASD Mail Flow Settings Check"

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
    
    $requiredFields = @(
        @{Path = 'general'; Type = 'Object'; Description = 'General mail flow settings'},
        @{Path = 'security'; Type = 'Object'; Description = 'Security settings'},
        @{Path = 'replyAllStormProtection'; Type = 'Object'; Description = 'Reply all storm protection configuration'},
        @{Path = 'messageRecall'; Type = 'Object'; Description = 'Message recall configuration'}
    )
    
    # Note: 'metadata' is optional and not validated as a required field
    
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
    param(
        [string]$Path
    )
    
    Write-Log "Starting baseline settings load from: $Path" -Level "INFO"
    
    $baselineSettings = $null
    $script:baselineLoaded = $false
    $jsonContent = $null

    # Check if Path is a URL
    $isUrl = $Path -match '^https?://'
    Write-Log "Baseline source type: $(if ($isUrl) { 'URL' } else { 'Local File' })" -Level "INFO"
    
    if ($isUrl) {
        try {
            Write-Progress -Activity "Loading Baseline" -Status "Downloading from GitHub..." -PercentComplete 30
            Write-ColorOutput "Downloading baseline settings from: $Path" -Type Info
            $jsonContent = (Invoke-WebRequest -Uri $Path -UseBasicParsing -ErrorAction Stop).Content
            Write-Progress -Activity "Loading Baseline" -Status "Parsing JSON..." -PercentComplete 60
            $json = $jsonContent | ConvertFrom-Json -ErrorAction Stop
            
            # Validate schema
            Write-Progress -Activity "Loading Baseline" -Status "Validating schema..." -PercentComplete 80
            if (Test-BaselineSchema -Baseline $json) {
                $baselineSettings = $json
                $script:baselineLoaded = $true
                Write-Progress -Activity "Loading Baseline" -Completed
                Write-Log "Baseline loaded successfully from GitHub" -Level "INFO"
                Write-ColorOutput "✓ Baseline loaded successfully from GitHub.`n" -Type Success
            }
            else {
                Write-Progress -Activity "Loading Baseline" -Completed
                Write-Log "Baseline schema validation failed" -Level "ERROR"
                $baselineSettings = $null
            }
        }
        catch {
            Write-Progress -Activity "Loading Baseline" -Completed
            Write-Log "Failed to download baseline from URL: $($_.Exception.Message)" -Level "ERROR"
            Write-ColorOutput "Failed to download or parse baseline from URL: $($_.Exception.Message)" -Type Error
            Write-ColorOutput "⚠️  Using built-in ASD Blueprint defaults instead`n" -Type Warning
            $baselineSettings = $null
        }
    }
    elseif (Test-Path $Path) {
        try {
            Write-Progress -Activity "Loading Baseline" -Status "Reading local file..." -PercentComplete 30
            Write-ColorOutput "Loading baseline settings from: $Path" -Type Info
            $json = Get-Content -Path $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            
            # Validate schema
            Write-Progress -Activity "Loading Baseline" -Status "Validating schema..." -PercentComplete 70
            if (Test-BaselineSchema -Baseline $json) {
                $baselineSettings = $json
                $script:baselineLoaded = $true
                Write-Progress -Activity "Loading Baseline" -Completed
                Write-Log "Baseline loaded successfully from local file" -Level "INFO"
                Write-ColorOutput "✓ Baseline loaded successfully from JSON file.`n" -Type Success
            }
            else {
                Write-Progress -Activity "Loading Baseline" -Completed
                Write-Log "Baseline schema validation failed for local file" -Level "ERROR"
                $baselineSettings = $null
            }
        }
        catch {
            Write-Progress -Activity "Loading Baseline" -Completed
            Write-Log "Failed to parse baseline JSON: $($_.Exception.Message)" -Level "ERROR"
            Write-ColorOutput "Failed to parse baseline JSON: $($_.Exception.Message)" -Type Error
            Write-ColorOutput "Error at line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())" -Type Error
            Write-ColorOutput "⚠️  Using built-in ASD Blueprint defaults instead`n" -Type Warning
            $baselineSettings = $null
        }
    }
    else {
        Write-Progress -Activity "Loading Baseline" -Completed
        Write-Log "Baseline file not found at: $Path - using defaults" -Level "WARN"
        Write-ColorOutput "Baseline file not found at: $Path" -Type Warning
        Write-ColorOutput "⚠️  Using built-in ASD Blueprint defaults instead`n" -Type Warning
    }

    # Build and return ASD Blueprint requirements (from baseline if available, otherwise defaults)
    return @{
        # General settings
        DisablePlusAddressInRecipients = (Get-BaselineValue -Parent $baselineSettings.general -Property 'disablePlusAddressInRecipients' -Default $true)
        SendFromAliasEnabled = (Get-BaselineValue -Parent $baselineSettings.general -Property 'sendFromAliasEnabled' -Default $false)
        
        # Security settings
        SmtpAuthProtocolEnabled = (Get-BaselineValue -Parent $baselineSettings.security -Property 'smtpAuthProtocolEnabled' -Default $false)
        LegacyTlsClientsAllowed = (Get-BaselineValue -Parent $baselineSettings.security -Property 'legacyTlsClientsAllowed' -Default $false)
        
        # Reply all storm protection
        ReplyAllStormEnabled = (Get-BaselineValue -Parent $baselineSettings.replyAllStormProtection -Property 'enabled' -Default $true)
        ReplyAllStormMinimumRecipients = (Get-BaselineValue -Parent $baselineSettings.replyAllStormProtection -Property 'minimumRecipients' -Default 2500)
        ReplyAllStormMinimumReplyAlls = (Get-BaselineValue -Parent $baselineSettings.replyAllStormProtection -Property 'minimumReplyAlls' -Default 10)
        ReplyAllStormBlockDurationHours = (Get-BaselineValue -Parent $baselineSettings.replyAllStormProtection -Property 'blockDurationHours' -Default 6)
        
        # Message recall
        MessageRecallEnabled = (Get-BaselineValue -Parent $baselineSettings.messageRecall -Property 'enabled' -Default $true)
        MessageRecallAllowRecallReadMessages = (Get-BaselineValue -Parent $baselineSettings.messageRecall -Property 'allowRecallReadMessages' -Default $true)
        MessageRecallEnableRecipientAlerts = (Get-BaselineValue -Parent $baselineSettings.messageRecall -Property 'enableRecipientAlerts' -Default $true)
        MessageRecallAlertReadMessagesOnly = (Get-BaselineValue -Parent $baselineSettings.messageRecall -Property 'alertReadMessagesOnly' -Default $true)
        MessageRecallMaxAgeDays = (Get-BaselineValue -Parent $baselineSettings.messageRecall -Property 'maxAgeDays' -Default 1)
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
    
    $result = @{
        Setting = $SettingName
        Description = $Description
        CurrentValue = if ($null -eq $CurrentValue) { "Not set" } else { $CurrentValue.ToString() }
        RequiredValue = if ($null -eq $RequiredValue) { "Not set" } else { $RequiredValue.ToString() }
        Compliant = $false
        Status = ""
    }
    
    # Compare values
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
    
    # Log the check result
    $logMsg = "Check: $SettingName - Current: $($result.CurrentValue), Required: $($result.RequiredValue), Status: $($result.Status)"
    Write-Log -Message $logMsg -Level $(if ($result.Compliant) { "INFO" } else { "WARN" })
    
    return $result
}

# Generate HTML Report
function New-HTMLReport {
    param(
        [array]$CheckResults,
        [object]$OrgConfig,
        [string]$OutputPath
    )
    
    $totalChecks = $CheckResults.Count
    $passedChecks = ($CheckResults | Where-Object { $_.Compliant }).Count
    $failedChecks = $totalChecks - $passedChecks
    $compliancePercentage = [math]::Round(($passedChecks / $totalChecks) * 100, 2)
    $overallStatus = if ($compliancePercentage -eq 100) { "COMPLIANT" } else { "NON-COMPLIANT" }
    $statusColor = if ($compliancePercentage -eq 100) { "#28a745" } else { "#dc3545" }
    
    $reportDate = Get-Date -Format "dd MMMM yyyy - HH:mm:ss"
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ASD Mail Flow Settings Compliance Report</title>
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
            <h1>🛡️ ASD Mail Flow Settings Compliance Report</h1>
            <p>Exchange Online Mail Flow Configuration Check</p>
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
            <h2>📋 Organization Information</h2>
            <div class="info-grid">
                <div class="info-item">
                    <strong>Organization:</strong> $($OrgConfig.DisplayName)
                </div>
                <div class="info-item">
                    <strong>Identity:</strong> $($OrgConfig.Identity)
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
            <p><strong>Reference:</strong> <a href="https://blueprint.asd.gov.au/configuration/exchange-online/settings/mail-flow/" target="_blank">ASD's Blueprint for Secure Cloud - Mail Flow Settings</a></p>
            <p style="margin-top: 10px;"><strong>Security Controls Explanation:</strong> <a href="https://github.com/directorcia/bp/wiki/Exchange-Online-Mail-Flow-Security-Controls" target="_blank">Why These Recommendations Matter</a></p>
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
function Invoke-MailFlowCheck {
    param(
        [hashtable]$Requirements
    )
    
    Write-ColorOutput "`n========================================" -Type Info
    Write-ColorOutput "  $scriptName v$scriptVersion" -Type Info
    Write-ColorOutput "  ASD Blueprint Compliance Check" -Type Info
    Write-ColorOutput "========================================`n" -Type Info
    
    # Initialize progress
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Initializing..." -PercentComplete 0
    
    # Get the organization configuration
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Retrieving organization configuration..." -PercentComplete 10
    Write-ColorOutput "Retrieving organization mail flow configuration..." -Type Info
    
    try {
        $orgConfig = Get-OrganizationConfig -ErrorAction Stop
        
        if (-not $orgConfig) {
            Write-ColorOutput "Organization configuration not found!" -Type Error
            return
        }
        
        Write-ColorOutput "Organization configuration retrieved: $($orgConfig.DisplayName)" -Type Success
        
        # Get transport configuration for Reply-All Storm Protection settings
        Write-ColorOutput "Retrieving transport configuration..." -Type Info
        $transportConfig = Get-TransportConfig -ErrorAction Stop
        
        if (-not $transportConfig) {
            Write-ColorOutput "Transport configuration not found!" -Type Error
            return
        }
        
        Write-ColorOutput "Transport configuration retrieved.`n" -Type Success
        
    }
    catch {
        Write-Progress -Activity "ASD Mail Flow Settings Check" -Completed
        Write-ColorOutput "Failed to retrieve configuration: $($_.Exception.Message)" -Type Error
        return
    }
    
    # Array to store all check results
    $checkResults = @()
    
    # Check each setting
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking settings against ASD Blueprint requirements..." -PercentComplete 20
    Write-ColorOutput "Checking settings against ASD Blueprint requirements...`n" -Type Info
    
    # Define total checks for progress calculation
    $totalChecks = 13
    $currentCheck = 0
    
    # General Settings - Plus Addressing
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking PlusAddressingEnabled ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # Treat null as False for DisablePlusAddressInRecipients (meaning plus addressing is enabled by default)
    $disablePlusAddressValue = if ($null -eq $orgConfig.DisablePlusAddressInRecipients) { $false } else { $orgConfig.DisablePlusAddressInRecipients }
    $checkResults += Test-Setting -SettingName "DisablePlusAddressInRecipients" `
        -CurrentValue $disablePlusAddressValue `
        -RequiredValue $Requirements.DisablePlusAddressInRecipients `
        -Description "Disable plus addressing (user+tag@domain.com) - ASD recommends disabling"
    
    # General Settings - Send From Aliases
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking SendFromAliasEnabled ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # Treat null as False for boolean settings
    $sendFromAliasValue = if ($null -eq $orgConfig.SendFromAliasEnabled) { $false } else { $orgConfig.SendFromAliasEnabled }
    $checkResults += Test-Setting -SettingName "SendFromAliasEnabled" `
        -CurrentValue $sendFromAliasValue `
        -RequiredValue $Requirements.SendFromAliasEnabled `
        -Description "Allow sending from email aliases"
    
    # Security Settings - SMTP Auth
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking SmtpClientAuthenticationDisabled ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # Treat null as True for SmtpClientAuthenticationDisabled (default is disabled)
    $smtpAuthDisabledValue = if ($null -eq $orgConfig.SmtpClientAuthenticationDisabled) { $true } else { $orgConfig.SmtpClientAuthenticationDisabled }
    $checkResults += Test-Setting -SettingName "SmtpClientAuthenticationDisabled" `
        -CurrentValue $smtpAuthDisabledValue `
        -RequiredValue (-not $Requirements.SmtpAuthProtocolEnabled) `
        -Description "SMTP client authentication (should be disabled)"
    
    # Security Settings - Legacy TLS
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking AllowLegacyTLSClients ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # Treat null as False for boolean settings
    $allowLegacyTlsValue = if ($null -eq $orgConfig.AllowLegacyTLSClients) { $false } else { $orgConfig.AllowLegacyTLSClients }
    $checkResults += Test-Setting -SettingName "AllowLegacyTLSClients" `
        -CurrentValue $allowLegacyTlsValue `
        -RequiredValue $Requirements.LegacyTlsClientsAllowed `
        -Description "Allow legacy TLS clients (TLS 1.0/1.1)"
    
    # Reply All Storm Protection - Enabled
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking ReplyAllStormProtectionEnabled ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # Treat null as False for boolean settings (default is disabled)
    $replyAllStormEnabledValue = if ($null -eq $transportConfig.ReplyAllStormProtectionEnabled) { $false } else { $transportConfig.ReplyAllStormProtectionEnabled }
    $checkResults += Test-Setting -SettingName "ReplyAllStormProtectionEnabled" `
        -CurrentValue $replyAllStormEnabledValue `
        -RequiredValue $Requirements.ReplyAllStormEnabled `
        -Description "Enable reply all storm protection"
    
    # Reply All Storm Protection - Minimum Recipients
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking ReplyAllStormDetectionMinimumRecipients ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # Treat null as 0 for numeric settings
    $replyAllStormMinRecipientsValue = if ($null -eq $transportConfig.ReplyAllStormDetectionMinimumRecipients) { 0 } else { $transportConfig.ReplyAllStormDetectionMinimumRecipients }
    $checkResults += Test-Setting -SettingName "ReplyAllStormProtectionMinimumRecipients" `
        -CurrentValue $replyAllStormMinRecipientsValue `
        -RequiredValue $Requirements.ReplyAllStormMinimumRecipients `
        -Description "Minimum recipients to trigger protection"
    
    # Reply All Storm Protection - Minimum Reply Alls
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking ReplyAllStormDetectionMinimumReplies ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # Treat null as 0 for numeric settings
    $replyAllStormMinRepliesValue = if ($null -eq $transportConfig.ReplyAllStormDetectionMinimumReplies) { 0 } else { $transportConfig.ReplyAllStormDetectionMinimumReplies }
    $checkResults += Test-Setting -SettingName "ReplyAllStormProtectionMinimumReplies" `
        -CurrentValue $replyAllStormMinRepliesValue `
        -RequiredValue $Requirements.ReplyAllStormMinimumReplyAlls `
        -Description "Minimum reply-alls to trigger protection"
    
    # Reply All Storm Protection - Block Duration
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking ReplyAllStormBlockDurationHours ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # Treat null as 0 for numeric settings
    $replyAllStormBlockDurationValue = if ($null -eq $transportConfig.ReplyAllStormBlockDurationHours) { 0 } else { $transportConfig.ReplyAllStormBlockDurationHours }
    $checkResults += Test-Setting -SettingName "ReplyAllStormBlockDurationHours" `
        -CurrentValue $replyAllStormBlockDurationValue `
        -RequiredValue $Requirements.ReplyAllStormBlockDurationHours `
        -Description "Block duration in hours"
    
    # Message Recall - Enabled
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking MessageRecallEnabled ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # SPECIAL CASE: For Message Recall, null/blank = TRUE (Microsoft documentation states this)
    # Per https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/work-with-cloud-based-message-recall
    $messageRecallEnabledValue = if ([string]::IsNullOrWhiteSpace($orgConfig.MessageRecallEnabled)) { $true } else { $orgConfig.MessageRecallEnabled }
    $checkResults += Test-Setting -SettingName "MessageRecallEnabled" `
        -CurrentValue $messageRecallEnabledValue `
        -RequiredValue $Requirements.MessageRecallEnabled `
        -Description "Enable message recall"
    
    # Message Recall - Allow Recall Read Messages
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking RecallReadMessagesEnabled ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # SPECIAL CASE: For Message Recall, null/blank = TRUE (Microsoft documentation states this)
    $recallReadMessagesValue = if ([string]::IsNullOrWhiteSpace($orgConfig.RecallReadMessagesEnabled)) { $true } else { $orgConfig.RecallReadMessagesEnabled }
    $checkResults += Test-Setting -SettingName "RecallReadMessagesEnabled" `
        -CurrentValue $recallReadMessagesValue `
        -RequiredValue $Requirements.MessageRecallAllowRecallReadMessages `
        -Description "Allow recalling already-read messages"
    
    # Message Recall - Enable Recipient Alerts
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking MessageRecallAlertRecipientsEnabled ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # Treat null as False for boolean settings
    $messageRecallAlertRecipientsValue = if ($null -eq $orgConfig.MessageRecallAlertRecipientsEnabled) { $false } else { $orgConfig.MessageRecallAlertRecipientsEnabled }
    $checkResults += Test-Setting -SettingName "MessageRecallAlertRecipientsEnabled" `
        -CurrentValue $messageRecallAlertRecipientsValue `
        -RequiredValue $Requirements.MessageRecallEnableRecipientAlerts `
        -Description "Enable recipient alerts for recalls"
    
    # Message Recall - Alert Read Messages Only
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking MessageRecallAlertRecipientsReadMessagesOnlyEnabled ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # Treat null as False for boolean settings
    $messageRecallAlertReadOnlyValue = if ($null -eq $orgConfig.MessageRecallAlertRecipientsReadMessagesOnlyEnabled) { $false } else { $orgConfig.MessageRecallAlertRecipientsReadMessagesOnlyEnabled }
    # Add dependency note if parent setting is disabled
    $alertReadOnlyDescription = "Alert recipients only for read messages"
    if (-not $messageRecallAlertRecipientsValue) {
        $alertReadOnlyDescription += " (Requires MessageRecallAlertRecipientsEnabled to be True first)"
    }
    $checkResults += Test-Setting -SettingName "MessageRecallAlertRecipientsReadMessagesOnlyEnabled" `
        -CurrentValue $messageRecallAlertReadOnlyValue `
        -RequiredValue $Requirements.MessageRecallAlertReadMessagesOnly `
        -Description $alertReadOnlyDescription
    
    # Message Recall - Max Age Days
    $currentCheck++
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Checking MessageRecallMaxRecallableAge ($currentCheck of $totalChecks)" -PercentComplete (20 + ($currentCheck / $totalChecks * 40))
    # MessageRecallMaxRecallableAge is a TimeSpan string like "365.00:00:00"
    # We need to convert it to days for comparison
    $maxRecallAge = $orgConfig.MessageRecallMaxRecallableAge
    if ([string]::IsNullOrWhiteSpace($maxRecallAge)) {
        $messageRecallMaxAgeValue = 0
    } else {
        # Extract days from TimeSpan format "days.HH:mm:ss"
        if ($maxRecallAge -match '^\d+') {
            $messageRecallMaxAgeValue = [int]($matches[0])
        } else {
            $messageRecallMaxAgeValue = 0
        }
    }
    $checkResults += Test-Setting -SettingName "MessageRecallMaxRecallableAge" `
        -CurrentValue $messageRecallMaxAgeValue `
        -RequiredValue $Requirements.MessageRecallMaxAgeDays `
        -Description "Maximum age of messages that can be recalled (days)"
    
    # Display results
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Analyzing results..." -PercentComplete 60
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
        Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Exporting results to CSV..." -PercentComplete 70
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
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Generating HTML report..." -PercentComplete 80
    Write-ColorOutput "`nGenerating HTML report..." -Type Info
    if (New-HTMLReport -CheckResults $checkResults -OrgConfig $orgConfig -OutputPath $script:HTMLPath) {
        Write-ColorOutput "HTML report generated: $script:HTMLPath" -Type Success
        Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Opening report in browser..." -PercentComplete 90
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
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Status "Completed" -PercentComplete 100
    Start-Sleep -Milliseconds 500
    Write-Progress -Activity "ASD Mail Flow Settings Check" -Completed
    
    return $checkResults
}

# Main execution
try {
    # Initialize logging if enabled
    if ($script:DetailedLogging) {
        Write-Log "=== ASD Mail Flow Settings Check Started ===" -Level "INFO"
        Write-Log "Script Version: $scriptVersion" -Level "INFO"
        Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)" -Level "INFO"
        Write-Log "Detailed Logging: Enabled" -Level "INFO"
        Write-Log "Log Path: $script:LogPath" -Level "INFO"
    }
    
    Write-ColorOutput "`n========================================" -Type Info
    Write-ColorOutput "  ASD Mail Flow Settings Check" -Type Info
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
    Write-Log "Starting mail flow settings compliance check" -Level "INFO"
    $results = Invoke-MailFlowCheck -Requirements $asdRequirements
    
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
        Write-Log "=== ASD Mail Flow Settings Check Completed ===" -Level "INFO"
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
