#Requires -Version 5.1
<#
.SYNOPSIS
    Tests connection speed and latency to Microsoft 365 services
    
.DESCRIPTION
    This script tests network connectivity, DNS resolution times, and HTTP response times 
    to various Microsoft 365 service endpoints. It provides comprehensive network 
    performance metrics for M365 services including Exchange Online, SharePoint Online,
    Teams, OneDrive, and other core services.
    
.PARAMETER TestDuration
    Duration in seconds for bandwidth testing (default: 10)
    
.PARAMETER IncludeBandwidthTest
    Include bandwidth testing using Microsoft's network test endpoints
    
.PARAMETER OutputFormat
    Output format: Console, CSV, JSON, or HTML (default: HTML)
    
.PARAMETER OutputPath
    Path to save results when using CSV or JSON output
    
.PARAMETER DetailedLogging
    Enable detailed logging to file

.PARAMETER IncludeAuthentication
    Include authentication tests for M365 services (requires interactive login)

.PARAMETER SkipAuthenticationCheck
    Skip checking for existing authentication tokens

.PARAMETER UploadResults
    Upload test results to CIAOPS secure Azure Blob Storage for analysis and support (unlimited uploads)

.PARAMETER ShowUploadStats
    Display current upload statistics

.EXAMPLE
    .\test-m365-connection-speed.ps1
    
.EXAMPLE
    .\test-m365-connection-speed.ps1 -IncludeAuthentication -DetailedLogging
    
.EXAMPLE
    .\test-m365-connection-speed.ps1 -IncludeBandwidthTest -OutputFormat CSV -OutputPath "c:\temp\m365-test-results.csv"
    
.EXAMPLE
    .\test-m365-connection-speed.ps1 -IncludeBandwidthTest -OutputFormat HTML
    
.EXAMPLE
    .\test-m365-connection-speed.ps1 -TestDuration 30 -Verbose

.EXAMPLE
    .\test-m365-connection-speed.ps1 -UploadResults
    
.EXAMPLE
    .\test-m365-connection-speed.ps1 -ShowUploadStats

.NOTES
    Author: CIAOPS
    Version: 1.0
    Created: July 2025
    
    This script follows Azure/M365 best practices for network diagnostics and monitoring.
    It implements proper error handling, logging, and performance measurement techniques.
    
.LINK
    https://docs.microsoft.com/en-us/microsoft-365/enterprise/network-connectivity
    
.LINK
    https://docs.microsoft.com/en-us/microsoft-365/enterprise/office-365-network-connectivity-principles

.LINK
    https://github.com/directorcia/Office365/wiki/Microsoft-365-Connection-Speed-Test
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateRange(5, 300)]
    [int]$TestDuration = 10,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeBandwidthTest,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Console", "CSV", "JSON", "HTML")]
    [string]$OutputFormat = "HTML",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$DetailedLogging,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeAuthentication,
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipAuthenticationCheck,
    
    [Parameter(Mandatory = $false)]
    [switch]$UploadResults,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowUploadStats
)

# Error handling and logging setup
$ErrorActionPreference = "Continue"
$WarningPreference = "Continue"

# Azure Blob Storage configuration for CIAOPS upload service
$AzureUploadConfig = @{
    SasUrl = "https://m365testresults.blob.core.windows.net/m365-metric-uploads?sp=cw&st=2025-07-21T22:44:29Z&se=2027-03-01T05:59:29Z&spr=https&sv=2024-11-04&sr=c&sig=VW2DQRwcublyIkByKae7xgcgQwk7SEVbcD3NngPR4bg%3D"
    ContainerName = "m365-metric-uploads"
    StorageAccount = "m365testresults"
}

# Initialize logging
if ($DetailedLogging) {
    $LogPath = Join-Path $env:TEMP "m365-connection-test-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
    Start-Transcript -Path $LogPath -Append
    Write-Host "Detailed logging enabled. Log file: $LogPath" -ForegroundColor Green
}

# Define Microsoft 365 service endpoints for testing
$M365Endpoints = @{
    "Exchange Online" = @{
        Primary = "outlook.office365.com"
        Secondary = "outlook.office.com"
        Port = 443
        Protocol = "HTTPS"
        Description = "Exchange Online mail service"
    }
    "SharePoint Online" = @{
        Primary = "sharepoint.com"
        Secondary = "*.sharepoint.com"
        Port = 443
        Protocol = "HTTPS"
        Description = "SharePoint Online collaboration"
    }
    "Microsoft Teams" = @{
        Primary = "teams.microsoft.com"
        Secondary = "teams.live.com"
        Port = 443
        Protocol = "HTTPS"
        Description = "Microsoft Teams communication"
    }
    "OneDrive for Business" = @{
        Primary = "onedrive.live.com"
        Secondary = "*.files.1drv.com"
        Port = 443
        Protocol = "HTTPS"
        Description = "OneDrive file storage"
    }
    "Azure Active Directory" = @{
        Primary = "login.microsoftonline.com"
        Secondary = "login.windows.net"
        Port = 443
        Protocol = "HTTPS"
        Description = "Azure AD authentication"
    }
    "Office Online" = @{
        Primary = "office.com"
        Secondary = "www.office.com"
        Port = 443
        Protocol = "HTTPS"
        Description = "Office Online applications"
    }
    "Microsoft Graph" = @{
        Primary = "graph.microsoft.com"
        Secondary = "graph.windows.net"
        Port = 443
        Protocol = "HTTPS"
        Description = "Microsoft Graph API"
    }
    "Power Platform" = @{
        Primary = "make.powerapps.com"
        Secondary = "powerapps.microsoft.com"
        Port = 443
        Protocol = "HTTPS"
        Description = "Power Platform services"
    }
    "Microsoft Viva" = @{
        Primary = "myanalytics.microsoft.com"
        Secondary = "insights.viva.office.com"
        Port = 443
        Protocol = "HTTPS"
        Description = "Microsoft Viva employee experience"
    }
    "Microsoft Purview" = @{
        Primary = "compliance.microsoft.com"
        Secondary = "purview.microsoft.com"
        Port = 443
        Protocol = "HTTPS"
        Description = "Microsoft Purview compliance"
    }
}

# Bandwidth test endpoints - Using well-known, stable endpoints that are unlikely to change
$BandwidthTestEndpoints = @{
    "Microsoft.com Homepage" = "https://www.microsoft.com"
    "Office.com Portal" = "https://www.office.com"
    "Azure Portal" = "https://portal.azure.com"
}

# Results collection
$TestResults = @()
$BandwidthResults = @()
$Summary = @{
    TestStart = Get-Date
    TestEnd = $null
    TotalEndpoints = 0
    SuccessfulConnections = 0
    FailedConnections = 0
    AverageLatency = 0
    MaxLatency = 0
    MinLatency = [int]::MaxValue
}

# Helper Functions

function Write-TestProgress {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete = 0
    )
    
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
    if ($DetailedLogging) {
        Write-Verbose "[$((Get-Date).ToString('HH:mm:ss'))] $Activity - $Status"
    }
}

function Format-LocalDateTime {
    <#
    .SYNOPSIS
        Formats DateTime objects using local system format
    #>
    param(
        [DateTime]$DateTime,
        [switch]$IncludeTime = $true
    )
    
    if ($IncludeTime) {
        return $DateTime.ToString()  # Uses system's short date and time pattern
    }
    else {
        return $DateTime.ToString("d")  # Uses system's short date pattern
    }
}

function Get-PublicIPAddress {
    <#
    .SYNOPSIS
        Gets the public IP address of the workstation running the test
    #>
    param(
        [int]$TimeoutSeconds = 10
    )
    
    # List of reliable public IP services with fallbacks
    $ipServices = @(
        "https://api.ipify.org",
        "https://ipinfo.io/ip",
        "https://checkip.amazonaws.com",
        "https://icanhazip.com"
    )
    
    foreach ($service in $ipServices) {
        try {
            Write-Verbose "Trying to get public IP from: $service"
            $webClient = New-Object System.Net.WebClient
            $webClient.Headers.Add("User-Agent", "M365-Connection-Tester/1.0")
            
            # Set timeout
            $webRequest = [System.Net.WebRequest]::Create($service)
            $webRequest.Timeout = $TimeoutSeconds * 1000
            
            $response = $webRequest.GetResponse()
            $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
            $publicIP = $reader.ReadToEnd().Trim()
            $reader.Close()
            $response.Close()
            $webClient.Dispose()
            
            # Validate IP address format
            if ($publicIP -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
                Write-Verbose "Successfully retrieved public IP: $publicIP from $service"
                return @{
                    Success = $true
                    IPAddress = $publicIP
                    Source = $service
                    Error = $null
                }
            }
            else {
                Write-Verbose "Invalid IP format received from $service`: $publicIP"
                continue
            }
        }
        catch {
            Write-Verbose "Failed to get public IP from $service`: $($_.Exception.Message)"
            if ($webClient) { $webClient.Dispose() }
            continue
        }
    }
    
    # If all services failed, return error
    return @{
        Success = $false
        IPAddress = "Unknown"
        Source = $null
        Error = "Unable to determine public IP address from any service"
    }
}

function Test-DNSResolution {
    param(
        [string]$Hostname,
        [int]$Timeout = 5000
    )
    
    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $dnsResult = [System.Net.Dns]::GetHostAddresses($Hostname)
        $stopwatch.Stop()
        
        return @{
            Success = $true
            ResolutionTime = $stopwatch.ElapsedMilliseconds
            IPAddresses = $dnsResult | ForEach-Object { $_.ToString() }
            Error = $null
        }
    }
    catch {
        return @{
            Success = $false
            ResolutionTime = $null
            IPAddresses = @()
            Error = $_.Exception.Message
        }
    }
}

function Test-HTTPSConnection {
    param(
        [string]$Url,
        [int]$TimeoutSeconds = 30
    )
    
    # Try HEAD request first
    $headResult = Test-HTTPSRequest -Url $Url -Method "HEAD" -TimeoutSeconds $TimeoutSeconds
    
    # If HEAD returns certain status codes that indicate method not allowed, try GET
    if ($headResult.StatusCode -in @(405, 501)) {
        Write-Verbose "HEAD request returned $($headResult.StatusCode), trying GET request to root"
        $rootUrl = ([System.Uri]$Url).GetLeftPart([System.UriPartial]::Authority)
        $getResult = Test-HTTPSRequest -Url $rootUrl -Method "GET" -TimeoutSeconds $TimeoutSeconds -MaxBytes 1024
        
        # Return the better result (prefer successful GET over successful HEAD with method restrictions)
        if ($getResult.Success -and $getResult.StatusCode -lt 400) {
            return $getResult
        }
    }
    
    # If HEAD failed with 417 (Expectation Failed), try GET with the same URL
    if (-not $headResult.Success -and $headResult.StatusCode -eq 417) {
        Write-Verbose "HEAD request failed with 417 Expectation Failed, trying GET request"
        $getResult = Test-HTTPSRequest -Url $Url -Method "GET" -TimeoutSeconds $TimeoutSeconds -MaxBytes 1024
        
        if ($getResult.Success) {
            return $getResult
        }
    }
    
    return $headResult
}

function Test-HTTPSRequest {
    param(
        [string]$Url,
        [string]$Method = "HEAD",
        [int]$TimeoutSeconds = 30,
        [int]$MaxBytes = 0
    )
    
    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        
        # Create web request with proper configuration
        $webRequest = [System.Net.WebRequest]::Create($Url)
        $webRequest.Timeout = $TimeoutSeconds * 1000
        $webRequest.Method = $Method
        $webRequest.UserAgent = "M365-Connection-Tester/1.0"
        
        # Configure HTTP settings to avoid common issues
        if ($webRequest -is [System.Net.HttpWebRequest]) {
            $webRequest.ServicePoint.Expect100Continue = $false  # Avoid 417 Expectation Failed
            $webRequest.KeepAlive = $false  # Simplify connection handling
            $webRequest.AllowAutoRedirect = $true  # Follow redirects
            
            # For GET requests, limit the amount of data we read
            if ($Method -eq "GET" -and $MaxBytes -gt 0) {
                $webRequest.Headers.Add("Range", "bytes=0-$($MaxBytes-1)")
            }
        }
        
        # Configure TLS settings
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        
        $response = $webRequest.GetResponse()
        $stopwatch.Stop()
        
        $result = @{
            Success = $true
            ResponseTime = $stopwatch.ElapsedMilliseconds
            StatusCode = [int]$response.StatusCode
            StatusDescription = $response.StatusDescription
            ContentLength = $response.ContentLength
            Error = $null
        }
        
        $response.Close()
        return $result
    }
    catch [System.Net.WebException] {
        $stopwatch.Stop()
        $statusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { $null }
        
        # Treat certain HTTP errors as successful connectivity (server is responding)
        # These codes indicate the endpoint is reachable and the service is running
        $isConnectivitySuccess = $statusCode -in @(401, 403, 404, 405, 417, 429)  # Unauthorized, Forbidden, Not Found, Method Not Allowed, Expectation Failed, Too Many Requests
        
        return @{
            Success = $isConnectivitySuccess
            ResponseTime = $stopwatch.ElapsedMilliseconds
            StatusCode = $statusCode
            StatusDescription = if ($isConnectivitySuccess) { "Server responding (HTTP $statusCode)" } else { $_.Exception.Message }
            ContentLength = $null
            Error = if (-not $isConnectivitySuccess) { $_.Exception.Message } else { $null }
        }
    }
    catch {
        $stopwatch.Stop()
        return @{
            Success = $false
            ResponseTime = $stopwatch.ElapsedMilliseconds
            StatusCode = $null
            StatusDescription = $null
            ContentLength = $null
            Error = $_.Exception.Message
        }
    }
}

function Test-PortConnectivity {
    param(
        [string]$Hostname,
        [int]$Port,
        [int]$TimeoutSeconds = 10
    )
    
    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $tcpClient.BeginConnect($Hostname, $Port, $null, $null)
        $waitHandle = $asyncResult.AsyncWaitHandle
        
        if ($waitHandle.WaitOne($TimeoutSeconds * 1000)) {
            $tcpClient.EndConnect($asyncResult)
            $stopwatch.Stop()
            $tcpClient.Close()
            
            return @{
                Success = $true
                ConnectionTime = $stopwatch.ElapsedMilliseconds
                Error = $null
            }
        }
        else {
            $stopwatch.Stop()
            $tcpClient.Close()
            
            return @{
                Success = $false
                ConnectionTime = $stopwatch.ElapsedMilliseconds
                Error = "Connection timeout after $TimeoutSeconds seconds"
            }
        }
    }
    catch {
        $stopwatch.Stop()
        if ($tcpClient) { $tcpClient.Close() }
        
        return @{
            Success = $false
            ConnectionTime = $stopwatch.ElapsedMilliseconds
            Error = $_.Exception.Message
        }
    }
}

# Azure Blob Upload Functions

function Upload-ToAzureBlob {
    <#
    .SYNOPSIS
        Uploads file to Azure Blob Storage (unlimited uploads)
    #>
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,
        
        [Parameter(Mandatory)]
        [string]$SasUrl
    )
    
    try {
        if (-not (Test-Path $FilePath)) {
            return @{
                Success = $false
                Error = "File not found: $FilePath"
            }
        }
        
        Write-Host "🔄 Uploading test results to CIAOPS Azure Storage..." -ForegroundColor Yellow
        
        # Generate unique blob name
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $randomId = [Guid]::NewGuid().ToString().Substring(0,8)
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        $extension = [System.IO.Path]::GetExtension($FilePath)
        $uniqueBlobName = "$fileName-$timestamp-$randomId$extension"
        
        # Construct upload URL for the specific blob
        $containerUrl = $SasUrl.Split('?')[0]
        $sasParams = $SasUrl.Split('?')[1]
        $uploadUrl = "$containerUrl/$uniqueBlobName" + "?" + $sasParams
        
        # Read file content
        $fileContent = Get-Content -Path $FilePath -Raw
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($fileContent)
        $fileSize = $bytes.Length
        
        # Create web request
        $webRequest = [System.Net.WebRequest]::Create($uploadUrl)
        $webRequest.Method = "PUT"
        $webRequest.ContentType = "text/html"
        $webRequest.ContentLength = $fileSize
        $webRequest.Headers.Add("x-ms-blob-type", "BlockBlob")
        $webRequest.Headers.Add("x-ms-meta-uploaded", (Get-Date).ToString("o"))
        $webRequest.Headers.Add("x-ms-meta-source", "m365-connection-test")
        
        # Set timeouts
        $webRequest.Timeout = 60000  # 60 seconds
        $webRequest.ReadWriteTimeout = 300000  # 5 minutes
        
        # Upload file
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        
        $requestStream = $webRequest.GetRequestStream()
        $requestStream.Write($bytes, 0, $bytes.Length)
        $requestStream.Close()
        
        # Get response
        $response = $webRequest.GetResponse()
        $statusCode = [int]$response.StatusCode
        $response.Close()
        
        $stopwatch.Stop()
        
        if ($statusCode -eq 201) {
            Write-Host "✅ Upload successful!" -ForegroundColor Green
            Write-Host "Upload Time: $($stopwatch.ElapsedMilliseconds)ms" -ForegroundColor Gray
            Write-Host "File Size: $([math]::Round($fileSize / 1KB, 2)) KB" -ForegroundColor Gray
            
            return @{
                Success = $true
                BlobName = $uniqueBlobName
                UploadTime = $stopwatch.ElapsedMilliseconds
                FileSize = $fileSize
                HttpStatus = $statusCode
            }
        } else {
            return @{
                Success = $false
                Error = "HTTP Status: $statusCode"
                HttpStatus = $statusCode
            }
        }
    }
    catch [System.Net.WebException] {
        $statusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { "Unknown" }
        return @{
            Success = $false
            Error = "HTTP Error ${statusCode}: $($_.Exception.Message)"
            HttpStatus = $statusCode
        }
    }
    catch {
        return @{
            Success = $false
            Error = $_.Exception.Message
            HttpStatus = $null
        }
    }
}

function Test-BandwidthSpeed {
    param(
        [string]$TestUrl,
        [int]$DurationSeconds = 10
    )
    
    try {
        Write-TestProgress -Activity "Bandwidth Testing" -Status "Testing download speed from $TestUrl"
        
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $totalBytesRead = 0
        $testCount = 0
        $maxTests = 15  # Limit number of requests to avoid overwhelming servers
        
        # For homepage-style endpoints, we'll do multiple requests to measure throughput
        while ($stopwatch.ElapsedMilliseconds -lt ($DurationSeconds * 1000) -and $testCount -lt $maxTests) {
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.Headers.Add("User-Agent", "M365-Connection-Tester/1.0")
                
                # Use the URL directly without cache-busting parameters for homepages
                $requestStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                $data = $webClient.DownloadData($TestUrl)
                $requestStopwatch.Stop()
                
                $totalBytesRead += $data.Length
                $testCount++
                
                # Update progress
                if ($testCount % 2 -eq 0) {
                    $currentSpeedMbps = [math]::Round((($totalBytesRead * 8) / ($stopwatch.ElapsedMilliseconds / 1000)) / 1000000, 2)
                    Write-TestProgress -Activity "Bandwidth Testing" -Status "Completed $testCount requests - Downloaded $([math]::Round($totalBytesRead / 1KB, 2)) KB - Current speed: $currentSpeedMbps Mbps"
                }
                
                $webClient.Dispose()
                
                # Longer delay between requests to be respectful to servers
                Start-Sleep -Milliseconds 200
            }
            catch {
                Write-Verbose "Request failed: $($_.Exception.Message)"
                # Continue with other requests even if one fails
                if ($webClient) { $webClient.Dispose() }
                
                # If we're getting failures, add a longer delay before retrying
                Start-Sleep -Milliseconds 500
            }
        }
        
        $stopwatch.Stop()
        
        if ($totalBytesRead -gt 0 -and $testCount -gt 0) {
            $speedBps = $totalBytesRead / ($stopwatch.ElapsedMilliseconds / 1000)
            $speedMbps = ($speedBps * 8) / 1000000  # Convert to Mbps
            
            return @{
                Success = $true
                DownloadedBytes = $totalBytesRead
                DurationMs = $stopwatch.ElapsedMilliseconds
                SpeedBps = $speedBps
                SpeedMbps = [math]::Round($speedMbps, 2)
                RequestCount = $testCount
                Error = $null
            }
        }
        else {
            return @{
                Success = $false
                DownloadedBytes = 0
                DurationMs = $stopwatch.ElapsedMilliseconds
                SpeedBps = 0
                SpeedMbps = 0
                RequestCount = $testCount
                Error = "No successful requests completed within $DurationSeconds seconds"
            }
        }
    }
    catch {
        return @{
            Success = $false
            DownloadedBytes = 0
            DurationMs = 0
            SpeedBps = 0
            SpeedMbps = 0
            RequestCount = 0
            Error = $_.Exception.Message
        }
    }
}

function New-HTMLReport {
    <#
    .SYNOPSIS
        Generates an HTML report of the M365 connection test results
    #>
    param(
        [array]$TestResults,
        [array]$BandwidthResults,
        [hashtable]$Summary,
        [hashtable]$AuthenticationStatus,
        [string]$OutputPath,
        [bool]$IncludeBandwidthTest,
        [bool]$IncludeAuthentication
    )
    
    # Calculate connection quality
    $connectionQuality = "N/A"
    $qualityColor = "gray"
    $qualityDescription = ""
    
    if ($Summary.SuccessfulConnections -gt 0) {
        $connectionQuality = switch ($Summary.AverageLatency) {
            { $_ -lt 100 } { "Excellent"; break }
            { $_ -lt 200 } { "Good"; break }
            { $_ -lt 300 } { "Fair"; break }
            { $_ -lt 500 } { "Poor"; break }
            default { "Very Poor" }
        }
        
        $qualityColor = switch ($connectionQuality) {
            "Excellent" { "#28a745"; break }
            "Good" { "#28a745"; break }
            "Fair" { "#ffc107"; break }
            "Poor" { "#dc3545"; break }
            default { "#dc3545" }
        }
        
        $qualityDescription = switch ($connectionQuality) {
            "Excellent" { "Outstanding performance with minimal delays. Ideal for real-time collaboration and video conferencing." }
            "Good" { "Very responsive with slight delays. Suitable for most M365 activities including Teams calls." }
            "Fair" { "Noticeable delays but still functional. May experience minor delays in real-time applications." }
            "Poor" { "Significant delays affecting user experience. Video calls may be choppy, file uploads/downloads slower." }
            "Very Poor" { "Severe delays impacting productivity. Consider network optimization or contact IT support." }
        }
    }
    
    # Generate HTML content
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Connection Test Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
            color: #333;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            padding: 30px;
        }
        .header {
            text-align: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 3px solid #0078d4;
        }
        .header h1 {
            color: #0078d4;
            margin: 0;
            font-size: 2.5em;
        }
        .header .subtitle {
            color: #666;
            margin: 10px 0;
            font-size: 1.1em;
        }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        .summary-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
        }
        .summary-card h3 {
            margin: 0 0 10px 0;
            font-size: 1.2em;
        }
        .summary-card .value {
            font-size: 2em;
            font-weight: bold;
            margin: 10px 0;
        }
        .quality-card {
            background: linear-gradient(135deg, $qualityColor 0%, $qualityColor 100%);
        }
        .section {
            margin-bottom: 30px;
        }
        .section h2 {
            color: #0078d4;
            border-bottom: 2px solid #0078d4;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }
        .table-container {
            overflow-x: auto;
            margin-bottom: 20px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background-color: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        th, td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        th {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.9em;
            letter-spacing: 0.5px;
        }
        tr:hover {
            background-color: #f8f9fa;
        }
        .status-success {
            color: #28a745;
            font-weight: bold;
        }
        .status-error {
            color: #dc3545;
            font-weight: bold;
        }
        .status-warning {
            color: #ffc107;
            font-weight: bold;
        }
        /* Latency cell styling with strong background colors and readable text */
        .latency-excellent { 
            background-color: #28a745; 
            color: white; 
            font-weight: bold; 
            text-align: center;
            border-radius: 4px;
            padding: 8px 12px;
        }
        .latency-good { 
            background-color: #17a2b8; 
            color: white; 
            font-weight: bold; 
            text-align: center;
            border-radius: 4px;
            padding: 8px 12px;
        }
        .latency-fair { 
            background-color: #ffc107; 
            color: #212529; 
            font-weight: bold; 
            text-align: center;
            border-radius: 4px;
            padding: 8px 12px;
        }
        .latency-poor { 
            background-color: #fd7e14; 
            color: white; 
            font-weight: bold; 
            text-align: center;
            border-radius: 4px;
            padding: 8px 12px;
        }
        .latency-very-poor { 
            background-color: #dc3545; 
            color: white; 
            font-weight: bold; 
            text-align: center;
            border-radius: 4px;
            padding: 8px 12px;
        }
        .recommendations {
            background-color: #e7f3ff;
            border-left: 4px solid #0078d4;
            padding: 20px;
            border-radius: 0 8px 8px 0;
        }
        .recommendations h3 {
            color: #0078d4;
            margin-top: 0;
        }
        .recommendations ul {
            margin: 10px 0;
            padding-left: 20px;
        }
        .recommendations li {
            margin: 8px 0;
        }
        .footer {
            text-align: center;
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            color: #666;
            font-size: 0.9em;
        }
        .bandwidth-chart {
            margin: 20px 0;
        }
        .speed-bar {
            height: 30px;
            background: linear-gradient(90deg, #ff6b6b 0%, #feca57 50%, #48dbfb 100%);
            border-radius: 15px;
            position: relative;
            margin: 10px 0;
            overflow: hidden;
        }
        .speed-indicator {
            position: absolute;
            top: 0;
            left: 0;
            height: 100%;
            background-color: rgba(255,255,255,0.3);
            border-radius: 15px;
            transition: width 0.3s ease;
        }
        .speed-text {
            position: absolute;
            top: 50%;
            left: 10px;
            transform: translateY(-50%);
            color: white;
            font-weight: bold;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.5);
        }
        @media (max-width: 768px) {
            .container {
                padding: 15px;
            }
            .summary-grid {
                grid-template-columns: 1fr;
            }
            .header h1 {
                font-size: 2em;
            }
            table {
                font-size: 0.9em;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔗 Microsoft 365 Connection Test Report</h1>
            <div class="subtitle">Generated on $(Format-LocalDateTime -DateTime $Summary.TestStart)</div>
            <div class="subtitle">Test Duration: $([math]::Round((New-TimeSpan -Start $Summary.TestStart -End $Summary.TestEnd).TotalSeconds, 1)) seconds</div>
        </div>

        <div class="summary-grid">
            <div class="summary-card">
                <h3>📊 Total Endpoints</h3>
                <div class="value">$($Summary.TotalEndpoints)</div>
            </div>
            <div class="summary-card">
                <h3>✅ Successful</h3>
                <div class="value">$($Summary.SuccessfulConnections)</div>
            </div>
            <div class="summary-card">
                <h3>❌ Failed</h3>
                <div class="value">$($Summary.FailedConnections)</div>
            </div>
            <div class="summary-card">
                <h3>🌐 Workstation IP</h3>
                <div class="value" style="font-size: 1.2em;">$($Summary.WorkstationPublicIP)</div>
                <div style="font-size: 0.8em; margin-top: 5px; opacity: 0.9;">Public IP Address</div>
            </div>
            <div class="summary-card quality-card">
                <h3>🎯 Connection Quality</h3>
                <div class="value">$connectionQuality</div>
                <div style="font-size: 0.9em; margin-top: 5px;">$($Summary.AverageLatency)ms avg</div>
            </div>
        </div>

        <div class="section">
            <h2>📋 Connection Test Results</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Service</th>
                            <th>Endpoint</th>
                            <th>Status</th>
                            <th>Total Latency</th>
                            <th>DNS</th>
                            <th>Port</th>
                            <th>HTTPS</th>
"@

    if ($IncludeAuthentication) {
        $htmlContent += @"
                            <th>Auth</th>
"@
    }

    $htmlContent += @"
                        </tr>
                    </thead>
                    <tbody>
"@

    foreach ($result in $TestResults) {
        $statusClass = if ($result.OverallSuccess) { "status-success" } else { "status-error" }
        $statusIcon = if ($result.OverallSuccess) { "✅" } else { "❌" }
        $statusText = if ($result.OverallSuccess) { "Success" } else { "Failed" }
        
        $latencyClass = switch ($result.TotalLatency) {
            { $_ -lt 100 } { "latency-excellent"; break }
            { $_ -lt 200 } { "latency-good"; break }
            { $_ -lt 300 } { "latency-fair"; break }
            { $_ -lt 500 } { "latency-poor"; break }
            default { "latency-very-poor" }
        }
        
        $dnsStatus = if ($result.DNSResolution.Success) { "✅ $($result.DNSResolution.ResolutionTime)ms" } else { "❌ Failed" }
        $portStatus = if ($result.PortConnectivity.Success) { "✅ $($result.PortConnectivity.ConnectionTime)ms" } else { "❌ Failed" }
        $httpsStatus = if ($result.HTTPSResponse -and $result.HTTPSResponse.Success) { "✅ $($result.HTTPSResponse.ResponseTime)ms" } else { "❌ Failed" }
        
        $htmlContent += @"
                        <tr>
                            <td><strong>$($result.Service)</strong><br><small style="color: #666;">$($result.Description)</small></td>
                            <td><code>$($result.Endpoint)</code></td>
                            <td class="$statusClass">$statusIcon $statusText</td>
                            <td class="$latencyClass">$($result.TotalLatency)ms</td>
                            <td>$dnsStatus</td>
                            <td>$portStatus</td>
                            <td>$httpsStatus</td>
"@

        if ($IncludeAuthentication) {
            $authStatus = if ($result.AuthenticationTest -and $result.AuthenticationTest.AuthenticationTest) { 
                "✅ $($result.AuthenticationTest.AuthenticationTime)ms" 
            } elseif ($result.AuthenticationTest -and $result.AuthenticationTest.AuthenticationError) { 
                "❌ Error" 
            } else { 
                "➖ N/A" 
            }
            $htmlContent += @"
                            <td>$authStatus</td>
"@
        }

        $htmlContent += @"
                        </tr>
"@
    }

    $htmlContent += @"
                    </tbody>
                </table>
            </div>
            
            <!-- Connection Quality Legend -->
            <div style="margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 8px; border-left: 4px solid #0078d4;">
                <h4 style="margin: 0 0 15px 0; color: #0078d4;">🎯 Connection Quality Guide</h4>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 10px;">
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 12px; background-color: #28a745; border-radius: 6px; margin-right: 8px;"></div>
                        <span style="font-size: 0.9em;"><strong>Excellent (&lt; 100ms):</strong> Outstanding performance</span>
                    </div>
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 12px; background-color: #17a2b8; border-radius: 6px; margin-right: 8px;"></div>
                        <span style="font-size: 0.9em;"><strong>Good (100-200ms):</strong> Very responsive</span>
                    </div>
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 12px; background-color: #ffc107; border-radius: 6px; margin-right: 8px;"></div>
                        <span style="font-size: 0.9em;"><strong>Fair (200-300ms):</strong> Noticeable delays</span>
                    </div>
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 12px; background-color: #fd7e14; border-radius: 6px; margin-right: 8px;"></div>
                        <span style="font-size: 0.9em;"><strong>Poor (300-500ms):</strong> Significant delays</span>
                    </div>
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 12px; background-color: #dc3545; border-radius: 6px; margin-right: 8px;"></div>
                        <span style="font-size: 0.9em;"><strong>Very Poor (&gt; 500ms):</strong> Severe delays</span>
                    </div>
                </div>
                <p style="margin: 10px 0 0 0; font-size: 0.85em; color: #666; font-style: italic;">
                    💡 Latency values in the table above are color-coded according to these performance ranges
                </p>
            </div>
            
            <!-- Test Components Explanation -->
            <div style="margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 8px; border-left: 4px solid #0078d4;">
                <h4 style="margin: 0 0 15px 0; color: #0078d4;">🔍 Test Components Explained</h4>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 15px;">
                    <div style="padding: 10px; background-color: white; border-radius: 6px; border-left: 3px solid #28a745;">
                        <h5 style="margin: 0 0 8px 0; color: #28a745;">🌐 DNS Resolution</h5>
                        <p style="margin: 0; font-size: 0.9em; color: #666;">Tests how quickly your system can resolve the Microsoft 365 service domain names to IP addresses. Fast DNS resolution (under 20ms) is crucial for quick service connections.</p>
                    </div>
                    <div style="padding: 10px; background-color: white; border-radius: 6px; border-left: 3px solid #17a2b8;">
                        <h5 style="margin: 0 0 8px 0; color: #17a2b8;">🔌 Port Connectivity</h5>
                        <p style="margin: 0; font-size: 0.9em; color: #666;">Verifies that your network can establish a TCP connection to port 443 (HTTPS) on the M365 servers. This tests firewall and proxy configurations that might block access.</p>
                    </div>
                    <div style="padding: 10px; background-color: white; border-radius: 6px; border-left: 3px solid #ffc107;">
                        <h5 style="margin: 0 0 8px 0; color: #e68900;">🔒 HTTPS Response</h5>
                        <p style="margin: 0; font-size: 0.9em; color: #666;">Measures the time to receive an HTTP response from the M365 service. This includes SSL/TLS handshake time and indicates the overall responsiveness of the service endpoint.</p>
                    </div>
                </div>
                <div style="margin-top: 10px; padding: 10px; background-color: #e7f3ff; border-radius: 6px;">
                    <p style="margin: 0; font-size: 0.85em; color: #666; font-style: italic;">
                        <strong>💡 Pro Tip:</strong> The <strong>Total Latency</strong> column shows the combined time for all three tests, giving you the complete picture of how long it takes to establish a connection to each M365 service.
                    </p>
                </div>
            </div>
        </div>
"@

    # Add bandwidth results section if available
    if ($IncludeBandwidthTest -and $BandwidthResults.Count -gt 0) {
        $htmlContent += @"
        <div class="section">
            <h2>🚀 Bandwidth Test Results</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Test Endpoint</th>
                            <th>Status</th>
                            <th>Download Speed</th>
                            <th>Data Downloaded</th>
                            <th>Duration</th>
                            <th>Requests</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($bandwidth in $BandwidthResults) {
            $statusClass = if ($bandwidth.Success) { "status-success" } else { "status-error" }
            $statusIcon = if ($bandwidth.Success) { "✅" } else { "❌" }
            $statusText = if ($bandwidth.Success) { "Success" } else { "Failed" }
            $downloadedKB = [math]::Round($bandwidth.DownloadedBytes / 1KB, 2)
            
            # Format duration more appropriately
            $durationSec = $bandwidth.DurationMs / 1000
            $durationDisplay = if ($durationSec -lt 1) {
                "$([math]::Round($bandwidth.DurationMs, 0))ms"
            } elseif ($durationSec -lt 10) {
                "$([math]::Round($durationSec, 1))s"
            } else {
                "$([math]::Round($durationSec, 0))s"
            }
            
            # Calculate speed bar width and color based on speed ranges
            $speedPercentage = [math]::Min(($bandwidth.SpeedMbps / 20) * 100, 100)
            
            # Ensure minimum visible width for very slow speeds (at least 5% to show color)
            if ($speedPercentage -lt 5 -and $bandwidth.SpeedMbps -gt 0) {
                $speedPercentage = 5
            }
            
            # Determine color based on speed ranges
            $speedColor = switch ($bandwidth.SpeedMbps) {
                { $_ -ge 10 } { "#28a745"; break }  # Green for excellent (10+ Mbps)
                { $_ -ge 5 } { "#17a2b8"; break }   # Blue for very good (5-10 Mbps)
                { $_ -ge 2 } { "#ffc107"; break }   # Yellow for good (2-5 Mbps)
                { $_ -ge 1 } { "#fd7e14"; break }   # Orange for fair (1-2 Mbps)
                default { "#dc3545" }               # Red for poor (< 1 Mbps)
            }
            
            $htmlContent += @"
                        <tr>
                            <td><strong>$($bandwidth.TestName)</strong><br><small style="color: #666;"><code>$($bandwidth.TestUrl)</code></small></td>
                            <td class="$statusClass">$statusIcon $statusText</td>
                            <td>
                                <div style="display: flex; align-items: center;">
                                    <strong style="margin-right: 10px;">$($bandwidth.SpeedMbps) Mbps</strong>
                                    <div style="width: 100px; height: 20px; background-color: #e9ecef; border-radius: 10px; overflow: hidden;">
                                        <div style="width: $speedPercentage%; height: 100%; background-color: $speedColor; border-radius: 10px; transition: width 0.3s ease;"></div>
                                    </div>
                                </div>
                            </td>
                            <td>$downloadedKB KB</td>
                            <td>$durationDisplay</td>
                            <td>$($bandwidth.RequestCount)</td>
                        </tr>
"@
        }

        $htmlContent += @"
                    </tbody>
                </table>
            </div>
            
            <!-- Bandwidth Speed Legend -->
            <div style="margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 8px; border-left: 4px solid #0078d4;">
                <h4 style="margin: 0 0 15px 0; color: #0078d4;">📊 Bandwidth Speed Legend</h4>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 12px; background-color: #28a745; border-radius: 6px; margin-right: 8px;"></div>
                        <span style="font-size: 0.9em;"><strong>Excellent (10+ Mbps):</strong> Outstanding performance</span>
                    </div>
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 12px; background-color: #004085; border-radius: 6px; margin-right: 8px;"></div>
                        <span style="font-size: 0.9em;"><strong>Very Good (5-10 Mbps):</strong> Very responsive</span>
                    </div>
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 12px; background-color: #17a2b8; border-radius: 6px; margin-right: 8px;"></div>
                        <span style="font-size: 0.9em;"><strong>Good (2-5 Mbps):</strong> Very responsive</span>
                    </div>
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 12px; background-color: #ffc107; border-radius: 6px; margin-right: 8px;"></div>
                        <span style="font-size: 0.9em;"><strong>Fair (1-2 Mbps):</strong> Noticeable delays</span>
                    </div>
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 12px; background-color: #dc3545; border-radius: 6px; margin-right: 8px;"></div>
                        <span style="font-size: 0.9em;"><strong>Poor (&lt; 1 Mbps):</strong> Significant delays</span>
                    </div>
                </div>
                <p style="margin: 10px 0 0 0; font-size: 0.85em; color: #666; font-style: italic;">
                    💡 Bandwidth indicators show download speed performance for M365 connectivity testing
                </p>
            </div>
        </div>
"@
    }

    # Add authentication status if available
    if ($IncludeAuthentication -and $AuthenticationStatus) {
        $htmlContent += @"
        <div class="section">
            <h2>🔐 Authentication Status</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Service</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($auth in $AuthenticationStatus.GetEnumerator()) {
            $statusClass = if ($auth.Value) { "status-success" } else { "status-warning" }
            $statusIcon = if ($auth.Value) { "✅" } else { "⚠️" }
            $statusText = if ($auth.Value) { "Connected" } else { "Not Connected" }
            
            $htmlContent += @"
                        <tr>
                            <td>$($auth.Key)</td>
                            <td class="$statusClass">$statusIcon $statusText</td>
                        </tr>
"@
        }

        $htmlContent += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

    # Add recommendations section
    $htmlContent += @"
        <div class="section">
            <div class="recommendations">
                <h3>💡 Recommendations</h3>
"@

    if ($Summary.FailedConnections -gt 0) {
        $htmlContent += @"
                <p><strong>Connection Issues Detected:</strong></p>
                <ul>
                    <li>Some connections failed. Check firewall settings and proxy configuration.</li>
                    <li>Verify DNS settings and ensure M365 URLs are not blocked.</li>
                </ul>
"@
    }

    if ($Summary.AverageLatency -gt 200) {
        $htmlContent += @"
                <p><strong>High Latency Detected:</strong></p>
                <ul>
                    <li>Consider optimizing network routing.</li>
                    <li>Check for network congestion or bandwidth limitations.</li>
                </ul>
"@
    }

    $htmlContent += @"
                <p><strong>General Recommendations:</strong></p>
                <ul>
                    <li>For detailed M365 network guidance, visit: <a href="https://aka.ms/m365networkconnectivity" target="_blank">Microsoft 365 Network Connectivity</a></li>
                    <li>Consider using Microsoft 365 connectivity test tool: <a href="https://connectivity.office.com" target="_blank">connectivity.office.com</a></li>
"@

    if (-not $IncludeAuthentication) {
        $htmlContent += @"
                    <li>Run with -IncludeAuthentication for enhanced testing including authenticated API calls.</li>
"@
    }

    $htmlContent += @"
                </ul>
            </div>
        </div>

        <div class="footer">
            <p>📊 Generated by CIAOPS M365 Connection Test Tool | $(Format-LocalDateTime -DateTime (Get-Date))</p>
            <p>This report provides comprehensive network connectivity metrics for Microsoft 365 services.</p>
        </div>
    </div>

    <script>
        // Add some interactivity
        document.addEventListener('DOMContentLoaded', function() {
            // Animate speed bars
            const speedBars = document.querySelectorAll('.speed-indicator');
            speedBars.forEach(bar => {
                const width = bar.style.width;
                bar.style.width = '0%';
                setTimeout(() => {
                    bar.style.width = width;
                }, 100);
            });
            
            // Add click-to-copy functionality for endpoints
            const endpoints = document.querySelectorAll('code');
            endpoints.forEach(code => {
                code.style.cursor = 'pointer';
                code.title = 'Click to copy';
                code.addEventListener('click', function() {
                    navigator.clipboard.writeText(this.textContent).then(() => {
                        const originalText = this.textContent;
                        this.textContent = 'Copied!';
                        setTimeout(() => {
                            this.textContent = originalText;
                        }, 1000);
                    });
                });
            });
        });
    </script>
</body>
</html>
"@

    # Write the HTML content to file
    $htmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
}

# Authentication Functions

function Test-M365Authentication {
    <#
    .SYNOPSIS
        Tests authentication status for various M365 services
    #>
    param(
        [switch]$ShowDetails
    )
    
    $authStatus = @{
        AzureAD = $false
        ExchangeOnline = $false
        SharePointOnline = $false
        MicrosoftGraph = $false
        Teams = $false
        SecurityCompliance = $false
    }
    
    try {
        # Check Azure AD / Microsoft Graph PowerShell
        if (Get-Module -Name "Microsoft.Graph.Authentication" -ListAvailable -ErrorAction SilentlyContinue) {
            try {
                $context = Get-MgContext -ErrorAction SilentlyContinue
                if ($context -and $context.Account) {
                    $authStatus.MicrosoftGraph = $true
                    $authStatus.AzureAD = $true
                    if ($ShowDetails) {
                        Write-Host "  ✓ Microsoft Graph: Connected as $($context.Account)" -ForegroundColor Green
                    }
                }
            }
            catch {
                if ($ShowDetails) {
                    Write-Host "  ✗ Microsoft Graph: Not connected" -ForegroundColor Yellow
                }
            }
        }
        
        # Check Exchange Online PowerShell
        if (Get-Module -Name "ExchangeOnlineManagement" -ListAvailable -ErrorAction SilentlyContinue) {
            try {
                $exoSession = Get-PSSession | Where-Object { $_.ComputerName -like "*outlook.office365.com*" -and $_.State -eq "Opened" }
                if ($exoSession) {
                    $authStatus.ExchangeOnline = $true
                    if ($ShowDetails) {
                        Write-Host "  ✓ Exchange Online: Connected" -ForegroundColor Green
                    }
                }
                else {
                    if ($ShowDetails) {
                        Write-Host "  ✗ Exchange Online: Not connected" -ForegroundColor Yellow
                    }
                }
            }
            catch {
                if ($ShowDetails) {
                    Write-Host "  ✗ Exchange Online: Not connected" -ForegroundColor Yellow
                }
            }
        }
        
        # Check SharePoint Online PowerShell
        if (Get-Module -Name "Microsoft.Online.SharePoint.PowerShell" -ListAvailable -ErrorAction SilentlyContinue) {
            try {
                $spoConnection = Get-SPOSite -Limit 1 -ErrorAction SilentlyContinue
                if ($spoConnection) {
                    $authStatus.SharePointOnline = $true
                    if ($ShowDetails) {
                        Write-Host "  ✓ SharePoint Online: Connected" -ForegroundColor Green
                    }
                }
            }
            catch {
                if ($ShowDetails) {
                    Write-Host "  ✗ SharePoint Online: Not connected" -ForegroundColor Yellow
                }
            }
        }
        
        # Check Teams PowerShell
        if (Get-Module -Name "MicrosoftTeams" -ListAvailable -ErrorAction SilentlyContinue) {
            try {
                $teamsContext = Get-CsOnlineSession -ErrorAction SilentlyContinue
                if ($teamsContext) {
                    $authStatus.Teams = $true
                    if ($ShowDetails) {
                        Write-Host "  ✓ Microsoft Teams: Connected" -ForegroundColor Green
                    }
                }
            }
            catch {
                if ($ShowDetails) {
                    Write-Host "  ✗ Microsoft Teams: Not connected" -ForegroundColor Yellow
                }
            }
        }
        
        # Check Security & Compliance PowerShell
        if (Get-Module -Name "ExchangeOnlineManagement" -ListAvailable -ErrorAction SilentlyContinue) {
            try {
                $sccSession = Get-PSSession | Where-Object { $_.ComputerName -like "*ps.compliance.protection.outlook.com*" -and $_.State -eq "Opened" }
                if ($sccSession) {
                    $authStatus.SecurityCompliance = $true
                    if ($ShowDetails) {
                        Write-Host "  ✓ Security & Compliance: Connected" -ForegroundColor Green
                    }
                }
            }
            catch {
                if ($ShowDetails) {
                    Write-Host "  ✗ Security & Compliance: Not connected" -ForegroundColor Yellow
                }
            }
        }
        
    }
    catch {
        Write-Warning "Error checking authentication status: $($_.Exception.Message)"
    }
    
    return $authStatus
}

function Connect-M365Services {
    <#
    .SYNOPSIS
        Connects to Microsoft 365 services interactively
    #>
    param(
        [string[]]$Services = @("MicrosoftGraph", "ExchangeOnline", "SharePointOnline", "Teams"),
        [switch]$Force
    )
    
    Write-Host "`n=== Microsoft 365 Authentication ===" -ForegroundColor Cyan
    Write-Host "Connecting to M365 services for enhanced testing..." -ForegroundColor Yellow
    
    $connectionResults = @{}
    
    foreach ($service in $Services) {
        Write-Host "`nConnecting to $service..." -ForegroundColor White
        
        try {
            switch ($service) {
                "MicrosoftGraph" {
                    # Check if Microsoft Graph module is available
                    if (-not (Get-Module -Name "Microsoft.Graph.Authentication" -ListAvailable)) {
                        Write-Host "  ⚠ Microsoft Graph PowerShell module not found" -ForegroundColor Yellow
                        Write-Host "  Install with: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Gray
                        $connectionResults[$service] = $false
                        continue
                    }
                    
                    # Connect to Microsoft Graph
                    Connect-MgGraph -Scopes "User.Read", "Directory.Read.All", "Reports.Read.All" -NoWelcome -ErrorAction Stop
                    $context = Get-MgContext
                    Write-Host "  ✓ Connected to Microsoft Graph as $($context.Account)" -ForegroundColor Green
                    $connectionResults[$service] = $true
                }
                
                "ExchangeOnline" {
                    if (-not (Get-Module -Name "ExchangeOnlineManagement" -ListAvailable)) {
                        Write-Host "  ⚠ Exchange Online PowerShell module not found" -ForegroundColor Yellow
                        Write-Host "  Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Gray
                        $connectionResults[$service] = $false
                        continue
                    }
                    
                    Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop
                    Write-Host "  ✓ Connected to Exchange Online" -ForegroundColor Green
                    $connectionResults[$service] = $true
                }
                
                "SharePointOnline" {
                    if (-not (Get-Module -Name "Microsoft.Online.SharePoint.PowerShell" -ListAvailable)) {
                        Write-Host "  ⚠ SharePoint Online PowerShell module not found" -ForegroundColor Yellow
                        Write-Host "  Install with: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser" -ForegroundColor Gray
                        $connectionResults[$service] = $false
                        continue
                    }
                    
                    # Prompt for tenant name
                    $tenantName = Read-Host "Enter your SharePoint tenant name (e.g., contoso for contoso.sharepoint.com)"
                    if ($tenantName) {
                        Connect-SPOService -Url "https://$tenantName-admin.sharepoint.com" -ErrorAction Stop
                        Write-Host "  ✓ Connected to SharePoint Online" -ForegroundColor Green
                        $connectionResults[$service] = $true
                    }
                    else {
                        Write-Host "  ✗ Tenant name required for SharePoint Online" -ForegroundColor Red
                        $connectionResults[$service] = $false
                    }
                }
                
                "Teams" {
                    if (-not (Get-Module -Name "MicrosoftTeams" -ListAvailable)) {
                        Write-Host "  ⚠ Microsoft Teams PowerShell module not found" -ForegroundColor Yellow
                        Write-Host "  Install with: Install-Module MicrosoftTeams -Scope CurrentUser" -ForegroundColor Gray
                        $connectionResults[$service] = $false
                        continue
                    }
                    
                    Connect-MicrosoftTeams -ErrorAction Stop
                    Write-Host "  ✓ Connected to Microsoft Teams" -ForegroundColor Green
                    $connectionResults[$service] = $true
                }
                
                default {
                    Write-Host "  ✗ Unknown service: $service" -ForegroundColor Red
                    $connectionResults[$service] = $false
                }
            }
        }
        catch {
            Write-Host "  ✗ Failed to connect to $service`: $($_.Exception.Message)" -ForegroundColor Red
            $connectionResults[$service] = $false
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    # Summary
    $successfulConnections = ($connectionResults.Values | Where-Object { $_ -eq $true }).Count
    $totalAttempted = $connectionResults.Count
    
    Write-Host "`nAuthentication Summary:" -ForegroundColor Cyan
    Write-Host "Successfully connected to $successfulConnections of $totalAttempted services" -ForegroundColor White
    
    return $connectionResults
}

function Test-AuthenticatedEndpoint {
    <#
    .SYNOPSIS
        Tests authenticated endpoints using existing sessions
    #>
    param(
        [string]$Service,
        [hashtable]$AuthStatus
    )
    
    $result = @{
        Service = $Service
        AuthenticationTest = $false
        AuthenticationTime = 0
        AuthenticationError = $null
        UserInfo = $null
    }
    
    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        
        switch ($Service) {
            "Microsoft Graph" {
                if ($AuthStatus.MicrosoftGraph) {
                    $user = Get-MgUser -Top 1 -ErrorAction Stop
                    $result.AuthenticationTest = $true
                    $result.UserInfo = "Graph API accessible"
                }
            }
            
            "Exchange Online" {
                if ($AuthStatus.ExchangeOnline) {
                    $mailbox = Get-Mailbox -ResultSize 1 -ErrorAction Stop
                    $result.AuthenticationTest = $true
                    $result.UserInfo = "Exchange Online accessible"
                }
            }
            
            "SharePoint Online" {
                if ($AuthStatus.SharePointOnline) {
                    $site = Get-SPOSite -Limit 1 -ErrorAction Stop
                    $result.AuthenticationTest = $true
                    $result.UserInfo = "SharePoint Online accessible"
                }
            }
            
            "Microsoft Teams" {
                if ($AuthStatus.Teams) {
                    $tenant = Get-CsTenant -ErrorAction Stop
                    $result.AuthenticationTest = $true
                    $result.UserInfo = "Teams accessible"
                }
            }
        }
        
        $stopwatch.Stop()
        $result.AuthenticationTime = $stopwatch.ElapsedMilliseconds
    }
    catch {
        $stopwatch.Stop()
        $result.AuthenticationTime = $stopwatch.ElapsedMilliseconds
        $result.AuthenticationError = $_.Exception.Message
    }
    
    return $result
}

# Main Testing Logic

# Handle ShowUploadStats parameter independently
if ($ShowUploadStats -and -not $UploadResults) {
    Write-Host "`n=== Upload Statistics ===" -ForegroundColor Cyan
    Write-Host "✅ Unlimited uploads enabled - no restrictions!" -ForegroundColor Green
    
    Write-Host "`nTo upload results, use: .\test-m365-connection-speed.ps1 -UploadResults" -ForegroundColor Cyan
    Write-Host "Next run command: .\test-m365-connection-speed.ps1 -UploadResults" -ForegroundColor Gray
    return
}

Write-Host "`n=== Microsoft 365 Connection Speed Test ===" -ForegroundColor Cyan
Write-Host "Test started: $(Format-LocalDateTime -DateTime $Summary.TestStart)" -ForegroundColor Green

# Display upload option information
if ($UploadResults) {
    Write-Host "`n🔒 SECURE UPLOAD ENABLED" -ForegroundColor Green
    Write-Host "Your test results will be securely uploaded to CIAOPS Azure Blob Storage" -ForegroundColor White
    Write-Host "for analysis and to help improve M365 connectivity recommendations." -ForegroundColor White
    Write-Host "✅ Unlimited uploads enabled" -ForegroundColor Yellow
} elseif (-not $ShowUploadStats) {
    Write-Host "`n💡 TIP: Use -UploadResults to securely share your results with CIAOPS" -ForegroundColor Cyan
    Write-Host "   This helps improve M365 connectivity guidance and provides you with enhanced support." -ForegroundColor Gray
    Write-Host "   Your data is anonymized and stored securely. Unlimited uploads available." -ForegroundColor Gray
}

# Get workstation public IP address
Write-Host "`n=== Workstation Information ===" -ForegroundColor Cyan
Write-Host "Detecting public IP address..." -ForegroundColor Yellow

$PublicIPResult = Get-PublicIPAddress -TimeoutSeconds 10
if ($PublicIPResult.Success) {
    Write-Host "✅ Public IP Address: $($PublicIPResult.IPAddress) (Source: $($PublicIPResult.Source))" -ForegroundColor Green
    $Summary.Add("WorkstationPublicIP", $PublicIPResult.IPAddress)
    $Summary.Add("WorkstationIPSource", $PublicIPResult.Source)
} else {
    Write-Host "⚠️  Unable to determine public IP: $($PublicIPResult.Error)" -ForegroundColor Yellow
    $Summary.Add("WorkstationPublicIP", "Unknown")
    $Summary.Add("WorkstationIPSource", "N/A")
}

# Authentication section
$AuthenticationStatus = @{}
$AuthenticationResults = @{}

if ($IncludeAuthentication -and -not $SkipAuthenticationCheck) {
    Write-Host "`n=== Authentication Check ===" -ForegroundColor Cyan
    Write-Host "Checking existing authentication status..." -ForegroundColor Yellow
    
    $AuthenticationStatus = Test-M365Authentication -ShowDetails
    
    # Check if any services are not authenticated
    $unauthenticatedServices = @()
    if (-not $AuthenticationStatus.MicrosoftGraph) { $unauthenticatedServices += "MicrosoftGraph" }
    if (-not $AuthenticationStatus.ExchangeOnline) { $unauthenticatedServices += "ExchangeOnline" }
    if (-not $AuthenticationStatus.SharePointOnline) { $unauthenticatedServices += "SharePointOnline" }
    if (-not $AuthenticationStatus.Teams) { $unauthenticatedServices += "Teams" }
    
    if ($unauthenticatedServices.Count -gt 0) {
        Write-Host "`nSome services require authentication for enhanced testing." -ForegroundColor Yellow
        $response = Read-Host "Would you like to authenticate to M365 services? (Y/N)"
        
        if ($response -match "^[Yy]") {
            $connectionResults = Connect-M365Services -Services $unauthenticatedServices
            
            # Update authentication status
            foreach ($service in $connectionResults.Keys) {
                switch ($service) {
                    "MicrosoftGraph" { 
                        $AuthenticationStatus.MicrosoftGraph = $connectionResults[$service]
                        $AuthenticationStatus.AzureAD = $connectionResults[$service]
                    }
                    "ExchangeOnline" { $AuthenticationStatus.ExchangeOnline = $connectionResults[$service] }
                    "SharePointOnline" { $AuthenticationStatus.SharePointOnline = $connectionResults[$service] }
                    "Teams" { $AuthenticationStatus.Teams = $connectionResults[$service] }
                }
            }
        }
        else {
            Write-Host "Proceeding with network connectivity tests only..." -ForegroundColor Gray
        }
    }
    else {
        Write-Host "All available services are already authenticated!" -ForegroundColor Green
    }
}

Write-Host "`n=== Network Connectivity Testing ===" -ForegroundColor Cyan
Write-Host "Testing $($M365Endpoints.Count) M365 service endpoints..." -ForegroundColor Yellow

$currentEndpoint = 0
$Summary.TotalEndpoints = $M365Endpoints.Count

foreach ($service in $M365Endpoints.Keys) {
    $currentEndpoint++
    $endpoint = $M365Endpoints[$service]
    $percentComplete = [math]::Round(($currentEndpoint / $Summary.TotalEndpoints) * 100, 0)
    
    Write-TestProgress -Activity "Testing M365 Services" -Status "Testing $service ($currentEndpoint of $($Summary.TotalEndpoints))" -PercentComplete $percentComplete
    
    Write-Host "`nTesting: $service" -ForegroundColor White
    Write-Host "Endpoint: $($endpoint.Primary)" -ForegroundColor Gray
    
    # Initialize result object
    $result = @{
        Service = $service
        Endpoint = $endpoint.Primary
        Description = $endpoint.Description
        Timestamp = Get-Date
        DNSResolution = $null
        PortConnectivity = $null
        HTTPSResponse = $null
        AuthenticationTest = $null
        OverallSuccess = $false
        TotalLatency = 0
    }
    
    # Test DNS Resolution
    Write-Host "  → DNS Resolution..." -NoNewline
    $dnsTest = Test-DNSResolution -Hostname $endpoint.Primary
    $result.DNSResolution = $dnsTest
    
    if ($dnsTest.Success) {
        Write-Host " ✓ $($dnsTest.ResolutionTime)ms" -ForegroundColor Green
        $result.TotalLatency += $dnsTest.ResolutionTime
    }
    else {
        Write-Host " ✗ Failed: $($dnsTest.Error)" -ForegroundColor Red
    }
    
    # Test Port Connectivity
    Write-Host "  → Port $($endpoint.Port) connectivity..." -NoNewline
    $portTest = Test-PortConnectivity -Hostname $endpoint.Primary -Port $endpoint.Port
    $result.PortConnectivity = $portTest
    
    if ($portTest.Success) {
        Write-Host " ✓ $($portTest.ConnectionTime)ms" -ForegroundColor Green
        $result.TotalLatency += $portTest.ConnectionTime
    }
    else {
        Write-Host " ✗ Failed: $($portTest.Error)" -ForegroundColor Red
    }
    
    # Test HTTPS Response
    if ($endpoint.Protocol -eq "HTTPS" -and $dnsTest.Success -and $portTest.Success) {
        Write-Host "  → HTTPS response..." -NoNewline
        $httpsTest = Test-HTTPSConnection -Url "https://$($endpoint.Primary)"
        $result.HTTPSResponse = $httpsTest
        
        if ($httpsTest.Success) {
            Write-Host " ✓ $($httpsTest.ResponseTime)ms (Status: $($httpsTest.StatusCode))" -ForegroundColor Green
            $result.TotalLatency += $httpsTest.ResponseTime
            $result.OverallSuccess = $true
            $Summary.SuccessfulConnections++
        }
        else {
            Write-Host " ✗ Failed: $($httpsTest.Error)" -ForegroundColor Red
            $Summary.FailedConnections++
        }
    }
    else {
        $Summary.FailedConnections++
    }
    
    # Test Authentication (if enabled and authenticated)
    if ($IncludeAuthentication -and $result.OverallSuccess) {
        Write-Host "  → Authentication test..." -NoNewline
        $authTest = Test-AuthenticatedEndpoint -Service $service -AuthStatus $AuthenticationStatus
        $result.AuthenticationTest = $authTest
        
        if ($authTest.AuthenticationTest) {
            Write-Host " ✓ $($authTest.AuthenticationTime)ms ($($authTest.UserInfo))" -ForegroundColor Green
        }
        else {
            if ($authTest.AuthenticationError) {
                Write-Host " ✗ Auth failed: $($authTest.AuthenticationError)" -ForegroundColor Yellow
            }
            else {
                Write-Host " - Not authenticated" -ForegroundColor Gray
            }
        }
    }
    
    # Update latency statistics
    if ($result.OverallSuccess) {
        if ($result.TotalLatency -gt $Summary.MaxLatency) {
            $Summary.MaxLatency = $result.TotalLatency
        }
        if ($result.TotalLatency -lt $Summary.MinLatency) {
            $Summary.MinLatency = $result.TotalLatency
        }
    }
    
    $TestResults += $result
    
    # Brief pause between tests to avoid overwhelming endpoints
    Start-Sleep -Milliseconds 500
}

# Bandwidth Testing (if enabled)
if ($IncludeBandwidthTest) {
    Write-Host "`n=== Bandwidth Testing ===" -ForegroundColor Cyan
    
    foreach ($testName in $BandwidthTestEndpoints.Keys) {
        $testUrl = $BandwidthTestEndpoints[$testName]
        Write-Host "`nTesting bandwidth via: $testName" -ForegroundColor White
        
        $bandwidthResult = Test-BandwidthSpeed -TestUrl $testUrl -DurationSeconds $TestDuration
        
        # Store bandwidth result for output files
        $bandwidthTestResult = @{
            TestName = $testName
            TestUrl = $testUrl
            Success = $bandwidthResult.Success
            SpeedMbps = $bandwidthResult.SpeedMbps
            DownloadedBytes = $bandwidthResult.DownloadedBytes
            DurationMs = $bandwidthResult.DurationMs
            RequestCount = $bandwidthResult.RequestCount
            Error = $bandwidthResult.Error
            Timestamp = Get-Date
        }
        $BandwidthResults += $bandwidthTestResult
        
        if ($bandwidthResult.Success) {
            Write-Host "  → Download Speed: $($bandwidthResult.SpeedMbps) Mbps" -ForegroundColor Green
            Write-Host "  → Downloaded: $([math]::Round($bandwidthResult.DownloadedBytes / 1KB, 2)) KB in $([math]::Round($bandwidthResult.DurationMs / 1000, 1)) seconds ($($bandwidthResult.RequestCount) requests)" -ForegroundColor Green
        }
        else {
            Write-Host "  → Bandwidth test failed: $($bandwidthResult.Error)" -ForegroundColor Red
        }
    }
}

# Calculate summary statistics
$Summary.TestEnd = Get-Date
if ($Summary.SuccessfulConnections -gt 0) {
    $successfulLatencies = $TestResults | Where-Object { $_.OverallSuccess } | ForEach-Object { $_.TotalLatency }
    $Summary.AverageLatency = [math]::Round(($successfulLatencies | Measure-Object -Average).Average, 0)
}

# Display Summary
Write-Host "`n=== Test Summary ===" -ForegroundColor Cyan
Write-Host "Test Duration: $([math]::Round((New-TimeSpan -Start $Summary.TestStart -End $Summary.TestEnd).TotalSeconds, 1)) seconds" -ForegroundColor White
Write-Host "Test completed: $(Format-LocalDateTime -DateTime $Summary.TestEnd)" -ForegroundColor White
Write-Host "Total Endpoints Tested: $($Summary.TotalEndpoints)" -ForegroundColor White
Write-Host "Successful Connections: $($Summary.SuccessfulConnections)" -ForegroundColor Green
Write-Host "Failed Connections: $($Summary.FailedConnections)" -ForegroundColor Red

if ($Summary.SuccessfulConnections -gt 0) {
    Write-Host "Average Latency: $($Summary.AverageLatency)ms" -ForegroundColor White
    Write-Host "Min Latency: $($Summary.MinLatency)ms" -ForegroundColor White
    Write-Host "Max Latency: $($Summary.MaxLatency)ms" -ForegroundColor White
    
    # Connection quality assessment
    $connectionQuality = switch ($Summary.AverageLatency) {
        { $_ -lt 100 } { "Excellent"; break }
        { $_ -lt 200 } { "Good"; break }
        { $_ -lt 300 } { "Fair"; break }
        { $_ -lt 500 } { "Poor"; break }
        default { "Very Poor" }
    }
    
    $qualityColor = switch ($connectionQuality) {
        "Excellent" { "Green"; break }
        "Good" { "Green"; break }
        "Fair" { "Yellow"; break }
        "Poor" { "Red"; break }
        default { "Red" }
    }
    
    Write-Host "Connection Quality: $connectionQuality" -ForegroundColor $qualityColor
    
    # Explain the connection quality rating
    Write-Host "`nConnection Quality Ratings:" -ForegroundColor Cyan
    switch ($connectionQuality) {
        "Excellent" { 
            Write-Host "• Excellent (< 100ms): Outstanding performance with minimal delays" -ForegroundColor Green
            Write-Host "  - Ideal for real-time collaboration and video conferencing" -ForegroundColor Gray
        }
        "Good" { 
            Write-Host "• Good (100-199ms): Very responsive with slight delays" -ForegroundColor Green
            Write-Host "  - Suitable for most M365 activities including Teams calls" -ForegroundColor Gray
        }
        "Fair" { 
            Write-Host "• Fair (200-299ms): Noticeable delays but still functional" -ForegroundColor Yellow
            Write-Host "  - May experience minor delays in real-time applications" -ForegroundColor Gray
        }
        "Poor" { 
            Write-Host "• Poor (300-499ms): Significant delays affecting user experience" -ForegroundColor Red
            Write-Host "  - Video calls may be choppy, file uploads/downloads slower" -ForegroundColor Gray
        }
        "Very Poor" { 
            Write-Host "• Very Poor (≥ 500ms): Severe delays impacting productivity" -ForegroundColor Red
            Write-Host "  - Consider network optimization or contact IT support" -ForegroundColor Gray
        }
    }
}

# Output Results
switch ($OutputFormat) {
    "CSV" {
        if (-not $OutputPath) {
            # Use parent directory as default output location
            $parentPath = Split-Path (Get-Location) -Parent
            if (Test-Path $parentPath) {
                $OutputPath = Join-Path $parentPath "M365-Connection-Test.csv"
            }
            else {
                $OutputPath = Join-Path (Get-Location) "M365-Connection-Test.csv"
            }
        }
        
        # Ensure the output directory exists
        $outputDirectory = Split-Path $OutputPath -Parent
        if ([string]::IsNullOrEmpty($outputDirectory)) {
            # If no directory specified, use current directory
            $outputDirectory = Get-Location
        }
        if (-not (Test-Path $outputDirectory)) {
            New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
        }
        
        $csvData = $TestResults | ForEach-Object {
            [PSCustomObject]@{
                Service = $_.Service
                Endpoint = $_.Endpoint
                Description = $_.Description
                OverallSuccess = $_.OverallSuccess
                TotalLatency = $_.TotalLatency
                DNSSuccess = $_.DNSResolution.Success
                DNSTime = $_.DNSResolution.ResolutionTime
                PortSuccess = $_.PortConnectivity.Success
                PortTime = $_.PortConnectivity.ConnectionTime
                HTTPSSuccess = if ($_.HTTPSResponse) { $_.HTTPSResponse.Success } else { $null }
                HTTPSTime = if ($_.HTTPSResponse) { $_.HTTPSResponse.ResponseTime } else { $null }
                HTTPSStatus = if ($_.HTTPSResponse) { $_.HTTPSResponse.StatusCode } else { $null }
                AuthSuccess = if ($_.AuthenticationTest) { $_.AuthenticationTest.AuthenticationTest } else { $null }
                AuthTime = if ($_.AuthenticationTest) { $_.AuthenticationTest.AuthenticationTime } else { $null }
                AuthInfo = if ($_.AuthenticationTest) { $_.AuthenticationTest.UserInfo } else { $null }
                WorkstationPublicIP = $Summary.WorkstationPublicIP
            }
        }
        
        $csvData | Export-Csv -Path $OutputPath -NoTypeInformation
        Write-Host "`nResults exported to: $OutputPath" -ForegroundColor Green
        
        # Export bandwidth results to separate CSV file if bandwidth testing was performed
        if ($IncludeBandwidthTest -and $BandwidthResults.Count -gt 0) {
            $bandwidthOutputPath = $OutputPath -replace '\.csv$', '-Bandwidth.csv'
            
            $bandwidthCsvData = $BandwidthResults | ForEach-Object {
                [PSCustomObject]@{
                    TestName = $_.TestName
                    TestUrl = $_.TestUrl
                    Success = $_.Success
                    SpeedMbps = $_.SpeedMbps
                    DownloadedBytes = $_.DownloadedBytes
                    DurationMs = $_.DurationMs
                    RequestCount = $_.RequestCount
                    Error = $_.Error
                }
            }
            
            $bandwidthCsvData | Export-Csv -Path $bandwidthOutputPath -NoTypeInformation
            Write-Host "Bandwidth results exported to: $bandwidthOutputPath" -ForegroundColor Green
        }
    }
    
    "JSON" {
        if (-not $OutputPath) {
            # Use parent directory as default output location
            $parentPath = Split-Path (Get-Location) -Parent
            if (Test-Path $parentPath) {
                $OutputPath = Join-Path $parentPath "M365-Connection-Test.json"
            }
            else {
                $OutputPath = Join-Path (Get-Location) "M365-Connection-Test.json"
            }
        }
        
        # Ensure the output directory exists
        $outputDirectory = Split-Path $OutputPath -Parent
        if ([string]::IsNullOrEmpty($outputDirectory)) {
            # If no directory specified, use current directory
            $outputDirectory = Get-Location
        }
        if (-not (Test-Path $outputDirectory)) {
            New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
        }
        
        $jsonData = @{
            Summary = @{
                TestStart = Format-LocalDateTime -DateTime $Summary.TestStart
                TestEnd = if ($Summary.TestEnd) { Format-LocalDateTime -DateTime $Summary.TestEnd } else { $null }
                TotalEndpoints = $Summary.TotalEndpoints
                SuccessfulConnections = $Summary.SuccessfulConnections
                FailedConnections = $Summary.FailedConnections
                AverageLatency = $Summary.AverageLatency
                MaxLatency = $Summary.MaxLatency
                MinLatency = $Summary.MinLatency
            }
            TestResults = $TestResults | ForEach-Object {
                $result = $_ | ConvertTo-Json -Depth 5 | ConvertFrom-Json
                # Remove timestamp from result
                $result.PSObject.Properties.Remove('Timestamp')
                $result
            }
            BandwidthResults = if ($IncludeBandwidthTest -and $BandwidthResults.Count -gt 0) { 
                $BandwidthResults | ForEach-Object {
                    $result = $_ | ConvertTo-Json -Depth 3 | ConvertFrom-Json
                    # Remove timestamp from result
                    $result.PSObject.Properties.Remove('Timestamp')
                    $result
                }
            } else { $null }
            AuthenticationStatus = if ($IncludeAuthentication) { $AuthenticationStatus } else { $null }
        }
        
        $jsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Host "`nResults exported to: $OutputPath" -ForegroundColor Green
    }
    
    "HTML" {
        if (-not $OutputPath) {
            # Use parent directory as default output location
            $parentPath = Split-Path (Get-Location) -Parent
            if (Test-Path $parentPath) {
                $OutputPath = Join-Path $parentPath "M365-Connection-Test.html"
            }
            else {
                $OutputPath = Join-Path (Get-Location) "M365-Connection-Test.html"
            }
        }
        
        # Ensure the output directory exists
        $outputDirectory = Split-Path $OutputPath -Parent
        if ([string]::IsNullOrEmpty($outputDirectory)) {
            # If no directory specified, use current directory
            $outputDirectory = Get-Location
        }
        if (-not (Test-Path $outputDirectory)) {
            New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
        }
        
        # Generate HTML report
        New-HTMLReport -TestResults $TestResults -BandwidthResults $BandwidthResults -Summary $Summary -AuthenticationStatus $AuthenticationStatus -OutputPath $OutputPath -IncludeBandwidthTest $IncludeBandwidthTest -IncludeAuthentication $IncludeAuthentication
        
        Write-Host "`nHTML report exported to: $OutputPath" -ForegroundColor Green
        
        # Auto-open the HTML report in the default browser
        try {
            Start-Process $OutputPath
            Write-Host "Opening HTML report in your default browser..." -ForegroundColor Green
        }
        catch {
            Write-Host "HTML report saved. You can open it manually: $OutputPath" -ForegroundColor Yellow
        }
    }
    
    default {
        # Console output is already displayed above
    }
}

# Recommendations
Write-Host "`n=== Recommendations ===" -ForegroundColor Cyan

if ($Summary.FailedConnections -gt 0) {
    Write-Host "• Some connections failed. Check firewall settings and proxy configuration." -ForegroundColor Yellow
    Write-Host "• Verify DNS settings and ensure M365 URLs are not blocked." -ForegroundColor Yellow
}

if ($Summary.AverageLatency -gt 200) {
    Write-Host "• High latency detected. Consider optimizing network routing." -ForegroundColor Yellow
    Write-Host "• Check for network congestion or bandwidth limitations." -ForegroundColor Yellow
}

if ($IncludeAuthentication) {
    $authenticatedServices = ($AuthenticationStatus.GetEnumerator() | Where-Object { $_.Value -eq $true }).Count
    $totalServices = $AuthenticationStatus.Count
    
    if ($authenticatedServices -lt $totalServices) {
        Write-Host "• Consider installing missing PowerShell modules for complete testing:" -ForegroundColor Yellow
        Write-Host "  - Microsoft.Graph: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Gray
        Write-Host "  - ExchangeOnlineManagement: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Gray
        Write-Host "  - Microsoft.Online.SharePoint.PowerShell: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser" -ForegroundColor Gray
        Write-Host "  - MicrosoftTeams: Install-Module MicrosoftTeams -Scope CurrentUser" -ForegroundColor Gray
    }
    
    Write-Host "• Use -IncludeAuthentication for enhanced service testing with real API calls." -ForegroundColor Cyan
}
else {
    Write-Host "• Run with -IncludeAuthentication for enhanced testing including authenticated API calls." -ForegroundColor Cyan
}

Write-Host "• For detailed M365 network guidance, visit: https://aka.ms/m365networkconnectivity" -ForegroundColor Cyan
Write-Host "• Consider using Microsoft 365 connectivity test tool: https://connectivity.office.com" -ForegroundColor Cyan

# Cleanup
if ($DetailedLogging) {
    Stop-Transcript
    Write-Host "`nDetailed log saved to: $LogPath" -ForegroundColor Green
}

# Upload results if requested
if ($UploadResults) {
    Write-Host "`n=== Uploading Results to CIAOPS ===" -ForegroundColor Cyan
    Write-Host "Preparing to upload test results to CIAOPS secure Azure Blob Storage..." -ForegroundColor Yellow
    
    # Show upload statistics if requested
    if ($ShowUploadStats) {
        Write-Host "`n=== Upload Statistics ===" -ForegroundColor Cyan
        Write-Host "✅ Unlimited uploads enabled - no restrictions!" -ForegroundColor Green
        Write-Host ""
    }
    
    # Check if we have a file to upload
    if (-not $OutputPath -or -not (Test-Path $OutputPath)) {
        # Generate HTML report for upload if we don't have one
        if (-not $OutputPath) {
            $OutputPath = Join-Path $env:TEMP "m365-connection-test-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').html"
        }
        
        Write-Host "Generating HTML report for upload..." -ForegroundColor Yellow
        
        # Ensure the output directory exists
        $outputDirectory = Split-Path $OutputPath -Parent
        if (-not (Test-Path $outputDirectory)) {
            New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
        }
        
        # Generate HTML report
        New-HTMLReport -TestResults $TestResults -BandwidthResults $BandwidthResults -Summary $Summary -AuthenticationStatus $AuthenticationStatus -OutputPath $OutputPath -IncludeBandwidthTest $IncludeBandwidthTest -IncludeAuthentication $IncludeAuthentication
    }
    
    # Attempt upload
    $uploadResult = Upload-ToAzureBlob -FilePath $OutputPath -SasUrl $AzureUploadConfig.SasUrl
    
    if ($uploadResult.Success) {
        Write-Host "`n🎉 Results uploaded successfully to CIAOPS!" -ForegroundColor Green
        
        Write-Host "`n📧 Your test results have been securely uploaded to CIAOPS for analysis." -ForegroundColor Green
        Write-Host "   This helps improve M365 connectivity recommendations and support." -ForegroundColor White
        
    } else {
        Write-Host "`n❌ Upload failed: $($uploadResult.Error)" -ForegroundColor Red
        Write-Host "Please check your network connectivity and try again." -ForegroundColor Yellow
        Write-Host "If the problem persists, contact: support@ciaops.com" -ForegroundColor Yellow
    }
} else {
    # Offer upload option if user didn't specify -UploadResults
    Write-Host "`n=== Share Your Results ===" -ForegroundColor Cyan
    Write-Host "Help improve M365 connectivity insights by sharing your test results with CIAOPS." -ForegroundColor White
    Write-Host "Your data is uploaded securely and anonymously to our Azure Blob Storage." -ForegroundColor Gray
    Write-Host ""
    
    $uploadChoice = Read-Host "Would you like to upload your test results to CIAOPS? (Y/N) [Default: Y]"
    
    # Default to "Y" if user just presses Enter (empty string)
    if ([string]::IsNullOrWhiteSpace($uploadChoice) -or $uploadChoice -match '^[Yy]') {
        Write-Host "`n=== Uploading Results to CIAOPS ===" -ForegroundColor Cyan
        Write-Host "Preparing to upload test results to CIAOPS secure Azure Blob Storage..." -ForegroundColor Yellow
        
        # Check if we have a file to upload
        if (-not $OutputPath -or -not (Test-Path $OutputPath)) {
            # Generate HTML report for upload if we don't have one
            if (-not $OutputPath) {
                $OutputPath = Join-Path $env:TEMP "m365-connection-test-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').html"
            }
            
            Write-Host "Generating HTML report for upload..." -ForegroundColor Yellow
            
            # Ensure the output directory exists
            $outputDirectory = Split-Path $OutputPath -Parent
            if (-not (Test-Path $outputDirectory)) {
                New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
            }
            
            # Generate HTML report
            New-HTMLReport -TestResults $TestResults -BandwidthResults $BandwidthResults -Summary $Summary -AuthenticationStatus $AuthenticationStatus -OutputPath $OutputPath -IncludeBandwidthTest $IncludeBandwidthTest -IncludeAuthentication $IncludeAuthentication
        }
        
        # Attempt upload
        $uploadResult = Upload-ToAzureBlob -FilePath $OutputPath -SasUrl $AzureUploadConfig.SasUrl
        
        if ($uploadResult.Success) {
            Write-Host "`n🎉 Results uploaded successfully to CIAOPS!" -ForegroundColor Green
            
            Write-Host "`n Your test results have been securely uploaded to CIAOPS for analysis." -ForegroundColor Green
            Write-Host "   This helps improve M365 connectivity recommendations and support." -ForegroundColor White
            
        } else {
            Write-Host "`n❌ Upload failed: $($uploadResult.Error)" -ForegroundColor Red
            Write-Host "Please check your network connectivity and try again." -ForegroundColor Yellow
            Write-Host "If the problem persists, contact: support@ciaops.com" -ForegroundColor Yellow
        }
    } else {
        Write-Host "`nNo problem! Your test results remain on your local machine." -ForegroundColor Green
        Write-Host "Tip: Use -UploadResults parameter to upload automatically in future runs." -ForegroundColor Cyan
    }
}

Write-Host "`nTest completed successfully!" -ForegroundColor Green

# Return results for programmatic use
return @{
    Summary = $Summary
    Results = $TestResults
}
