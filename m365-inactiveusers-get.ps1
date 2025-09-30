#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(HelpMessage="Enable detailed logging and transcript recording")]
    [switch]$DebugMode = $false,
    
    [Parameter(HelpMessage="Number of days to consider a user inactive (default: 90 days)")]
    [ValidateRange(1, 365)]
    [int]$InactiveDays = 90,
    
    [Parameter(HelpMessage="Include guest/external users in the analysis (enabled by default)")]
    [switch]$IncludeGuests = $true,
    
    [Parameter(HelpMessage="Generate HTML report")]
    [switch]$GenerateHtmlReport = $true,
    
    [Parameter(HelpMessage="Automatically open HTML report in browser when completed")]
    [switch]$OpenReportInBrowser = $true
)

<#CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.
Description: Check for inactive users in Microsoft 365 using Microsoft Graph
Documentation: https://github.com/directorcia/Office365/wiki/Microsoft-365-Inactive-Users-Check-Script
Source: https://github.com/directorcia/Office365/blob/master/m365-inactiveusers-get.ps1

#>

# Initialize
$LogPath = "..\check-inactive-users-$(Get-Date -Format 'yyyy-MM-dd-HHmm').txt"
$Script:UserResults = @()

function Write-LogMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet("Information", "Warning", "Error", "Success")]
        [string]$Level = "Information"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    try {
        $logEntry | Out-File -FilePath $LogPath -Append -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore log file errors
    }
    
    # Write to console with color coding
    switch ($Level) {
        "Information" { Write-Host $Message -ForegroundColor White }
        "Success" { Write-Host $Message -ForegroundColor Green }
        "Warning" { Write-Host $Message -ForegroundColor Yellow }
        "Error" { Write-Host $Message -ForegroundColor Red }
    }
}

function Initialize-GraphConnection {
    try {
        # Check if already connected
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context) {
            Write-LogMessage "✅ Already connected to Graph as: $($context.Account)" -Level Success
            return
        }
        
        Write-LogMessage "Connecting to Microsoft Graph..." -Level Information
        
        $requiredScopes = @(
            'User.Read.All',
            'Directory.Read.All'
        )
        
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        
        $context = Get-MgContext
        if ($context) {
            Write-LogMessage "✅ Successfully connected to Graph as: $($context.Account)" -Level Success
        } else {
            throw "Failed to establish Graph connection"
        }
    }
    catch {
        Write-LogMessage "Failed to initialize Graph connection: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Get-InactiveUsers {
    try {
        Write-LogMessage "🔍 Starting user analysis..." -Level Information
        
        # Calculate cutoff date
        $cutoffDate = (Get-Date).AddDays(-$InactiveDays)
        Write-LogMessage "Users inactive since: $($cutoffDate.ToString('yyyy-MM-dd'))" -Level Information
        
        # Get all users with required properties using Invoke-MgGraphRequest for better reliability
        Write-LogMessage "Fetching users from Microsoft Graph..." -Level Information
        
        $allUsers = @()
        $uri = "https://graph.microsoft.com/v1.0/users?`$select=id,userPrincipalName,displayName,userType,accountEnabled,createdDateTime,signInActivity,assignedLicenses,licenseAssignmentStates"
        
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            $allUsers += $response.value
            $uri = $response.'@odata.nextLink'
        } while ($uri)
        
        Write-LogMessage "Processing $($allUsers.Count) users..." -Level Information
        
        $stats = @{
            TotalUsers = $allUsers.Count
            ActiveUsers = 0
            InactiveUsers = 0
            NeverLoggedOn = 0
            UnlicensedUsers = 0
            GuestUsers = 0
        }
        
        foreach ($user in $allUsers) {
            # Skip guest users if not included
            if ($user.userType -eq "Guest" -and -not $IncludeGuests) {
                continue
            }
            
            # Determine activity status
            $lastSignIn = $null
            $activityStatus = "Unknown"
            $daysSinceLastSignIn = $null
            
            if ($user.signInActivity -and $user.signInActivity.lastSignInDateTime) {
                $lastSignIn = [DateTime]$user.signInActivity.lastSignInDateTime
                $daysSinceLastSignIn = ((Get-Date) - $lastSignIn).Days
                
                if ($lastSignIn -lt $cutoffDate) {
                    $activityStatus = "Inactive"
                    $stats.InactiveUsers++
                } else {
                    $activityStatus = "Active"
                    $stats.ActiveUsers++
                }
            } else {
                $activityStatus = "Never Logged On"
                $stats.NeverLoggedOn++
            }
            
            # Check license status - simplified approach
            $isUnlicensed = $true
            $licenseInfo = "No licenses assigned"
            
            if ($user.assignedLicenses -and $user.assignedLicenses.Count -gt 0) {
                $isUnlicensed = $false
                $licenseInfo = "Licensed ($($user.assignedLicenses.Count) licenses)"
            }
            
            if ($isUnlicensed) {
                $stats.UnlicensedUsers++
            }
            
            # Track guest users
            if ($user.userType -eq "Guest") {
                $stats.GuestUsers++
            }
            
            # Debug output for specific users
            if ($user.userPrincipalName -match "consulting|director|information") {
                Write-LogMessage "🔍 DEBUG: $($user.userPrincipalName)" -Level Warning
                Write-LogMessage "  Display Name: $($user.displayName)" -Level Information
                Write-LogMessage "  User Type: $($user.userType)" -Level Information
                Write-LogMessage "  Account Enabled: $($user.accountEnabled)" -Level Information
                Write-LogMessage "  Activity Status: $activityStatus" -Level Information
                Write-LogMessage "  Assigned Licenses Count: $(if ($user.assignedLicenses) { $user.assignedLicenses.Count } else { 0 })" -Level Information
                Write-LogMessage "  Is Unlicensed: $isUnlicensed" -Level Information
                Write-LogMessage "  License Info: $licenseInfo" -Level Information
            }
            
            # Create user result object
            $userResult = [PSCustomObject]@{
                DisplayName = $user.displayName
                UserPrincipalName = $user.userPrincipalName
                UserType = $user.userType
                AccountEnabled = $user.accountEnabled
                ActivityStatus = $activityStatus
                LastSignIn = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd HH:mm:ss') } else { "Never" }
                DaysSinceLastSignIn = $daysSinceLastSignIn
                IsUnlicensed = $isUnlicensed
                Licenses = $licenseInfo
                CreatedDate = if ($user.createdDateTime) { ([DateTime]$user.createdDateTime).ToString('yyyy-MM-dd') } else { "Unknown" }
            }
            
            $Script:UserResults += $userResult
        }
        
        Write-LogMessage "✅ User analysis completed" -Level Success
        Write-LogMessage "📊 SUMMARY STATISTICS:" -Level Success
        Write-LogMessage "  • Total Users Processed: $($stats.TotalUsers)" -Level Information
        Write-LogMessage "  • Active Users: $($stats.ActiveUsers)" -Level Success
        Write-LogMessage "  • Inactive Users: $($stats.InactiveUsers)" -Level Warning
        Write-LogMessage "  • Never Logged On: $($stats.NeverLoggedOn)" -Level Warning
        Write-LogMessage "  • Unlicensed Users: $($stats.UnlicensedUsers)" -Level Warning
        Write-LogMessage "  • External/Guest Users: $($stats.GuestUsers)" -Level Information
        
    }
    catch {
        Write-LogMessage "Error during user analysis: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Generate-HtmlReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )
    
    $reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $context = Get-MgContext
    $tenantInfo = if ($context) { $context.Account } else { "Unknown" }
    
    # Calculate statistics
    $activeUsers = $Script:UserResults | Where-Object { $_.ActivityStatus -eq "Active" }
    $inactiveUsers = $Script:UserResults | Where-Object { $_.ActivityStatus -eq "Inactive" }
    $neverLoggedUsers = $Script:UserResults | Where-Object { $_.ActivityStatus -eq "Never Logged On" }
    $unlicensedUsers = $Script:UserResults | Where-Object { $_.IsUnlicensed -eq $true }
    $guestUsers = $Script:UserResults | Where-Object { $_.UserType -eq "Guest" }
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Inactive Users Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f5f5f5; padding: 20px; }
        .container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); overflow: hidden; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p { opacity: 0.9; font-size: 1.1em; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; padding: 30px; background: #f8f9fa; }
        .stat-card { background: white; padding: 20px; border-radius: 8px; text-align: center; border-left: 4px solid #667eea; }
        .stat-card.warning { border-left-color: #ffc107; }
        .stat-card.danger { border-left-color: #dc3545; }
        .stat-card.success { border-left-color: #28a745; }
        .stat-card h3 { font-size: 2em; margin-bottom: 5px; }
        .stat-card p { color: #666; font-weight: 500; }
        .content { padding: 30px; }
        .search-box { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 6px; margin-bottom: 20px; font-size: 16px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #eee; }
        th { background: #f8f9fa; font-weight: 600; color: #495057; position: sticky; top: 0; }
        tr:hover { background: #f8f9fa; }
        .status-active { color: #28a745; font-weight: 600; }
        .status-inactive { color: #dc3545; font-weight: 600; }
        .status-never { color: #ffc107; font-weight: 600; }
        .license-yes { color: #28a745; }
        .license-no { color: #dc3545; }
        .footer { text-align: center; padding: 20px; background: #f8f9fa; color: #666; }
        .footer a { color: #667eea; text-decoration: none; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔍 Microsoft 365 Inactive Users Report</h1>
            <p>Generated on $reportDate</p>
            <p>Connected as: $tenantInfo</p>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card success">
                <h3>$($activeUsers.Count)</h3>
                <p>Active Users</p>
            </div>
            <div class="stat-card danger">
                <h3>$($inactiveUsers.Count)</h3>
                <p>Inactive Users</p>
            </div>
            <div class="stat-card warning">
                <h3>$($neverLoggedUsers.Count)</h3>
                <p>Never Logged In</p>
            </div>
            <div class="stat-card danger">
                <h3>$($unlicensedUsers.Count)</h3>
                <p>Unlicensed Users</p>
            </div>
            <div class="stat-card">
                <h3>$($guestUsers.Count)</h3>
                <p>External/Guest Users</p>
            </div>
            <div class="stat-card">
                <h3>$($Script:UserResults.Count)</h3>
                <p>Total Users</p>
            </div>
        </div>
        
        <div class="content">
            <input type="text" id="searchBox" class="search-box" placeholder="🔍 Search users..." onkeyup="filterTable()">
"@

    # Add user tables for each category
    $categories = @(
        @{ Name = "Active Users"; Users = $activeUsers; Icon = "✅" },
        @{ Name = "Inactive Users"; Users = $inactiveUsers; Icon = "⚠️" },
        @{ Name = "Never Logged In Users"; Users = $neverLoggedUsers; Icon = "🚫" },
        @{ Name = "Unlicensed Users"; Users = $unlicensedUsers; Icon = "💳" },
        @{ Name = "External/Guest Users"; Users = $guestUsers; Icon = "🌐" }
    )

    foreach ($category in $categories) {
        if ($category.Users -and $category.Users.Count -gt 0) {
            $html += @"
            
            <h2>$($category.Icon) $($category.Name) ($($category.Users.Count))</h2>
            <table id="usersTable$($category.Name.Replace(' ', ''))">
                <thead>
                    <tr>
                        <th>Display Name</th>
                        <th>User Principal Name</th>
                        <th>Status</th>
                        <th>Last Sign In</th>
                        <th>Days Since Last Sign In</th>
                        <th>Licensed</th>
                        <th>Licenses</th>
                        <th>User Type</th>
                        <th>Account Enabled</th>
                        <th>Created Date</th>
                    </tr>
                </thead>
                <tbody>
"@

            foreach ($user in $category.Users) {
                $statusClass = switch ($user.ActivityStatus) {
                    "Active" { "status-active" }
                    "Inactive" { "status-inactive" }
                    "Never Logged On" { "status-never" }
                    default { "" }
                }
                
                $licensedClass = if ($user.IsUnlicensed) { "license-no" } else { "license-yes" }
                $licensedText = if ($user.IsUnlicensed) { "❌ No" } else { "✅ Yes" }
                $enabledText = if ($user.AccountEnabled) { "✅ Yes" } else { "❌ No" }
                $daysDisplay = if ($user.DaysSinceLastSignIn) { $user.DaysSinceLastSignIn } else { "N/A" }

                $html += @"
                    <tr>
                        <td>$($user.DisplayName)</td>
                        <td>$($user.UserPrincipalName)</td>
                        <td class="$statusClass">$($user.ActivityStatus)</td>
                        <td>$($user.LastSignIn)</td>
                        <td>$daysDisplay</td>
                        <td class="$licensedClass">$licensedText</td>
                        <td>$($user.Licenses)</td>
                        <td>$($user.UserType)</td>
                        <td>$enabledText</td>
                        <td>$($user.CreatedDate)</td>
                    </tr>
"@
            }

            $html += @"
                </tbody>
            </table>
"@
        }
    }

    $html += @"
        </div>
        
        <div class="footer">
            <p>Generated by CIAOPS Inactive Users Check Script</p>
        </div>
    </div>
    
    <script>
        function filterTable() {
            var input, filter, tables, tr, td, i, txtValue;
            input = document.getElementById("searchBox");
            filter = input.value.toUpperCase();
            tables = document.getElementsByTagName("table");
            
            for (var t = 0; t < tables.length; t++) {
                tr = tables[t].getElementsByTagName("tr");
                for (i = 1; i < tr.length; i++) {
                    tr[i].style.display = "none";
                    td = tr[i].getElementsByTagName("td");
                    for (var j = 0; j < td.length; j++) {
                        if (td[j]) {
                            txtValue = td[j].textContent || td[j].innerText;
                            if (txtValue.toUpperCase().indexOf(filter) > -1) {
                                tr[i].style.display = "";
                                break;
                            }
                        }
                    }
                }
            }
        }
    </script>
</body>
</html>
"@

    $html | Out-File -FilePath $OutputPath -Encoding UTF8
}

function Export-Results {
    try {
        # Export to CSV
        $csvPath = "..\inactive-users-$(Get-Date -Format 'yyyy-MM-dd-HHmm').csv"
        Write-LogMessage "📄 Exporting results to CSV: $csvPath" -Level Information
        
        $Script:UserResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-LogMessage "✅ CSV export completed: $csvPath" -Level Success
        
        # Generate HTML report if requested
        $htmlPath = $null
        if ($GenerateHtmlReport) {
            $htmlPath = "..\inactive-users-report-$(Get-Date -Format 'yyyy-MM-dd-HHmm').html"
            Write-LogMessage "📊 Generating HTML report: $htmlPath" -Level Information
            
            Generate-HtmlReport -OutputPath $htmlPath
            Write-LogMessage "✅ HTML report generated: $htmlPath" -Level Success
        }
        
        return $htmlPath
        
    }
    catch {
        Write-LogMessage "Error during export: $($_.Exception.Message)" -Level Error
        return $null
    }
}

# Main execution
try {
    Write-LogMessage "=== CIAOPS Inactive Users Check - Started ===" -Level Information
    Write-LogMessage "Configuration:" -Level Information
    Write-LogMessage "  • Inactive Days Threshold: $InactiveDays days" -Level Information
    Write-LogMessage "  • Include Guests: $IncludeGuests" -Level Information
    Write-LogMessage "  • Generate HTML Report: $GenerateHtmlReport" -Level Information
    Write-LogMessage "  • Auto-open Report in Browser: $OpenReportInBrowser" -Level Information
    
    # Initialize Graph connection
    Initialize-GraphConnection
    
    # Get inactive users
    Get-InactiveUsers
    
    # Export results
    $htmlReportPath = Export-Results
    
    Write-LogMessage "=== CIAOPS Inactive Users Check - Completed ===" -Level Success
    
    # Open HTML report if it was generated and auto-open is enabled
    if ($OpenReportInBrowser -and $htmlReportPath -and (Test-Path $htmlReportPath)) {
        Write-LogMessage "🌐 Opening HTML report in default browser..." -Level Information
        try {
            Start-Process $htmlReportPath
            Write-LogMessage "✅ HTML report opened successfully" -Level Success
        }
        catch {
            Write-LogMessage "⚠️ Could not open HTML report automatically: $($_.Exception.Message)" -Level Warning
            Write-LogMessage "📁 HTML report location: $htmlReportPath" -Level Information
        }
    }
    elseif ($htmlReportPath -and (Test-Path $htmlReportPath)) {
        Write-LogMessage "📁 HTML report saved to: $htmlReportPath" -Level Information
    }
    
}
catch {
    Write-LogMessage "Script execution failed: $($_.Exception.Message)" -Level Error
    exit 1
}