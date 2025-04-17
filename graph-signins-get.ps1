param(                        
    [switch]$debug = $false,    ## if -debug parameter don't prompt for input
    [switch]$csv = $false,      ## export to CSV
    [switch]$prompt = $false,   ## if -prompt parameter used user prompted for input
    [int]$maxPages = 10,        ## maximum number of pages to retrieve (default: 10)
    [switch]$allPages = $false  ## retrieve all available pages, ignoring maxPages
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on signins for tenant
Source - https://github.com/directorcia/Office365/blob/master/graph-signins-get.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Get-tenant-signins

Prerequisites = 1
1. Ensure the MS Graph module is installed

If you find value in this script please support the author of these scripts by:

- https://ko-fi.com/ciaops

or 

- becoming a CIAOPS Patron: https://www.ciaops.com/patron

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"
$outputFile = "..\graph-signins.csv"

if ($debug) {
    # create a log file of process if option enabled
    write-host "Script activity logged at .\graph-signins-get.txt"
    start-transcript ".\graph-signins-get.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

Clear-Host
write-host -foregroundcolor $systemmessagecolor "Tenant signins report script - Started`n"
write-host -foregroundcolor $processmessagecolor "Connect to MS Graph"
$scopes = "AuditLog.Read.All","Directory.Read.All"
connect-mggraph -scopes $scopes -nowelcome | Out-Null
$graphcontext = Get-MgContext
write-host -foregroundcolor $processmessagecolor "Connected account =", $graphcontext.Account
if ($prompt) {
    do {
        $response = read-host -Prompt "`nIs this correct? [Y/N]"
    } until (-not [string]::isnullorempty($response))
    if ($response -ne "Y" -and $response -ne "y") {
        Disconnect-MgGraph | Out-Null
        write-host -foregroundcolor $warningmessagecolor "`n[001] Disconnected from current Graph environment. Re-run script to login to desired environment"
        exit 1
    }
    else {
        write-host
    }
}
If ($prompt) { Read-Host -Prompt "`n[PROMPT] -- Press Enter to continue" }

# Function to handle retries with exponential backoff
function Invoke-WithRetry {
    param (
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$InitialBackoffSeconds = 2
    )
    
    $retryCount = 0
    $success = $false
    $result = $null

    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            $result = & $ScriptBlock
            $success = $true
        }
        catch {
            $retryCount++            
            if ($retryCount -ge $MaxRetries) {
                Write-Host -ForegroundColor $errormessagecolor "Failed after $MaxRetries retries: $_"
                throw
            }
            
            $backoffSeconds = $InitialBackoffSeconds * [Math]::Pow(2, $retryCount - 1)
            Write-Host -ForegroundColor $warningmessagecolor "Attempt $retryCount failed. Retrying in $backoffSeconds seconds..."
            Start-Sleep -Seconds $backoffSeconds
        }
    }
    
    return $result
}

# Start timing execution
$overallTimer = [System.Diagnostics.Stopwatch]::StartNew()

# Get signins with pagination and field selection for better performance
# https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0&tabs=http
$baseUrl = "https://graph.microsoft.com/beta/auditLogs/signIns"
$selectFields = "clientAppUsed,ipAddress,isInteractive,userPrincipalName,createdDateTime,status"
$url = "$baseUrl`?`$select=$selectFields&`$top=100"
write-host -foregroundcolor $processmessagecolor "Make Graph request for signins with pagination"

$signinsummary = @()
$pageCount = 0
$totalRecords = 0

do {
    $pageTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $pageCount++
    
    try {
        $response = Invoke-WithRetry -ScriptBlock {
            Invoke-MgGraphRequest -Uri $url -Method GET
        }
        
        $results = $response.value
        $totalRecords += $results.Count
        
        foreach ($result in $results) {
            # Convert UTC time to local time using InvariantCulture to handle different regional settings
            try {
                # Try to parse using DateTime.ParseExact with ISO 8601 format that Graph API typically returns
                $utcDateTime = [DateTime]::ParseExact($result.createdDateTime, "yyyy-MM-ddTHH:mm:ss.fffffffZ", [System.Globalization.CultureInfo]::InvariantCulture)
                $localDateTime = $utcDateTime.ToLocalTime()
            }
            catch {
                # Fallback method if the format is different
                try {
                    # Try parsing with a more flexible approach
                    $utcDateTime = [DateTime]::Parse($result.createdDateTime, [System.Globalization.CultureInfo]::InvariantCulture)
                    $localDateTime = $utcDateTime.ToLocalTime()
                }
                catch {
                    # If all parsing fails, just use the original string
                    Write-Host -ForegroundColor $warningmessagecolor "Could not parse date: $($result.createdDateTime). Using as-is."
                    $localDateTime = $result.createdDateTime
                }
            }
            
            $SigninSummary += [pscustomobject]@{                                                  ## Build array item
                ClientAppUsed     = $result.ClientAppUsed
                IpAddress         = $result.ipaddress
                IsInteractive     = $result.isinteractive
                UserPrincipalName = $result.UserPrincipalName
                CreatedDateTime   = $localDateTime
                Status            = if ($result.status.errorCode -eq 0) { "Success" } else { "Failure: $($result.status.failureReason)" }
            }
        }
        
        $pageTimer.Stop()
        write-host -foregroundcolor $processmessagecolor "Retrieved page $pageCount with $($results.Count) records in $($pageTimer.Elapsed.TotalSeconds) seconds"
        
        # Get URL for next page if it exists
        $url = $response.'@odata.nextLink'
    }
    catch {
        Write-Host -ForegroundColor $errormessagecolor "`n"$_.Exception.Message
        exit (0)
    }

    # Check if we've reached the maximum number of pages and we're not retrieving all pages
    if (-not $allPages -and $pageCount -ge $maxPages) {
        write-host -foregroundcolor $warningmessagecolor "Reached maximum page limit of $maxPages. Use -allPages to retrieve all available data."
        $url = $null  # Clear the URL to stop the loop
        break
    }
} while ($url)

write-host -foregroundcolor $processmessagecolor "Retrieved $totalRecords total records across $pageCount pages"

# Stop the overall timer and display total execution time
$overallTimer.Stop()
write-host -foregroundcolor $processmessagecolor "Total execution time: $($overallTimer.Elapsed.TotalSeconds) seconds"

# Process results more efficiently using a parallel approach for large datasets
if ($signinsummary.Count -gt 1000) {
    write-host -foregroundcolor $processmessagecolor "Processing large dataset with optimized approach"
    
    # Define the number of items to process per batch
    $batchSize = [Math]::Min(5000, $signinsummary.Count)
    
    # Process data in batches to avoid memory issues
    $processedResults = @()
    for ($i = 0; $i -lt $signinsummary.Count; $i += $batchSize) {
        $batch = $signinsummary | Select-Object -Skip $i -First $batchSize
        $processedResults += $batch
    }
    
    # Replace the original array with the processed results
    $signinsummary = $processedResults
}

# Output the Signins with selective properties for better performance
$signinsummary | Format-Table ClientAppUsed, IpAddress, IsInteractive, UserPrincipalName, CreatedDateTime, Status

if ($csv) {
    write-host -foregroundcolor $processmessagecolor "`nOutput to CSV", $outputFile
    $signinsummary | export-csv $outputFile -NoTypeInformation
}

write-host -foregroundcolor $systemmessagecolor "`nGraph devices script - Finished"
if ($debug) {
    Stop-Transcript | Out-Null      
}
