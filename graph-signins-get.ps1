param(                        
    [switch]$debug = $false,    ## if -debug parameter don't prompt for input
    [switch]$csv = $false,      ## export to CSV
    [switch]$prompt = $false,   ## if -prompt parameter used user prompted for input
    [ValidateRange(1,100000)]
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
$scriptFailed = $false

if ($debug) {
    Write-Host "Script activity logged at .\graph-signins-get.txt"
    try {
        Start-Transcript ".\graph-signins-get.txt" | Out-Null
    }
    catch {
        Write-Host -ForegroundColor $warningmessagecolor "Unable to start transcript logging: $($_.Exception.Message)"
    }
}

function Invoke-WithRetry {
    param (
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$InitialBackoffSeconds = 2
    )

    $retryCount = 0
    while ($retryCount -lt $MaxRetries) {
        try {
            return & $ScriptBlock
        }
        catch {
            $retryCount++
            if ($retryCount -ge $MaxRetries) {
                throw
            }

            $backoffSeconds = $InitialBackoffSeconds * [Math]::Pow(2, $retryCount - 1)
            Write-Host -ForegroundColor $warningmessagecolor "Attempt $retryCount failed. Retrying in $backoffSeconds seconds..."
            Start-Sleep -Seconds $backoffSeconds
        }
    }
}

function Test-GraphContextHasScopes {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$RequiredScopes,
        $Context
    )

    if ($null -eq $Context -or [string]::IsNullOrWhiteSpace($Context.Account)) {
        return $false
    }

    $currentScopes = @($Context.Scopes)
    if ($currentScopes.Count -eq 0) {
        return $false
    }

    foreach ($scope in $RequiredScopes) {
        if ($scope -notin $currentScopes) {
            return $false
        }
    }

    return $true
}

try {
    Clear-Host
    Write-Host -ForegroundColor $systemmessagecolor "Tenant signins report script - Started`n"

    if (-not (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue)) {
        throw "Microsoft Graph PowerShell SDK is not installed. Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    $requiredScopes = @("AuditLog.Read.All")
    $graphcontext = Get-MgContext

    if (Test-GraphContextHasScopes -RequiredScopes $requiredScopes -Context $graphcontext) {
        Write-Host -ForegroundColor $processmessagecolor "Using existing MS Graph connection"
    }
    else {
        Write-Host -ForegroundColor $processmessagecolor "Connect to MS Graph"
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome | Out-Null
        $graphcontext = Get-MgContext
    }

    Write-Host -ForegroundColor $processmessagecolor "Connected account =", $graphcontext.Account

    if ($prompt) {
        do {
            $response = Read-Host -Prompt "`nIs this correct? [Y/N]"
        } until (-not [string]::IsNullOrWhiteSpace($response))

        if ($response -notmatch '^[Yy]$') {
            Disconnect-MgGraph | Out-Null
            Write-Host -ForegroundColor $warningmessagecolor "`n[001] Disconnected from current Graph environment. Re-run script to login to desired environment"
            exit 1
        }

        Read-Host -Prompt "`n[PROMPT] -- Press Enter to continue" | Out-Null
    }

    $overallTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $baseUrl = "https://graph.microsoft.com/beta/auditLogs/signIns"
    $selectFields = "clientAppUsed,ipAddress,isInteractive,userPrincipalName,createdDateTime,status"
    $url = "$baseUrl`?`$select=$selectFields&`$top=100"

    Write-Host -ForegroundColor $processmessagecolor "Make Graph request for signins with pagination"

    $signinSummary = [System.Collections.Generic.List[object]]::new()
    $pageCount = 0
    $totalRecords = 0

    do {
        $pageCount++
        $pageTimer = [System.Diagnostics.Stopwatch]::StartNew()

        $response = Invoke-WithRetry -ScriptBlock {
            Invoke-MgGraphRequest -Uri $url -Method GET
        }

        $results = @($response.value)
        $totalRecords += $results.Count

        foreach ($result in $results) {
            $parsedDate = [DateTimeOffset]::MinValue
            $localDateTime = $result.createdDateTime
            if ([DateTimeOffset]::TryParse($result.createdDateTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsedDate)) {
                $localDateTime = $parsedDate.ToLocalTime().DateTime
            }

            $statusText = "Unknown"
            if ($null -ne $result.status) {
                if ($result.status.errorCode -eq 0) {
                    $statusText = "Success"
                }
                elseif (-not [string]::IsNullOrWhiteSpace($result.status.failureReason)) {
                    $statusText = "Failure: $($result.status.failureReason)"
                }
                else {
                    $statusText = "Failure"
                }
            }

            $null = $signinSummary.Add([pscustomobject]@{
                ClientAppUsed     = $result.clientAppUsed
                IpAddress         = $result.ipAddress
                IsInteractive     = $result.isInteractive
                UserPrincipalName = $result.userPrincipalName
                CreatedDateTime   = $localDateTime
                Status            = $statusText
            })
        }

        $pageTimer.Stop()
        Write-Host -ForegroundColor $processmessagecolor "Retrieved page $pageCount with $($results.Count) records in $([Math]::Round($pageTimer.Elapsed.TotalSeconds, 2)) seconds"

        $url = $response.'@odata.nextLink'

        if (-not $allPages -and $pageCount -ge $maxPages) {
            Write-Host -ForegroundColor $warningmessagecolor "Reached maximum page limit of $maxPages. Use -allPages to retrieve all available data."
            $url = $null
        }
    } while ($url)

    $overallTimer.Stop()
    Write-Host -ForegroundColor $processmessagecolor "Retrieved $totalRecords total records across $pageCount pages"
    Write-Host -ForegroundColor $processmessagecolor "Total execution time: $([Math]::Round($overallTimer.Elapsed.TotalSeconds, 2)) seconds"

    $signinSummary | Format-Table ClientAppUsed, IpAddress, IsInteractive, UserPrincipalName, CreatedDateTime, Status

    if ($csv) {
        Write-Host -ForegroundColor $processmessagecolor "`nOutput to CSV", $outputFile
        $signinSummary | Export-Csv $outputFile -NoTypeInformation -Encoding UTF8
    }

    Write-Host -ForegroundColor $systemmessagecolor "`nGraph signins script - Finished"
}
catch {
    $scriptFailed = $true
    Write-Host -ForegroundColor $errormessagecolor "`n$($_.Exception.Message)"
}
finally {
    if ($debug) {
        try {
            Stop-Transcript | Out-Null
        }
        catch {
        }
    }
}

if ($scriptFailed) {
    exit 1
}
