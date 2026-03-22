param(                        
    [switch]$debug = $false, ## if -debug parameter don't prompt for input
    [switch]$csv = $false, ## export to CSV
    [switch]$prompt = $false    ## if -prompt parameter used user prompted for input
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on licenses for tenant
Source - https://github.com/directorcia/Office365/blob/master/graph-users-get.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Report-Tenant-Users

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
$outputFile = "..\graph-users.csv"
$scriptFailed = $false

if ($debug) {
    Write-Host "Script activity logged at .\graph-users-get.txt"
    try {
        Start-Transcript ".\graph-users-get.txt" | Out-Null
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
    Write-Host -ForegroundColor $systemmessagecolor "Tenant user report script - Started`n"

    if (-not (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue)) {
        throw "Microsoft Graph PowerShell SDK is not installed. Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    $requiredScopes = @("User.Read.All")
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

    # Get all users (paged)
    # https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
    $selectFields = "displayName,userPrincipalName,accountEnabled,userType"
    $url = "https://graph.microsoft.com/v1.0/users?`$select=$selectFields&`$top=999"
    Write-Host -ForegroundColor $processmessagecolor "Make Graph request for all users"

    $userSummary = [System.Collections.Generic.List[object]]::new()
    $pageCount = 0
    $totalRecords = 0

    do {
        $pageCount++
        $response = Invoke-WithRetry -ScriptBlock {
            Invoke-MgGraphRequest -Uri $url -Method GET
        }

        $results = @($response.value)
        $totalRecords += $results.Count

        foreach ($result in $results) {
            $null = $userSummary.Add([pscustomobject]@{
                DisplayName       = $result.displayName
                UserPrincipalName = $result.userPrincipalName
                AccountEnabled    = $result.accountEnabled
                UserType          = $result.userType
            })
        }

        Write-Host -ForegroundColor $processmessagecolor "Retrieved page $pageCount with $($results.Count) records"
        $url = $response.'@odata.nextLink'
    } while ($url)

    Write-Host -ForegroundColor $processmessagecolor "Retrieved $totalRecords total user records across $pageCount pages"

    # Output the users
    $userSummary | Sort-Object DisplayName | Format-Table DisplayName, UserPrincipalName, AccountEnabled, UserType

    if ($csv) {
        Write-Host -ForegroundColor $processmessagecolor "`nOutput to CSV", $outputFile
        $userSummary | Export-Csv $outputFile -NoTypeInformation -Encoding UTF8
    }

    Write-Host -ForegroundColor $systemmessagecolor "`nGraph users script - Finished"
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
