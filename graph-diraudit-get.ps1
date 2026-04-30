<#
.SYNOPSIS
    Retrieves and reports directory audit records from a Microsoft 365 tenant using Microsoft Graph API.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves directory audit logs. Displays results in a formatted table
    and optionally exports to CSV. Handles pagination automatically for large result sets.

.PARAMETER Debug
    If specified, logs script activity to a transcript file.

.PARAMETER Csv
    If specified, exports audit records to a CSV file.

.PARAMETER Prompt
    If specified, prompts user to confirm the connected account before proceeding.

.PARAMETER PageSize
    Number of records to retrieve per API call. Default is 1000 (maximum).

.PARAMETER OutputFile
    Path to the output CSV file. Default is "..\graph-diraudit.csv".

.EXAMPLE
    .\graph-diraudit-get.ps1 -Csv -Debug

.NOTES
    Prerequisites: MS Graph PowerShell module must be installed
    Requires: AuditLog.Read.All and Directory.Read.All scopes
#>

param(
    [switch]$debug = $false,

    [switch]$csv = $false,
    [switch]$prompt = $false,

    [ValidateRange(1, 1000)]
    [int]$PageSize = 1000,

    [ValidateNotNullOrEmpty()]
    [string]$OutputFile = "..\graph-diraudit.csv"
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report directory audit for tenant
Source - https://github.com/directorcia/Office365/blob/master/graph-diraudit-get.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Report-directory-activity-in-a-tenant

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

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

## Validate output file path if CSV export is requested
if ($csv) {
    $outputDir = Split-Path -Path $OutputFile -Parent
    if (-not (Test-Path -Path $outputDir -PathType Container)) {
        Write-Host -ForegroundColor $errormessagecolor "Output directory does not exist: $outputDir"
        exit 1
    }
}

function Confirm-YesResponse {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return $Value.Trim() -match '^(?i:y|yes)$'
}

if ($debug) {
    # create a log file of process if option enabled
    write-host "Script activity logged at .\graph-diraudit-get.txt"
    start-transcript ".\graph-diraudit-get.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

try {
    write-host -foregroundcolor $systemmessagecolor "Tenant directory audit report script - Started`n"
    write-host -foregroundcolor $processmessagecolor "Connect to MS Graph"

    $scopes = "AuditLog.Read.All", "Directory.Read.All"
    Connect-MgGraph -Scopes $scopes -NoWelcome | Out-Null

    $graphcontext = Get-MgContext
    write-host -foregroundcolor $processmessagecolor "Connected account = $($graphcontext.Account)"

    if ($prompt) {
        do {
            $response = Read-Host -Prompt "`nIs this correct? [Y/N]"
        } until (-not [string]::IsNullOrWhiteSpace($response))

        if (-not (Confirm-YesResponse -Value $response)) {
            Disconnect-MgGraph | Out-Null
            write-host -foregroundcolor $warningmessagecolor "`n[001] Disconnected from current Graph environment. Re-run script to login to desired environment"
            exit 1
        }

        Read-Host -Prompt "`n[PROMPT] -- Press Enter to continue" | Out-Null
    }

    # Get all records from directory audit
    # https://learn.microsoft.com/en-us/graph/api/directoryaudit-list?view=graph-rest-1.0&tabs=http
    $url = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$top=$PageSize"
    $results = [System.Collections.ArrayList]::new()
    $pageCount = 0
    write-host -foregroundcolor $processmessagecolor "Retrieving directory audit records (page size: $PageSize)...`n"

    while ($null -ne $url) {
        try {
            $pageCount++
            $response = Invoke-MgGraphRequest -Uri $url -Method GET -ErrorAction Stop
            $pageResultCount = @($response.value).Count
            $results.AddRange([object[]]$response.value)
            write-host -foregroundcolor $processmessagecolor "[Page $pageCount] Retrieved $pageResultCount records (Total: $($results.Count))"

            $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
            if ($null -ne $nextLinkProperty -and -not [string]::IsNullOrWhiteSpace([string]$nextLinkProperty.Value)) {
                $url = [string]$nextLinkProperty.Value
            }
            else {
                $url = $null
            }
        }
        catch {
            Write-Host -ForegroundColor $errormessagecolor "Error retrieving page $pageCount from Graph API: $($_.Exception.Message)"
            throw
        }
    }

    if ($results.Count -eq 0) {
        Write-Host -ForegroundColor $warningmessagecolor "No directory audit records returned."
    }
    else {
        write-host -foregroundcolor $processmessagecolor "`nProcessing $($results.Count) audit records...`n"
        
        # Output the directory audit records sorted by newest first.
        $sortedResults = $results | Sort-Object ActivityDateTime -Descending
        $sortedResults |
            Select-Object LoggedByService, ActivityDisplayName, Result, OperationType, Category, ActivityDateTime |
            Format-Table -AutoSize

        if ($csv) {
            write-host -foregroundcolor $processmessagecolor "Exporting $($results.Count) records to CSV: $OutputFile"
            $sortedResults | Export-Csv $OutputFile -NoTypeInformation -Encoding UTF8 -Force
            write-host -foregroundcolor $processmessagecolor "CSV export completed successfully"
        }
    }

    write-host -foregroundcolor $systemmessagecolor "`nGraph directory audit script - Finished"
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`nError occurred during script execution:"
    Write-Host -ForegroundColor $errormessagecolor "  Exception: $($_.Exception.GetType().Name)"
    Write-Host -ForegroundColor $errormessagecolor "  Message: $($_.Exception.Message)"
    Write-Host -ForegroundColor $errormessagecolor "  Line: $($_.InvocationInfo.ScriptLineNumber)"
    exit 1
}
finally {
    try {
        Disconnect-MgGraph | Out-Null
        write-host -foregroundcolor $processmessagecolor "Disconnected from Graph"
    }
    catch {
        # Ignore disconnect failures so script can complete with original error state.
    }

    if ($debug) {
        Stop-Transcript | Out-Null
    }
}
