param(
    [switch]$debug = $false, ## if -debug parameter, log transcript

    [switch]$csv = $false, ## export to CSV
    [switch]$prompt = $false, ## if -prompt parameter used user prompted for input

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
    write-host -foregroundcolor $processmessagecolor "Connected account =", $graphcontext.Account

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
    $url = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$top=1000"
    $results = [System.Collections.ArrayList]::new()
    write-host -foregroundcolor $processmessagecolor "Make Graph request for audit records"

    while ($null -ne $url) {
        $response = Invoke-MgGraphRequest -Uri $url -Method GET
        foreach ($item in @($response.value)) {
            [void]$results.Add($item)
        }

        $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
        if ($null -ne $nextLinkProperty -and -not [string]::IsNullOrWhiteSpace([string]$nextLinkProperty.Value)) {
            $url = [string]$nextLinkProperty.Value
        }
        else {
            $url = $null
        }
    }

    if ($results.Count -eq 0) {
        Write-Host -ForegroundColor $warningmessagecolor "No directory audit records returned."
    }
    else {
        # Output the directory audit records sorted by newest first.
        $sortedResults = $results | Sort-Object ActivityDateTime -Descending
        $sortedResults |
            Select-Object LoggedByService, ActivityDisplayName, Result, OperationType, Category, ActivityDateTime |
            Format-Table -AutoSize

        if ($csv) {
            write-host -foregroundcolor $processmessagecolor "`nOutput to CSV", $OutputFile
            $sortedResults | Export-Csv $OutputFile -NoTypeInformation -Encoding UTF8
        }
    }

    write-host -foregroundcolor $systemmessagecolor "`nGraph directory audit script - Finished"
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n$($_.Exception.Message)"
    exit 1
}
finally {
    try {
        Disconnect-MgGraph | Out-Null
    }
    catch {
        # Ignore disconnect failures so script can complete with original error state.
    }

    if ($debug) {
        Stop-Transcript | Out-Null
    }
}
