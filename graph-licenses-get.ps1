<#
.SYNOPSIS
    Retrieves and reports Microsoft 365 license information from a tenant using Microsoft Graph API.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves subscribed SKU (license) information. Displays available,
    assigned, and unassigned license counts. Optionally exports license data to CSV format.
    Automatically retrieves product names from CIAOPS community repository.

.PARAMETER Debug
    If specified, logs script activity to a transcript file.

.PARAMETER Csv
    If specified, exports license data to a CSV file.

.PARAMETER Prompt
    If specified, prompts user to confirm the connected account before proceeding.

.PARAMETER OutputFile
    Path to the output CSV file. Default is "..\/graph-licenses.csv".

.EXAMPLE
    .\graph-licenses-get.ps1 -Csv -Debug

.NOTES
    Prerequisites: MS Graph PowerShell module must be installed
    Requires: LicenseAssignment.Read.All scope
#>

param(
    [switch]$debug = $false,
    [switch]$csv = $false,
    [switch]$prompt = $false,
    
    [ValidateNotNullOrEmpty()]
    [string]$OutputFile = "..\graph-licenses.csv"
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on licenses for tenant
Source - https://github.com/directorcia/Office365/blob/master/graph-licenses-get.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Report-tenant-licenses

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
    write-host "Script activity logged at .\graph-licenses-get.txt"
    start-transcript ".\graph-licenses-get.txt" | Out-Null
}

try {
    Clear-Host
    write-host -foregroundcolor $systemmessagecolor "Tenant license report script - Started`n"
    write-host -foregroundcolor $processmessagecolor "Connect to MS Graph"
    $scopes = "LicenseAssignment.Read.All"
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

    Write-host -ForegroundColor $processmessagecolor "Retrieving product codes from repository..."
    try {
        $query = Invoke-WebRequest -Method GET -ContentType "application/json" -Uri "https://raw.githubusercontent.com/directorcia/bp/refs/heads/main/skus.json" -UseBasicParsing -ErrorAction Stop
        $skulist = $query.Content | ConvertFrom-Json
        write-host -foregroundcolor $processmessagecolor "Product codes retrieved successfully"
    }
    catch {
        Write-Host -ForegroundColor $errormessagecolor "Error retrieving product codes: $($_.Exception.Message)"
        $skulist = @{}
    }

    if ($prompt) { Read-Host -Prompt "`n[PROMPT] -- Press Enter to continue" | Out-Null }

    $url = "https://graph.microsoft.com/beta/subscribedSkus"
    write-host -foregroundcolor $processmessagecolor "Retrieving license information from Graph API..."
    
    try {
        $results = (Invoke-MgGraphRequest -Uri $url -Method GET -ErrorAction Stop).value
        write-host -foregroundcolor $processmessagecolor "Retrieved $($results.Count) license(s)"
    }
    catch {
        Write-Host -ForegroundColor $errormessagecolor "Error retrieving licenses: $($_.Exception.Message)"
        throw
    }

    $licensesummary = @()
    foreach ($result in $results) {
        $partnumber = $result.skupartnumber
        $unassigned = $result.prepaidunits.enabled - $result.consumedunits
        
        $licenseSummary += [pscustomobject]@{
            License   = $result.skupartnumber
            Name      = $skulist.$partnumber
            Available = $result.prepaidunits.enabled
            Assigned  = $result.consumedunits
            Unassigned = $unassigned
        }
    }

    write-host -foregroundcolor $processmessagecolor "`nProcessing $($licensesummary.Count) license records...`n"
    
    $licenseSummary | Sort-Object License | Select-Object License, Name, Available, Assigned, Unassigned | Format-Table

    if ($csv) {
        write-host -foregroundcolor $processmessagecolor "Exporting $($licensesummary.Count) licenses to CSV: $OutputFile"
        $licenseSummary | Export-Csv $OutputFile -NoTypeInformation -Encoding UTF8 -Force
        write-host -foregroundcolor $processmessagecolor "CSV export completed successfully"
    }

    write-host -foregroundcolor $systemmessagecolor "`nGraph license script - Finished"
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
        # Ignore disconnect failures
    }

    if ($debug) {
        Stop-Transcript | Out-Null
    }
}
