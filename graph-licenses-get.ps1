<#
.SYNOPSIS
    Retrieves and reports Microsoft 365 license information from a tenant using Microsoft Graph.

.DESCRIPTION
    Connects to Microsoft Graph, retrieves subscribed SKU (license) data, and displays available,
    assigned, and unassigned license counts. Optionally exports the results to CSV.
    Product names are resolved from the CIAOPS community SKU list when available.

.PARAMETER Debug
    If specified, logs script activity to a transcript file.

.PARAMETER Csv
    If specified, exports license data to a CSV file.

.PARAMETER Prompt
    If specified, prompts for confirmation before proceeding with the report.

.PARAMETER OutputFile
    Path to the output CSV file. Default is "..\graph-licenses.csv".

.EXAMPLE
    .\graph-licenses-get.ps1 -Csv -Debug

.NOTES
    Prerequisites: Microsoft Graph PowerShell module must be installed.
    Requires: LicenseAssignment.Read.All scope.
#>

param(
    [switch]$Debug = $false,
    [switch]$Csv = $false,
    [switch]$Prompt = $false,

    [ValidateNotNullOrEmpty()]
    [string]$OutputFile = "..\graph-licenses.csv"
)

<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on licenses for tenant
Source - https://github.com/directorcia/Office365/blob/master/graph-licenses-get.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Report-tenant-licenses

If you find value in this script please support the author of these scripts by:

- https://ko-fi.com/ciaops

or

- becoming a CIAOPS Patron: https://www.ciaops.com/patron
#>

$systemMessageColor = "Cyan"
$processMessageColor = "Green"
$errorMessageColor = "Red"
$warningMessageColor = "Yellow"

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Confirm-YesResponse {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return $Value.Trim() -match '^(?i:y|yes)$'
}

function Get-SkuDisplayName {
    param(
        [string]$PartNumber,
        $SkuList
    )

    if ([string]::IsNullOrWhiteSpace($PartNumber)) {
        return ""
    }

    if ($null -eq $SkuList) {
        return $PartNumber
    }

    if ($SkuList -is [System.Collections.IDictionary]) {
        if ($SkuList.Contains($PartNumber)) {
            return [string]$SkuList[$PartNumber]
        }

        return $PartNumber
    }

    if ($SkuList -is [pscustomobject]) {
        $property = $SkuList.PSObject.Properties[$PartNumber]
        if ($null -ne $property) {
            return [string]$property.Value
        }

        return $PartNumber
    }

    if ($SkuList -is [hashtable]) {
        if ($SkuList.ContainsKey($PartNumber)) {
            return [string]$SkuList[$PartNumber]
        }

        return $PartNumber
    }

    return $PartNumber
}

if ($Csv) {
    $outputDir = Split-Path -Path $OutputFile -Parent
    if ([string]::IsNullOrWhiteSpace($outputDir)) {
        $outputDir = (Get-Location).Path
    }

    if (-not (Test-Path -Path $outputDir -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
            Write-Host -ForegroundColor $processMessageColor "Created output directory: $outputDir"
        }
        catch {
            Write-Host -ForegroundColor $errorMessageColor "Unable to create output directory: $outputDir"
            exit 1
        }
    }
}

if ($Debug) {
    Write-Host "Script activity logged at .\graph-licenses-get.txt"
    Start-Transcript -Path ".\graph-licenses-get.txt" -Force | Out-Null
}

$connected = $false

try {
    Clear-Host
    Write-Host -ForegroundColor $systemMessageColor "Tenant license report script - Started`n"

    $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication
    if (-not $module) {
        throw "Microsoft Graph PowerShell module is not installed. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    Write-Host -ForegroundColor $processMessageColor "Connecting to Microsoft Graph"
    $scopes = "LicenseAssignment.Read.All"
    Connect-MgGraph -Scopes $scopes -NoWelcome | Out-Null

    $graphContext = Get-MgContext
    if (-not $graphContext) {
        throw "Failed to establish a Microsoft Graph connection."
    }

    $connected = $true
    Write-Host -ForegroundColor $processMessageColor "Connected account = $($graphContext.Account)"

    if ($Prompt) {
        do {
            $response = Read-Host -Prompt "`nIs this correct? [Y/N]"
        } until (-not [string]::IsNullOrWhiteSpace($response))

        if (-not (Confirm-YesResponse -Value $response)) {
            Disconnect-MgGraph | Out-Null
            Write-Host -ForegroundColor $warningMessageColor "`n[001] Disconnected from the current Graph environment. Re-run the script to connect to the desired environment."
            exit 1
        }

        Read-Host -Prompt "`n[PROMPT] -- Press Enter to continue" | Out-Null
    }

    Write-Host -ForegroundColor $processMessageColor "Retrieving product codes from repository..."
    $skuList = @{}
    try {
        $query = Invoke-WebRequest -Method GET -ContentType "application/json" -Uri "https://raw.githubusercontent.com/directorcia/bp/refs/heads/main/skus.json" -UseBasicParsing -ErrorAction Stop
        $skuData = $query.Content | ConvertFrom-Json

        if ($skuData -is [pscustomobject]) {
            foreach ($property in $skuData.PSObject.Properties) {
                $skuList[$property.Name] = [string]$property.Value
            }
        }
        elseif ($skuData -is [System.Collections.IDictionary]) {
            foreach ($key in $skuData.Keys) {
                $skuList[$key] = [string]$skuData[$key]
            }
        }

        Write-Host -ForegroundColor $processMessageColor "Product codes retrieved successfully"
    }
    catch {
        Write-Host -ForegroundColor $warningMessageColor "Unable to retrieve product codes: $($_.Exception.Message)"
        Write-Host -ForegroundColor $warningMessageColor "Falling back to SKU part numbers for display names."
    }

    if ($Prompt) {
        Read-Host -Prompt "`n[PROMPT] -- Press Enter to continue" | Out-Null
    }

    $url = "https://graph.microsoft.com/beta/subscribedSkus"
    Write-Host -ForegroundColor $processMessageColor "Retrieving license information from Microsoft Graph..."

    try {
        $results = (Invoke-MgGraphRequest -Uri $url -Method GET -ErrorAction Stop).value
        if ($null -eq $results) {
            $results = @()
        }

        Write-Host -ForegroundColor $processMessageColor "Retrieved $($results.Count) license(s)"
    }
    catch {
        Write-Host -ForegroundColor $errorMessageColor "Error retrieving licenses: $($_.Exception.Message)"
        throw
    }

    $licenseSummary = @()
    foreach ($result in $results) {
        $partNumber = [string]$result.skuPartNumber
        $availableUnits = 0
        $assignedUnits = 0

        if ($null -ne $result.prepaidUnits -and $null -ne $result.prepaidUnits.enabled) {
            $availableUnits = [int]$result.prepaidUnits.enabled
        }

        if ($null -ne $result.consumedUnits) {
            $assignedUnits = [int]$result.consumedUnits
        }

        $unassignedUnits = $availableUnits - $assignedUnits

        $licenseSummary += [pscustomobject]@{
            License    = $partNumber
            Name       = Get-SkuDisplayName -PartNumber $partNumber -SkuList $skuList
            Available  = $availableUnits
            Assigned   = $assignedUnits
            Unassigned = $unassignedUnits
        }
    }

    Write-Host -ForegroundColor $processMessageColor "`nProcessing $($licenseSummary.Count) license records...`n"
    $licenseSummary |
        Sort-Object License |
        Select-Object License, Name, Available, Assigned, Unassigned |
        Format-Table -AutoSize

    if ($Csv) {
        Write-Host -ForegroundColor $processMessageColor "Exporting $($licenseSummary.Count) licenses to CSV: $OutputFile"
        $licenseSummary | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8 -Force
        Write-Host -ForegroundColor $processMessageColor "CSV export completed successfully"
    }

    Write-Host -ForegroundColor $systemMessageColor "`nGraph license script - Finished"
}
catch {
    Write-Host -ForegroundColor $errorMessageColor "`nError occurred during script execution:"
    Write-Host -ForegroundColor $errorMessageColor "  Exception: $($_.Exception.GetType().Name)"
    Write-Host -ForegroundColor $errorMessageColor "  Message: $($_.Exception.Message)"
    Write-Host -ForegroundColor $errorMessageColor "  Line: $($_.InvocationInfo.ScriptLineNumber)"
    exit 1
}
finally {
    if ($connected) {
        try {
            Disconnect-MgGraph | Out-Null
            Write-Host -ForegroundColor $processMessageColor "Disconnected from Graph"
        }
        catch {
            # Ignore disconnect failures
        }
    }

    if ($Debug) {
        Stop-Transcript | Out-Null
    }
}
