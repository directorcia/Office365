param(
    [switch]$csv = $false,
    [switch]$debug = $false,
    [string]$outputFile = "..\intune-policy-report.csv"
)

<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Return the name of all policies configured in EndPoint Manager (Intune and Endpoint)
Source - https://github.com/directorcia/office365/blob/master/intune-policy-get.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Return-the-name-of-all-policies-configured-in-Endpoint-Manager-(Intune-and-Endpoint)

Prerequisites = 1
1. Ensure connected to Intune - Use https://github.com/directorcia/Office365/blob/master/Intune-connect.ps1

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"
$scriptFailed = $false

if ($debug) {
    Write-Host "Script activity logged at .\intune-policy-get.txt"
    try {
        Start-Transcript ".\intune-policy-get.txt" | Out-Null
    }
    catch {
        Write-Host -ForegroundColor $warningmessagecolor "Unable to start transcript logging: $($_.Exception.Message)"
    }
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$InitialBackoffSeconds = 2
    )

    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try {
            return & $ScriptBlock
        }
        catch {
            $attempt++
            if ($attempt -ge $MaxRetries) {
                throw
            }

            $waitSeconds = $InitialBackoffSeconds * [Math]::Pow(2, $attempt - 1)
            Write-Host -ForegroundColor $warningmessagecolor "Attempt $attempt failed. Retrying in $waitSeconds seconds..."
            Start-Sleep -Seconds $waitSeconds
        }
    }
}

function Get-PolicyObjects {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Category,
        [Parameter(Mandatory = $true)]
        [scriptblock]$Query
    )

    $result = [System.Collections.Generic.List[object]]::new()
    $policies = @()

    try {
        $policies = @(Invoke-WithRetry -ScriptBlock $Query)
    }
    catch {
        Write-Host -ForegroundColor $warningmessagecolor "Unable to retrieve ${Category}: $($_.Exception.Message)"
    }

    foreach ($policy in $policies) {
        if (-not [string]::IsNullOrWhiteSpace($policy.displayName)) {
            $null = $result.Add([pscustomobject]@{
                Category = $Category
                Name     = $policy.displayName
            })
        }
    }

    return $result
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

function Get-GraphCollection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri

    do {
        $response = Invoke-WithRetry -ScriptBlock {
            Invoke-MgGraphRequest -Uri $nextUri -Method GET
        }

        foreach ($item in @($response.value)) {
            $null = $items.Add($item)
        }

        $nextUri = $response.'@odata.nextLink'
    } while ($nextUri)

    return $items
}

try {
    Clear-Host
    Write-Host -ForegroundColor $systemmessagecolor "Script started"

    if (-not (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue)) {
        throw "Microsoft Graph PowerShell SDK is not installed. Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    $requiredScopes = @(
        "DeviceManagementConfiguration.Read.All",
        "DeviceManagementApps.Read.All"
    )

    $graphContext = Get-MgContext
    if (Test-GraphContextHasScopes -RequiredScopes $requiredScopes -Context $graphContext) {
        Write-Host -ForegroundColor $processmessagecolor "Using existing Intune Graph connection"
    }
    else {
        Write-Host -ForegroundColor $processmessagecolor "Connect to Intune Graph"
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome | Out-Null
        $graphContext = Get-MgContext
    }

    Write-Host -ForegroundColor $processmessagecolor "Connected account = $($graphContext.Account)"

    $allPolicies = [System.Collections.Generic.List[object]]::new()

    $compliancePolicies = Get-PolicyObjects -Category "Intune Compliance" -Query {
        Get-GraphCollection -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?`$select=displayName"
    }
    foreach ($item in $compliancePolicies) { $null = $allPolicies.Add($item) }

    $configurationPolicies = Get-PolicyObjects -Category "Intune Configuration" -Query {
        Get-GraphCollection -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$select=displayName"
    }
    foreach ($item in $configurationPolicies) { $null = $allPolicies.Add($item) }

    $appProtectionPolicies = Get-PolicyObjects -Category "Intune App Protection" -Query {
        Get-GraphCollection -Uri "https://graph.microsoft.com/beta/deviceAppManagement/managedAppPolicies?`$select=displayName"
    }
    foreach ($item in $appProtectionPolicies) { $null = $allPolicies.Add($item) }

    $appConfigurationPolicies = Get-PolicyObjects -Category "Intune App Configuration (Targeted)" -Query {
        Get-GraphCollection -Uri "https://graph.microsoft.com/beta/deviceAppManagement/targetedManagedAppConfigurations?`$select=displayName"
    }
    foreach ($item in $appConfigurationPolicies) { $null = $allPolicies.Add($item) }

    $endpointIntentPolicies = Get-PolicyObjects -Category "Endpoint Policies" -Query {
        Get-GraphCollection -Uri "https://graph.microsoft.com/beta/deviceManagement/intents?`$select=displayName"
    }
    foreach ($item in $endpointIntentPolicies) { $null = $allPolicies.Add($item) }

    Write-Host -ForegroundColor $processmessagecolor "`nPolicy summary"
    $allPolicies |
        Sort-Object Category, Name |
        Format-Table Category, Name

    if ($csv) {
        Write-Host -ForegroundColor $processmessagecolor "`nOutput to CSV $outputFile"
        $allPolicies |
            Sort-Object Category, Name |
            Export-Csv $outputFile -NoTypeInformation -Encoding UTF8
    }

    Write-Host -ForegroundColor $systemmessagecolor "`nScript finished"
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