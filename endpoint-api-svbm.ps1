<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Connect to the Defender for Endpoint API and return software vulnerabilities by machine

Source - https://github.com/directorcia/Office365/blob/master/endpoint-api-svbm.ps1
Documentation - https://blog.ciaops.com/2021/06/15/using-the-defender-for-endpoint-api-and-powershell/

Prerequisites = 1
1. Azure AD app setup per - https://blog.ciaops.com/2019/04/17/using-interactive-powershell-to-access-the-microsoft-graph/
2. Pass Client Id, Tenant Id and Client Secret as parameters when running the script

More scripts available by joining http://www.ciaopspatron.com

#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [SecureString]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [string]$CsvOutput
)

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor   = "red"

Clear-Host

Write-Host -ForegroundColor $systemmessagecolor "Script started`n"

try {
    # Decode the secure client secret for the token request body
    $bstr          = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
    $plainSecret   = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)

    # Construct token URI and body
    $tokenUri  = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $tokenBody = @{
        client_id     = $ClientId
        scope         = "https://api.securitycenter.microsoft.com/.default"
        client_secret = $plainSecret
        grant_type    = "client_credentials"
    }

    Write-Host -ForegroundColor $processmessagecolor "Getting OAuth 2.0 token"
    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUri -ContentType "application/x-www-form-urlencoded" -Body $tokenBody -ErrorAction Stop
    $token = $tokenResponse.access_token
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "Failed to acquire token: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($bstr) { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

$headers    = @{ Authorization = "Bearer $token" }
$apiUri     = "https://api.securitycenter.microsoft.com/api/machines/SoftwareVulnerabilitiesByMachine"
$allResults = [System.Collections.Generic.List[object]]::new()

try {
    Write-Host -ForegroundColor $processmessagecolor "Querying Defender for Endpoint API (with pagination)"

    do {
        $response = Invoke-RestMethod -Method Get -Uri $apiUri -ContentType "application/json" -Headers $headers -ErrorAction Stop
        if ($response.value) {
            $allResults.AddRange($response.value)
        }
        $apiUri = $response.'@odata.nextLink'
    } while ($apiUri)
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "API query failed: $($_.Exception.Message)"
    exit 1
}

Write-Host -ForegroundColor $processmessagecolor "Total records returned: $($allResults.Count)"

$selected = $allResults |
    Select-Object DeviceName, CveId, LastSeenTimestamp, SoftwareName, SoftwareVendor, SoftwareVersion, VulnerabilitySeverityLevel |
    Sort-Object DeviceName, LastSeenTimestamp

$selected | Format-Table -AutoSize

if ($CsvOutput) {
    try {
        $selected | Export-Csv -Path $CsvOutput -NoTypeInformation -Encoding UTF8
        Write-Host -ForegroundColor $processmessagecolor "Results exported to: $CsvOutput"
    }
    catch {
        Write-Host -ForegroundColor $errormessagecolor "CSV export failed: $($_.Exception.Message)"
    }
}

Write-Host -ForegroundColor $systemmessagecolor "`nScript Completed`n"
