param(                        
    # Enable transcript logging to a local text file for troubleshooting.
    [switch]$debug = $false, ## if -debug parameter don't prompt for input
    # Export final device results to CSV in addition to console output.
    [switch]$csv = $false, ## export to CSV
    # Ask the operator to confirm the signed-in Graph account before continuing.
    [switch]$prompt = $false    ## if -prompt parameter used user prompted for input
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on licenses for tenant
Source - https://github.com/directorcia/Office365/blob/master/graph-devices-get.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Report-tenant-devices

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
$outputFile = "..\graph-devices.csv"

# Retrieve all pages for a Graph collection endpoint by following @odata.nextLink.
function Get-GraphCollection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    # Use a generic list to avoid slow PowerShell array expansion in large tenants.
    $items = New-Object System.Collections.Generic.List[object]
    $nextLink = $Uri

    while ($null -ne $nextLink) {
        $response = Invoke-MgGraphRequest -Uri $nextLink -Method GET -OutputType PSObject
        if ($null -ne $response.value) {
            foreach ($entry in $response.value) {
                [void]$items.Add($entry)
            }
        }

        # When nextLink is null, paging is complete.
        $nextLink = $response.'@odata.nextLink'
    }

    return $items
}

if ($debug) {
    # Start a transcript so all host output can be reviewed after execution.
    write-host "Script activity logged at .\graph-devices-get.txt"
    start-transcript ".\graph-devices-get.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

Clear-Host
write-host -foregroundcolor $systemmessagecolor "Tenant device report script - Started`n"
write-host -foregroundcolor $processmessagecolor "Connect to MS Graph"
$scopes = "Device.Read.All"
# Fail early with a clear message if the Graph SDK is not available.
if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
    Write-Host -ForegroundColor $errormessagecolor "Microsoft Graph PowerShell SDK is not installed. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    exit 1
}

try {
    # Prompt for interactive auth and request only the scope needed by this report.
    connect-mggraph -scopes $scopes -nowelcome | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`nUnable to connect to Microsoft Graph."
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    exit 1
}

$graphcontext = Get-MgContext
write-host -foregroundcolor $processmessagecolor "Connected account =", $graphcontext.Account
if ($prompt) {
    # Optional safety confirmation to avoid running against the wrong tenant/account.
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


# Get all devices with paging support
# https://learn.microsoft.com/graph/api/device-list
$Url = "https://graph.microsoft.com/v1.0/devices?`$top=999"
write-host -foregroundcolor $processmessagecolor "Make Graph request for all devices"
try {
    # Collect every page of results so large tenants are fully represented.
    $results = Get-GraphCollection -Uri $Url
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`nFailed to retrieve devices from Microsoft Graph."
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    exit 1
}

$devicesummary = New-Object System.Collections.Generic.List[object]
foreach ($result in $results) {
    # Keep only key device properties for readable console/CSV reporting.
    $devicesummary.Add([pscustomobject]@{                                                  ## Build array item
        Displayname     = $result.displayname
        DeviceID        = $result.DeviceId
        OperatingSystem = $result.OperatingSystem
        Trusttype       = $result.Trusttype
    }) | Out-Null
}

# Output the devices
write-host -foregroundcolor $processmessagecolor "Devices returned:", $devicesummary.Count
# Sort by display name to make both console and CSV output easier to scan.
$devicesummary | Sort-Object DisplayName | Format-Table DisplayName, DeviceId, OperatingSystem, Trusttype

if ($csv) {
    # Optional export path for post-processing in Excel or other tooling.
    write-host -foregroundcolor $processmessagecolor "`nOutput to CSV", $outputFile
    $devicesummary | export-csv $outputFile -NoTypeInformation
}

write-host -foregroundcolor $systemmessagecolor "`nGraph devices script - Finished"
if ($debug) {
    # End transcript only when debug logging was enabled.
    Stop-Transcript | Out-Null      
}
