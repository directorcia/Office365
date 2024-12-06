param(                        
    [switch]$debug = $false,    ## if -debug parameter don't prompt for input
    [switch]$csv = $false,      ## export to CSV
    [switch]$prompt = $false    ## if -prompt parameter used user prompted for input
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on ODFB for tenant
Source - https://github.com/directorcia/Office365/blob/master/graph-licenses-get.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Report-tenant-licenses

Prerequisites = 1
1. Ensure the MS Graph module is installed

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"
$outputFile = "..\graph-licenses.csv"

if ($debug) {
    # create a log file of process if option enabled
    write-host "Script activity logged at .\graph-licenses-get.txt"
    start-transcript ".\graph-licenses-get.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

Clear-Host
write-host -foregroundcolor $systemmessagecolor "Tenant license report script - Started`n"
write-host -foregroundcolor $processmessagecolor "Connect to MS Graph"
# https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
$scopes = "LicenseAssignment.Read.All"
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

$Url = "https://graph.microsoft.com/beta/subscribedSkus"
write-host -foregroundcolor $processmessagecolor "Make Graph request for all licenses"
try {
    $results = (Invoke-MgGraphRequest -Uri $Url -Method GET).value
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n"$_.Exception.Message
    exit (0)
}
$licensesummary = @()
foreach ($result in $results) {
    $licenseSummary += [pscustomobject]@{                                                  ## Build array item
        license   = $result.skupartnumber
        available = $result.prepaidunits.enabled
        assigned  = $result.consumedunits
    }
}

$licenseSummary | sort-object skupartnumber | select-object license,available,assigned | format-table

if ($csv) {
    write-host -foregroundcolor $processmessagecolor "`nOutput to CSV", $outputFile
    $licenseSummary | export-csv $outputFile -NoTypeInformation
}

write-host -foregroundcolor $systemmessagecolor "`nGraph license script - Finished"
if ($debug) {
    Stop-Transcript | Out-Null      
}
