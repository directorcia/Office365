param(                        
    [switch]$debug = $false,    ## if -debug parameter don't prompt for input
    [switch]$csv = $false,      ## export to CSV
    [switch]$prompt = $false    ## if -prompt parameter used user prompted for input
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
If ($prompt) { Read-Host -Prompt "`n[PROMPT] -- Press Enter to continue" }

# Make call out to CIAOPS BP repository to get a list of all the product codes and store in a variable called $skulist
Write-host -ForegroundColor $processmessagecolor "Get Product codes via web request"
try {
    $query = invoke-webrequest -method GET -ContentType "application/json" -uri https://raw.githubusercontent.com/directorcia/bp/refs/heads/main/skus.json -UseBasicParsing
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[001]", $_.Exception.Message
}
$skulist = $query.content | ConvertFrom-Json

If ($prompt) { Read-Host -Prompt "`n[PROMPT] -- Press Enter to continue" }

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
    $partnumber=$result.skupartnumber
    $licenseSummary += [pscustomobject]@{                                                  ## Build array item
        license   = $result.skupartnumber
        name      = $skulist.$partnumber
        available = $result.prepaidunits.enabled
        assigned  = $result.consumedunits
    }
}

$licenseSummary | sort-object skupartnumber | select-object License,Name,Available,Assigned | format-table

if ($csv) {
    write-host -foregroundcolor $processmessagecolor "`nOutput to CSV", $outputFile
    $licenseSummary | export-csv $outputFile -NoTypeInformation
}

write-host -foregroundcolor $systemmessagecolor "`nGraph license script - Finished"
if ($debug) {
    Stop-Transcript | Out-Null      
}
