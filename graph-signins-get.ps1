param(                        
    [switch]$debug = $false, ## if -debug parameter don't prompt for input
    [switch]$csv = $false, ## export to CSV
    [switch]$prompt = $false    ## if -prompt parameter used user prompted for input
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on licenses for tenant
Source - https://github.com/directorcia/Office365/blob/master/graph-signins-get.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Get-tenant-signins

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
$outputFile = "..\graph-signins.csv"

if ($debug) {
    # create a log file of process if option enabled
    write-host "Script activity logged at .\graph-signins-get.txt"
    start-transcript ".\graph-signins-get.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

Clear-Host
write-host -foregroundcolor $systemmessagecolor "Tenant signins report script - Started`n"
write-host -foregroundcolor $processmessagecolor "Connect to MS Graph"
$scopes = "AuditLog.Read.All","Directory.Read.All"
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


# Get all devices
# https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0&tabs=http
$Url = "https://graph.microsoft.com/beta/auditLogs/signIns"
write-host -foregroundcolor $processmessagecolor "Make Graph request for signins"
try {
    $results = (Invoke-MgGraphRequest -Uri $Url -Method GET).value
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n"$_.Exception.Message
    exit (0)
}
$signinsummary = @()
foreach ($result in $results) {
    $SigninSummary += [pscustomobject]@{                                                  ## Build array item
        ClientAppUsed     = $result.ClientAppUsed
        IpAddress         = $result.ipaddress
        IsINteractive     = $result.isinteractive
        UserPrincipalName = $result.UserPrincipalName
    }
}

# Output the Signins
$signinsummary | Format-Table ClientAppUsed, IpAddress, IsINteractive, UserPrincipalName

if ($csv) {
    write-host -foregroundcolor $processmessagecolor "`nOutput to CSV", $outputFile
    $signinsummary | export-csv $outputFile -NoTypeInformation
}

write-host -foregroundcolor $systemmessagecolor "`nGraph devices script - Finished"
if ($debug) {
    Stop-Transcript | Out-Null      
}
