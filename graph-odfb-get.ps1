param(                        
    [switch]$debug = $false, ## if -debug parameter don't prompt for input
    [switch]$csv = $false, ## export to CSV
    [switch]$prompt = $false    ## if -prompt parameter used user prompted for input
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on ODFB for tenant
Source - https://github.com/directorcia/Office365/blob/master/graph-odfb-get.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Report-on-OneDrive-for-a-tenant

Prerequisites = 1
1. Ensure the MS Graph module is installed

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"
$outputFile = "..\odfb-summary.csv"

if ($debug) {
    # create a log file of process if option enabled
    write-host "Script activity logged at .\graph-odfb-get.txt"
    start-transcript ".\graph-odfb-get.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

Clear-Host
write-host -foregroundcolor $systemmessagecolor "ODFB report script - Started`n"
write-host -foregroundcolor $processmessagecolor "Connect to MS Graph"
# https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
$scopes = "User.ReadBasic.All", "User.Read.All", "User.ReadWrite.All","User.Read",`
          "Directory.Read.All", "Directory.ReadWrite.All",`
          "Files.Read"
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

$Url = 'https://graph.microsoft.com/beta/users?$select=displayName,userPrincipalName,id&$top=999'
write-host -foregroundcolor $processmessagecolor "Make Graph request for all users"
try {
    $results = (Invoke-MgGraphRequest -Uri $Url -Method GET).value
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n"$_.Exception.Message
    exit (0)
}
$userloginsummary = @()
foreach ($result in $results) {
    $userloginSummary += [pscustomobject]@{                                                  ## Build array item
        Displayname = $result.displayname
        UPN         = $result.userprincipalname
        Id          = $result.id
    }
}

$userloginSummary | sort-object upn | select-object displayname, upn | format-table

write-host -foregroundcolor $processmessagecolor "Make Graph request for user ODFB"
$odfbsummary = @()
foreach ($user in $userloginsummary) {
    $found = $true
    write-host -foregroundcolor $processmessagecolor "`nChecking user = $($user.UPN)"
    $Url = "https://graph.microsoft.com/beta/users/$($user.id)/drive/"
    try {
        $result = (Invoke-MgGraphRequest -Uri $Url -Method GET)
    }
    catch {
        Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
        $found = $false
    }
    if ($found) {
        write-host "  Found ="$result.weburl.replace("https://ciaopslabs-my.sharepoint.com/personal","")
        $odfbSummary += [pscustomobject]@{                                                  ## Build array item
            Displayname = $user.displayname
            UPN         = $user.UPN
            UserId      = $user.id
            ODFBweburl  = $result.weburl
            odfbId      = $result.id
            state       = $result.quota.state
            total       = $result.quota.total
            remaining   = $result.quota.remaining
            used        = $result.quota.used
            deleted     = $result.quota.deleted        
        }
    }
}
write-host -foregroundcolor $processmessagecolor "`nSummary of ODFB usage"
foreach ($odfb in $odfbsummary) {
    write-host
    write-host -ForegroundColor Cyan  $odfb.displayname," - "$odfb.upn
    write-host "  Total(MB)   = "('{0:N2}' -f ($odfb.total/1MB))
    write-host "  Used(MB)    = "('{0:N2}' -f ($odfb.used/1MB))
    write-host "  Deleted(MB) = "('{0:N2}' -f ($odfb.deleted/1MB))
}

if ($csv) {
    write-host -foregroundcolor $processmessagecolor "`nOutput to CSV", $outputFile
    $odfbSummary | export-csv $outputFile -NoTypeInformation
}

write-host -foregroundcolor $systemmessagecolor "`nGraph user last login script - Finished"
if ($debug) {
    Stop-Transcript | Out-Null      
}
