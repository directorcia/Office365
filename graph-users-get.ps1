param(                        
    [switch]$debug = $false, ## if -debug parameter don't prompt for input
    [switch]$csv = $false, ## export to CSV
    [switch]$prompt = $false    ## if -prompt parameter used user prompted for input
)
<#CIAOPS

Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on licenses for tenant
Source - https://github.com/directorcia/Office365/blob/master/graph-users-get.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Report-Tenant-Users

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
$outputFile = "..\graph-users.csv"

if ($debug) {
    # create a log file of process if option enabled
    write-host "Script activity logged at .\graph-users-get.txt"
    start-transcript ".\graph-users-get.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

Clear-Host
write-host -foregroundcolor $systemmessagecolor "Tenant user report script - Started`n"
write-host -foregroundcolor $processmessagecolor "Connect to MS Graph"
$scopes = "User.ReadBasic.All, User.Read.All, User.ReadWrite.All, Directory.Read.All, Directory.ReadWrite.All"
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
# https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
$Url = "https://graph.microsoft.com/beta/users?&`$top=999" 
write-host -foregroundcolor $processmessagecolor "Make Graph request for all users"
try {
    $results = (Invoke-MgGraphRequest -Uri $Url -Method GET).value
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n"$_.Exception.Message
    exit (0)
}
$usersummary = @()
foreach ($result in $results) {
    $userSummary += [pscustomobject]@{                                                  ## Build array item
        Displayname       = $result.displayname
        UserPrincipalName = $result.userPrincipalName
        AccountEnabled    = $result.accountEnabled
        UserType          = $result.userType
    }
}

# Output the devices
$usersummary | sort-object displayname | Format-Table DisplayName, UserPrincipalName, AccountEnabled, UserType

if ($csv) {
    write-host -foregroundcolor $processmessagecolor "`nOutput to CSV", $outputFile
    $usersummary | export-csv $outputFile -NoTypeInformation
}

write-host -foregroundcolor $systemmessagecolor "`nGraph devices script - Finished"
if ($debug) {
    Stop-Transcript | Out-Null      
}
