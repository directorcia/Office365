param(                          ## if no parameter used then login without MFA and use interactive mode
    [switch]$debug = $false,    ## if -debug parameter capture log information
    [switch]$prompt = $false  ## if -noprompt parameter used don't prompt user for input
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Reset permissions in a SharePoint document library to all be inherit
Source - https://github.com/directorcia/patron/blob/master/spo-doclib-reset.txt

Prerequisites = 2
1. Ensure connected to SharePoint PnP  - Use the script https://github.com/directorcia/Office365/blob/master/o365-connect-pnp.ps1
2. The latest version of PnP PowerShell module require PowerShell version 7. 

More scripts are available via www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

if ($debug) {
    start-transcript "..\spo-doclib-reset.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

write-host -foregroundcolor $systemmessagecolor "Reset Document Library permissions - start`n"
write-host -ForegroundColor $processmessagecolor "Debug =", $debug
write-host -ForegroundColor $processmessagecolor "Prompt =", ($prompt)
if ($debug) {
        write-host "Script activity logged at ..\spo-doclib-reset.txt"
}

write-host -ForegroundColor $processmessagecolor "Get Site Lists"
try {
    $lists = Get-PnPList
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "`n[001] - Failed to connect via PnP ", $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null
    }
    exit 1
}
write-host -ForegroundColor $processmessagecolor "Select List"
$result = $lists | select-object Title, Id | Sort-Object Title | Out-GridView -OutputMode Single -title "Select List"

write-host -ForegroundColor $processmessagecolor "Read items"
$items=(Get-PnPListItem -List $result.id -Fields "Title","GUID").FieldValues

write-host -ForegroundColor $processmessagecolor "Reset all permissions on all items to inherit`n"
if ($prompt) {
    do {
        $result=read-host -Prompt "`nAre you sure [Y/N]"
    } until (-not [string]::isnullorempty($result))
    if ($result -eq 'N' -or $result -eq 'n') {
        Stop-Transcript | Out-Null
        exit 2
    }
}

foreach ($item in $items) {
    write-host "Id =",$item.ID,"FileRef =",$item.FileRef
    Set-PnPListItemPermission -List $result.id -Identity $item.ID -InheritPermissions
}

write-host -foregroundcolor $systemmessagecolor "`nReset Document Library permissions - complete`n"
if ($debug) {
    Stop-Transcript | Out-Null
}