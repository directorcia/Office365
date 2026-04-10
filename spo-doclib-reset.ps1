param(                          ## if no parameter used then login without MFA and use interactive mode
    [switch]$debug = $false,    ## if -debug parameter capture log information
    [switch]$prompt = $false  ## if -noprompt parameter used don't prompt user for input
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Reset permissions in a SharePoint document library to all be inherit
Source - https://github.com/directorcia/patron/blob/master/spo-doclib-reset.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Resets-the-permissions-on-every-item-in-a-SharePoint-document-library

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
$selectedList = $lists | select-object Title, Id | Sort-Object Title | Out-GridView -OutputMode Single -title "Select List"
if ($null -eq $selectedList) {
    write-host -ForegroundColor $warningmessagecolor "No list selected. Exiting."
    if ($debug) { Stop-Transcript | Out-Null }
    exit 0
}

write-host -ForegroundColor $processmessagecolor "[Launch] = Read items (this may take some time depending on the number of items in the list)"
$items=(Get-PnPListItem -List $selectedList.id -pagesize 5000 -Fields "Title","GUID","FileRef").FieldValues
write-host -ForegroundColor $processmessagecolor "[Finish] = Read items"
Write-Host -ForegroundColor $processmessagecolor "Total Number of List Items: $($items.Count)`n"

write-host -ForegroundColor $processmessagecolor "Reset all permissions on all items to inherit`n"
if ($prompt) {
    do {
        $confirm = read-host -Prompt "`nAre you sure [Y/N]"
    } until (-not [string]::isnullorempty($confirm))
    if ($confirm -eq 'N' -or $confirm -eq 'n') {
        if ($debug) { Stop-Transcript | Out-Null }
        exit 2
    }
}
$count=0
$totalitems = $items.Count
foreach ($item in $items) {
    ++$count
    write-host -nonewline "[$count of $totalitems] Id =",$item.ID,"FileRef =",$item.FileRef
    try {
        Set-PnPListItemPermission -List $selectedList.id -Identity $item.ID -InheritPermissions
        write-host -foregroundcolor $processmessagecolor " - Success"
    }
    catch {
        write-host -foregroundcolor $errormessagecolor " - Failed"
        write-host -foregroundcolor $errormessagecolor "`n", $_.Exception.Message
    }
}

write-host -foregroundcolor $systemmessagecolor "`nReset Document Library permissions - complete`n"
if ($debug) {
    Stop-Transcript | Out-Null
}