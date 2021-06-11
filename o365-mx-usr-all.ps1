param(                        
    [switch]$debug = $false,    ## if -debug parameter create activity log file
    [switch]$select = $false,   ## if -user parameter allow selection of individual mailboxes
    [switch]$prompt = $false,   ## if -prompt wait for user input to continue
    [switch]$verbose = $false   ## if -verbose paramter add more information
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Display audit details for all Exchange users
Documentation - https://blog.ciaops.com/2021/06/09/exchange-user-best-practices-script/
Source - https://github.com/directorcia/office365/blob/master/o365-mx-usr-all.ps1

Prerequisites = 1
1. Ensure connected to Exchange Online. Use the script https://github.com/directorcia/Office365/blob/master/o365-connect-exo.ps1

for more scripts visit www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"

function displayresult($result){
    switch ($result) {
        0 {write-host -nonewline -foregroundcolor $errormessagecolor "X"}
        1 {write-host -nonewline -foregroundcolor $processmessagecolor "."}
    }
}

Clear-Host
if ($debug) {
    if ($verbose) { write-host -ForegroundColor $processmessagecolor "Create log file at ..\o365-mx-usr-all.txt"}
    start-transcript "..\o365-mx-usr-all.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

if (get-module -listavailable -name ExchangeOnlineManagement) {    ## Has the Exchange Online PowerShell module been loaded?
    if ($verbose) {write-host -ForegroundColor $processmessagecolor "Exchange Online PowerShell found"}
}
else {
    write-host -ForegroundColor yellow -backgroundcolor $errormessagecolor "[001] - Exchange Online PowerShell module not installed. Please install and re-run script`n"
    if ($debug) {
        Stop-Transcript                 ## Terminate transcription
    }
    exit 1                          ## Terminate script
}
if ($verbose) {write-host -ForegroundColor $processmessagecolor "Get best practice settings"}
try {
    $query = invoke-webrequest -method GET -ContentType "application/json" -uri https://ciaopsgraph.azurewebsites.net/api/f9833ef6b5db63746a2322e085c39eff?id=7e92468d6e6de5183db5fde815adb06f -UseBasicParsing
}
catch {
    Write-Host "[002]", $_.Exception.Message -ForegroundColor $errormessagecolor                       
}
$convertedOutput = $query.content | ConvertFrom-Json

if ($verbose) {write-host -ForegroundColor $processmessagecolor "Get all user mailbox information"}
$allmailboxes = get-mailbox -ResultSize unlimited | Where-Object {$_.name -NOTMATCH "Discovery"}
if ($verbose) {write-host -ForegroundColor $processmessagecolor "Start user check`n"}

## remove the settings specified as arguments on command line to not be executed as part of the script
## i.e if bc and arguments on command line retaindeleteditemsfor and delivertomailboxandforward are not excluded
$fulllist ="abcdefghijklmnop"   # Define the full list of settings to check
$list = $args[0]                # Get arguments from command line. These will be settings NOT to be displayed and checked
for ($i=0; $i -lt $list.length; $i++) {
    $fulllist = $fulllist -replace ($list[$i],'')           # If a setting is found as a parameter remove from the full list of paramters
}
if ($verbose) {
    switch -wildcard ($fulllist) {
        "*a*" { write-host -ForegroundColor $processmessagecolor "a = Mailbox type: S = Shared, R = Resource, U = User"}
        "*b*" { write-host -ForegroundColor $processmessagecolor "b = Enabled"}
        "*c*" { write-host -ForegroundColor $processmessagecolor "c = Inactive"}
        "*d*" { write-host -ForegroundColor $processmessagecolor "d = Remote PowerShell Enabled"}
        "*e*" { write-host -ForegroundColor $processmessagecolor "e = Retain Deleted Items for at least 30 days"}
        "*f*" { write-host -ForegroundColor $processmessagecolor "f = Deliver to Mailbox and Forward"}
        "*g*" { write-host -ForegroundColor $processmessagecolor "g = Litigation Hold Enabled"}
        "*h*" { write-host -ForegroundColor $processmessagecolor "h = Archive Mailbox Status"}
        "*i*" { write-host -ForegroundColor $processmessagecolor "i = Auto-expanding Archive Enabled"}
        "*j*" { write-host -ForegroundColor $processmessagecolor "j = Hidden From Address Lists Enabled"}
        "*k*" { write-host -ForegroundColor $processmessagecolor "k = POP Enabled"}
        "*l*" { write-host -ForegroundColor $processmessagecolor "l = IMAP Enabled"}
        "*m*" { write-host -ForegroundColor $processmessagecolor "m = EWS Enabled"}
        "*n*" { write-host -ForegroundColor $processmessagecolor "n = EWS Allow Outlook"}
        "*o*" { write-host -ForegroundColor $processmessagecolor "o = EWS Allow Mac Outlook"}
        "*p*" { write-host -ForegroundColor $processmessagecolor "p = Mailbox Audit Enabled"}
    }
    write-host
}
$fulllist       # Display the updated list of paramters to be displayed and checked
if ($select) {  # If want to individually select users to be checked then prompt
    $mailboxes = $allmailboxes | select-object displayname,userprincipalname | Sort-Object displayname | Out-GridView -PassThru -title "Select mailboxes (Multiple selections permitted) "    
}
else {
    $mailboxes = $allmailboxes
}
foreach ($mailbox in $mailboxes) {  # loop through all mailboxes selected
    $mbox = get-mailbox -identity $mailbox.userprincipalname
    $userinfo = get-user -identity $mailbox.userprincipalname
    $extramailbox = get-casmailbox -identity $mailbox.userprincipalname
    switch -wildcard ($fulllist) {
        "*a*" {
            if ($mbox.isshared) {
                write-host -nonewline -ForegroundColor $processmessagecolor -BackgroundColor DarkGreen "S"
            }
            else {
                if ($mbox.isresource) {
                    write-host -nonewline -ForegroundColor $processmessagecolor -BackgroundColor Black "R"
                }
                else {
                    write-host -nonewline -ForegroundColor $processmessagecolor "U"
                }
            }
        }
        "*b*" {
            if ($mbox.ismailboxenabled) { displayresult(1)}
            else { displayresult(0) }
        }
        "*c*" {
            if ($mbox.isinactivemailbox) { displayresult(0) }
            else { displayresult(1) }
        }
        "*d*" {
            if ($userinfo.remotepowershellenabled -eq $false) { displayresult(1) }
            else { displayresult(0) }    
        }
        "*e*" {
            if ([timespan]::parse($mbox.retaindeleteditemsfor).days -ge $convertedoutput.retaindeleteditemsfor) { displayresult(1) }
            else { displayresult(0) }
        }
        "*f*" {
            if ($mbox.delivertomailboxandforward -eq $convertedoutput.delivertomailboxandforward) { displayresult(1) }
            else { displayresult(0) }
        }
        "*g*" {
            if ($mbox.LitigationHoldEnabled -eq $convertedoutput.LitigationHoldEnabled) { displayresult(1) }
            else { displayresult(0) }
        }
        "*h*" {
            if ($mbox.archivestatus -eq $convertedoutput.archivestatus) { displayresult(1) }
            else { displayresult(0) }
        }
        "*i*" {
            if ($mbox.autoexpandingarchiveenabled -eq $convertedoutput.AutoExpandingArchive) { displayresult(1) }
            else { displayresult(0) }
        }
        "*j*" {
            if ($mbox.HiddenFromAddressListsEnabled -eq $convertedoutput.HiddenFromAddressListsEnabled) { displayresult(1) }
            else { displayresult(0) }
        }
        "*k*" {
            if ($extramailbox.popenabled -eq $convertedoutput.popenabled) { displayresult(1) }
            else { displayresult(0) }
        }
        "*l*" {
            if ($extramailbox.imapenabled -eq $convertedoutput.imapenabled) { displayresult(1) }
            else { displayresult(0) }
        }
        "*m*" {
            if ($extramailbox.ewsenabled -eq $convertedoutput.ewsenabled) { displayresult(1) }
            else { displayresult(3) }                    
        }
        "*n*" {
            if (($extramailbox.ewsallowoutlook -eq $true) -or ($extramailbox.ewsallowoutlook -eq $null)) { displayresult(1) } 
            else { displayresult(0) }
        }
        "*o*" {
            if (($extramailbox.ewsallowmacoutlook -eq $true) -or ($extramailbox.ewsallowmacoutlook -eq $null)) { displayresult(1) }
            else { displayresult(0) }                    
        }
        "*p*" {
            if ($mbox.auditenabled -eq $convertedoutput.auditenabled) { displayresult(1) }
            else { displayresult(0) }                    
        }
    }
    write-host -nonewline ":"
    if ($prompt) {
        write-host -nonewline $userinfo.displayname
        read-host       # command unfortunately leaves CR/LF which spreads out display
    }
    else {
        write-host $userinfo.displayname
    }
}
$fulllist
if ($verbose) {Write-Host -ForegroundColor $systemmessagecolor "`nGet user information from Exchange Online script finished"}
if ($debug) {
    stop-transcript | Out-Null
}