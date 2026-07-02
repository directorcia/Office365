[CmdletBinding()]
param(
    [switch]$select = $false,   ## if -select parameter allows selection of individual mailboxes
    [switch]$prompt = $false,   ## if -prompt wait for user input to continue
    [string[]]$ExcludeSettings = @()
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

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

## Variables
$systemmessagecolor = 'cyan'
$processmessagecolor = 'green'
$errormessagecolor = 'red'
$scriptVerbose = $PSBoundParameters.ContainsKey('Verbose') -or $VerbosePreference -ne 'SilentlyContinue'
$scriptDebug = $PSBoundParameters.ContainsKey('Debug') -or $DebugPreference -ne 'SilentlyContinue'

function displayresult {
    param([int]$result)
    switch ($result) {
        0 { Write-Host -NoNewline -ForegroundColor $script:errormessagecolor 'X' }
        1 { Write-Host -NoNewline -ForegroundColor $script:processmessagecolor '.' }
        2 { Write-Host -NoNewline -ForegroundColor Yellow '!' }
        default { Write-Host -NoNewline -ForegroundColor Yellow '?' }
    }
}

function Get-PropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,
        [Parameter(Mandatory = $true)]
        [string[]]$PropertyNames,
        [object]$Default = $null
    )

    if ($null -eq $InputObject) {
        return $Default
    }

    foreach ($propertyName in $PropertyNames) {
        $property = $InputObject.PSObject.Properties[$propertyName]
        if ($null -ne $property) {
            return $property.Value
        }
    }

    return $Default
}

if ($args) {
    $ExcludeSettings += $args
}

Clear-Host
if ($scriptDebug) {
    if ($scriptVerbose) { Write-Host -ForegroundColor $processmessagecolor 'Create log file at ..\o365-mx-usr-all.txt' }
    Start-Transcript -Path '..\o365-mx-usr-all.txt' -ErrorAction SilentlyContinue | Out-Null
}

try {
    $exchangeModule = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Select-Object -First 1
    if (-not $exchangeModule) {
        throw 'Exchange Online PowerShell module not installed.'
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $connectionCmdlet = Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($connectionCmdlet) {
        $connectionInfo = & $connectionCmdlet.Name -ErrorAction SilentlyContinue
        if (-not $connectionInfo) {
            throw 'Not connected to Exchange Online. Run Connect-ExchangeOnline or the provided connect script first.'
        }
    }

    $mailboxCmdlet = Get-Command -Name Get-EXOMailbox -ErrorAction SilentlyContinue
    if (-not $mailboxCmdlet) { $mailboxCmdlet = Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue }
    if (-not $mailboxCmdlet) { throw 'Neither Get-EXOMailbox nor Get-Mailbox is available.' }

    $userCmdlet = Get-Command -Name Get-EXOUser -ErrorAction SilentlyContinue
    if (-not $userCmdlet) { $userCmdlet = Get-Command -Name Get-User -ErrorAction SilentlyContinue }

    $casMailboxCmdlet = Get-Command -Name Get-EXOCasMailbox -ErrorAction SilentlyContinue
    if (-not $casMailboxCmdlet) { $casMailboxCmdlet = Get-Command -Name Get-CASMailbox -ErrorAction SilentlyContinue }
    if (-not $casMailboxCmdlet) { throw 'Neither Get-EXOCasMailbox nor Get-CASMailbox is available.' }

    if ($scriptVerbose) { Write-Host -ForegroundColor $processmessagecolor 'Exchange Online PowerShell found' }
}
catch {
    Write-Host -ForegroundColor Yellow -BackgroundColor $errormessagecolor "[001] - $($_.Exception.Message) Please install and re-run the script.`n"
    if ($scriptDebug) { Stop-Transcript | Out-Null }
    exit 1
}

if ($scriptVerbose) { Write-Host -ForegroundColor $processmessagecolor 'Get best practice settings' }
try {
    $convertedOutput = Invoke-RestMethod -Method Get -Uri 'https://ciaopsgraph.azurewebsites.net/api/f9833ef6b5db63746a2322e085c39eff?id=7e92468d6e6de5183db5fde815adb06f' -ErrorAction Stop
}
catch {
    Write-Host "[002] - Unable to retrieve best-practice settings: $($_.Exception.Message)" -ForegroundColor $errormessagecolor
    $convertedOutput = $null
}

if ($scriptVerbose) { Write-Host -ForegroundColor $processmessagecolor 'Get all user mailbox information' }
try {
    $allmailboxes = & $mailboxCmdlet.Name -ResultSize Unlimited -ErrorAction Stop | Where-Object { $_.Name -NOTMATCH 'Discovery' }
}
catch {
    Write-Host "[003] - Unable to retrieve mailboxes: $($_.Exception.Message)" -ForegroundColor $errormessagecolor
    if ($scriptDebug) { Stop-Transcript | Out-Null }
    exit 2
}

if ($scriptVerbose) { Write-Host -ForegroundColor $processmessagecolor "Start user check`n" }

## remove the settings specified as arguments on command line to not be executed as part of the script
## i.e. if bc is passed, the checks for b and c are not performed
$fulllist = 'abcdefghijklmnop'
foreach ($setting in $ExcludeSettings) {
    foreach ($char in $setting.ToCharArray()) {
        $fulllist = $fulllist -replace [regex]::Escape([string]$char), ''
    }
}

if ($scriptVerbose) {
    switch -wildcard ($fulllist) {
        '*a*' { Write-Host -ForegroundColor $processmessagecolor 'a = Mailbox type: S = Shared, R = Resource, U = User' }
        '*b*' { Write-Host -ForegroundColor $processmessagecolor 'b = Enabled' }
        '*c*' { Write-Host -ForegroundColor $processmessagecolor 'c = Inactive' }
        '*d*' { Write-Host -ForegroundColor $processmessagecolor 'd = Remote PowerShell Enabled' }
        '*e*' { Write-Host -ForegroundColor $processmessagecolor 'e = Retain Deleted Items for at least 30 days' }
        '*f*' { Write-Host -ForegroundColor $processmessagecolor 'f = Deliver to Mailbox and Forward' }
        '*g*' { Write-Host -ForegroundColor $processmessagecolor 'g = Litigation Hold Enabled' }
        '*h*' { Write-Host -ForegroundColor $processmessagecolor 'h = Archive Mailbox Status' }
        '*i*' { Write-Host -ForegroundColor $processmessagecolor 'i = Auto-expanding Archive Enabled' }
        '*j*' { Write-Host -ForegroundColor $processmessagecolor 'j = Hidden From Address Lists Enabled' }
        '*k*' { Write-Host -ForegroundColor $processmessagecolor 'k = POP Enabled' }
        '*l*' { Write-Host -ForegroundColor $processmessagecolor 'l = IMAP Enabled' }
        '*m*' { Write-Host -ForegroundColor $processmessagecolor 'm = EWS Enabled' }
        '*n*' { Write-Host -ForegroundColor $processmessagecolor 'n = EWS Allow Outlook' }
        '*o*' { Write-Host -ForegroundColor $processmessagecolor 'o = EWS Allow Mac Outlook' }
        '*p*' { Write-Host -ForegroundColor $processmessagecolor 'p = Mailbox Audit Enabled' }
    }
    Write-Host
}

$fulllist

if ($select) {
    $mailboxes = $allmailboxes | Select-Object DisplayName, UserPrincipalName | Sort-Object DisplayName | Out-GridView -PassThru -Title 'Select mailboxes (Multiple selections permitted)'
}
else {
    $mailboxes = $allmailboxes
}

foreach ($mailbox in $mailboxes) {
    $mbox = $mailbox
    $mailboxIdentity = Get-PropertyValue -InputObject $mailbox -PropertyNames @('UserPrincipalName', 'PrimarySmtpAddress', 'ExternalDirectoryObjectId', 'Identity') -Default $null
    if ($null -eq $mailboxIdentity) {
        $mailboxIdentity = Get-PropertyValue -InputObject $mailbox -PropertyNames @('DisplayName') -Default '<unknown mailbox>'
    }

    $userinfo = [pscustomobject]@{
        DisplayName = (Get-PropertyValue -InputObject $mailbox -PropertyNames @('DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress') -Default '<unknown mailbox>')
        RemotePowerShellEnabled = $null
    }

    if ($userCmdlet) {
        try {
            $userinfo = & $userCmdlet.Name -Identity $mailboxIdentity -ErrorAction Stop
        }
        catch {
            if ($scriptVerbose) {
                Write-Host -ForegroundColor Yellow "User lookup skipped for ${mailboxIdentity}: $($_.Exception.Message)"
            }
        }
    }

    $extramailbox = [pscustomobject]@{
        PopEnabled = $null
        ImapEnabled = $null
        EwsEnabled = $null
        EwsAllowOutlook = $null
        EwsAllowMacOutlook = $null
    }

    try {
        $casResult = & $casMailboxCmdlet.Name -Identity $mailboxIdentity -ErrorAction Stop
        if ($null -ne $casResult) {
            $extramailbox = $casResult
        }
    }
    catch {
        if ($scriptVerbose) {
            Write-Host -ForegroundColor Yellow "CAS mailbox lookup skipped for ${mailboxIdentity}: $($_.Exception.Message)"
        }
    }

    switch -wildcard ($fulllist) {
        '*a*' {
            $recipientTypeDetails = [string](Get-PropertyValue -InputObject $mbox -PropertyNames @('RecipientTypeDetails') -Default '')
            $isShared = [bool](Get-PropertyValue -InputObject $mbox -PropertyNames @('IsShared') -Default ($recipientTypeDetails -eq 'SharedMailbox'))
            $isResource = [bool](Get-PropertyValue -InputObject $mbox -PropertyNames @('IsResource') -Default ($recipientTypeDetails -in @('RoomMailbox', 'EquipmentMailbox')))

            if ($isShared) {
                Write-Host -NoNewline -ForegroundColor $processmessagecolor -BackgroundColor DarkGreen 'S'
            }
            elseif ($isResource) {
                Write-Host -NoNewline -ForegroundColor $processmessagecolor -BackgroundColor Black 'R'
            }
            else {
                Write-Host -NoNewline -ForegroundColor $processmessagecolor 'U'
            }
        }
        '*b*' {
            $isMailboxEnabled = [bool](Get-PropertyValue -InputObject $mbox -PropertyNames @('IsMailboxEnabled') -Default $true)
            if ($isMailboxEnabled) { displayresult 1 }
            else { displayresult 0 }
        }
        '*c*' {
            $isInactiveMailbox = [bool](Get-PropertyValue -InputObject $mbox -PropertyNames @('IsInactiveMailbox') -Default $false)
            if ($isInactiveMailbox) { displayresult 0 }
            else { displayresult 1 }
        }
        '*d*' {
            if ($userinfo.RemotePowerShellEnabled -eq $false) { displayresult 1 }
            else { displayresult 0 }
        }
        '*e*' {
            $retainDeletedItemsFor = Get-PropertyValue -InputObject $mbox -PropertyNames @('RetainDeletedItemsFor') -Default $null
            if (($null -eq $retainDeletedItemsFor) -or ($null -eq $convertedOutput)) {
                displayresult 0
            }
            elseif ([timespan]::Parse([string]$retainDeletedItemsFor).Days -ge [int]$convertedOutput.retaindeleteditemsfor) {
                displayresult 1
            }
            else {
                displayresult 0
            }
        }
        '*f*' {
            $deliverToMailboxAndForward = Get-PropertyValue -InputObject $mbox -PropertyNames @('DeliverToMailboxAndForward') -Default $null
            if (($null -ne $deliverToMailboxAndForward) -and ($null -ne $convertedOutput) -and ($deliverToMailboxAndForward -eq $convertedOutput.delivertomailboxandforward)) { displayresult 1 }
            else { displayresult 0 }
        }
        '*g*' {
            $litigationHoldEnabled = Get-PropertyValue -InputObject $mbox -PropertyNames @('LitigationHoldEnabled') -Default $null
            if (($null -ne $litigationHoldEnabled) -and ($null -ne $convertedOutput) -and ($litigationHoldEnabled -eq $convertedOutput.LitigationHoldEnabled)) { displayresult 1 }
            else { displayresult 0 }
        }
        '*h*' {
            $archiveStatus = Get-PropertyValue -InputObject $mbox -PropertyNames @('ArchiveStatus') -Default $null
            if (($null -ne $archiveStatus) -and ($null -ne $convertedOutput) -and ($archiveStatus -eq $convertedOutput.ArchiveStatus)) { displayresult 1 }
            else { displayresult 0 }
        }
        '*i*' {
            $autoExpandingArchiveEnabled = Get-PropertyValue -InputObject $mbox -PropertyNames @('AutoExpandingArchiveEnabled') -Default $null
            if (($null -ne $autoExpandingArchiveEnabled) -and ($null -ne $convertedOutput) -and ($autoExpandingArchiveEnabled -eq $convertedOutput.AutoExpandingArchive)) { displayresult 1 }
            else { displayresult 0 }
        }
        '*j*' {
            $hiddenFromAddressListsEnabled = Get-PropertyValue -InputObject $mbox -PropertyNames @('HiddenFromAddressListsEnabled') -Default $null
            if (($null -ne $hiddenFromAddressListsEnabled) -and ($null -ne $convertedOutput) -and ($hiddenFromAddressListsEnabled -eq $convertedOutput.HiddenFromAddressListsEnabled)) { displayresult 1 }
            else { displayresult 0 }
        }
        '*k*' {
            $popEnabled = Get-PropertyValue -InputObject $extramailbox -PropertyNames @('PopEnabled') -Default $null
            if (($null -ne $popEnabled) -and ($null -ne $convertedOutput) -and ($popEnabled -eq $convertedOutput.PopEnabled)) { displayresult 1 }
            else { displayresult 0 }
        }
        '*l*' {
            $imapEnabled = Get-PropertyValue -InputObject $extramailbox -PropertyNames @('ImapEnabled') -Default $null
            if (($null -ne $imapEnabled) -and ($null -ne $convertedOutput) -and ($imapEnabled -eq $convertedOutput.ImapEnabled)) { displayresult 1 }
            else { displayresult 0 }
        }
        '*m*' {
            $ewsEnabled = Get-PropertyValue -InputObject $extramailbox -PropertyNames @('EwsEnabled') -Default $null
            if (($null -ne $ewsEnabled) -and ($null -ne $convertedOutput) -and ($ewsEnabled -eq $convertedOutput.EwsEnabled)) { displayresult 1 }
            else { displayresult 3 }
        }
        '*n*' {
            $ewsAllowOutlook = Get-PropertyValue -InputObject $extramailbox -PropertyNames @('EwsAllowOutlook') -Default $null
            if (($ewsAllowOutlook -eq $true) -or ($ewsAllowOutlook -eq $null)) { displayresult 1 }
            else { displayresult 0 }
        }
        '*o*' {
            $ewsAllowMacOutlook = Get-PropertyValue -InputObject $extramailbox -PropertyNames @('EwsAllowMacOutlook') -Default $null
            if (($ewsAllowMacOutlook -eq $true) -or ($ewsAllowMacOutlook -eq $null)) { displayresult 1 }
            else { displayresult 0 }
        }
        '*p*' {
            $auditEnabled = Get-PropertyValue -InputObject $mbox -PropertyNames @('AuditEnabled') -Default $null
            if (($null -ne $auditEnabled) -and ($null -ne $convertedOutput) -and ($auditEnabled -eq $convertedOutput.AuditEnabled)) { displayresult 1 }
            else { displayresult 0 }
        }
    }

    Write-Host -NoNewline ':'
    if ($prompt) {
        Write-Host -NoNewline $userinfo.DisplayName
        Read-Host
    }
    else {
        Write-Host $userinfo.DisplayName
    }
}

$fulllist
if ($scriptVerbose) { Write-Host -ForegroundColor $systemmessagecolor "`nGet user information from Exchange Online script finished" }
if ($scriptDebug) {
    Stop-Transcript | Out-Null
}