param(
    [switch]$csv = $false
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description
Script designed to check and report the status of mailboxes in the tenant

Source - https://github.com/directorcia/Office365/blob/master/o365-mx-check.ps1

Notes - You can extend many of the default limits at no additional cost

Prerequisites = 1
1. Connected to Exchange Online. Recommended script = https://github.com/directorcia/Office365/blob/master/o365-connect-exo.ps1

More scripts available by joining http://www.ciaopspatron.com

#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

## Variables
$auditlogagelimitdefault = 90
$retaindeleteditemsmax = 30
$systemmessagecolor = 'Cyan'
$processmessagecolor = 'Green'
$errormessagecolor = 'Red'
$version = '2.10'
$transcriptStarted = $false

function Get-PreferredCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Names
    )

    foreach ($name in $Names) {
        $command = Get-Command -Name $name -ErrorAction SilentlyContinue
        if ($command) {
            return $command
        }
    }

    return $null
}

function Get-SafeDaysValue {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    try {
        return [TimeSpan]::Parse([string]$Value).Days
    }
    catch {
        return $null
    }
}

$transcriptPath = '..\o365-mx-check.txt'
try {
    Start-Transcript -Path $transcriptPath -ErrorAction Stop | Out-Null
    $transcriptStarted = $true
}
catch {
    Write-Host -ForegroundColor Yellow "Unable to start transcript at ${transcriptPath}: $($_.Exception.Message)"
}

Clear-Host

Write-Host -ForegroundColor $systemmessagecolor "Script started. Version = $version`n"
Write-Host -ForegroundColor Cyan -BackgroundColor DarkBlue ">>>>>> Created by www.ciaops.com <<<<<<`n"
Write-Host "--- Script to display mailbox settings ---`n"

try {
    $exchangeModule = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Select-Object -First 1
    if (-not $exchangeModule) {
        throw 'ExchangeOnlineManagement module not installed. Please install and re-run script.'
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $getMailboxCommand = Get-PreferredCommand -Names @('Get-EXOMailbox', 'Get-Mailbox')
    if (-not $getMailboxCommand) {
        throw 'Neither Get-EXOMailbox nor Get-Mailbox is available in this session.'
    }

    $getCasMailboxCommand = Get-PreferredCommand -Names @('Get-EXOCasMailbox', 'Get-CASMailbox')
    if (-not $getCasMailboxCommand) {
        throw 'Neither Get-EXOCasMailbox nor Get-CASMailbox is available in this session.'
    }

    $connectionCmdlet = Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($connectionCmdlet) {
        $connectionInfo = & $connectionCmdlet.Name -ErrorAction SilentlyContinue
        if (-not $connectionInfo) {
            throw 'Not connected to Exchange Online. Run Connect-ExchangeOnline first.'
        }
    }

    Write-Host -ForegroundColor $processmessagecolor 'Exchange Online PowerShell found'
}
catch {
    Write-Host -ForegroundColor Yellow -BackgroundColor Red "`n[001] - $($_.Exception.Message)`n"
    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }
    exit 1
}

Write-Host -ForegroundColor $processmessagecolor 'Getting Mailboxes'

try {
    if ($getMailboxCommand.Name -eq 'Get-EXOMailbox') {
        $mailboxes = & $getMailboxCommand.Name -ResultSize Unlimited -Properties AuditEnabled, AuditLogAgeLimit, RetainDeletedItemsFor, ForwardingAddress, ForwardingSmtpAddress, LitigationHoldEnabled, ArchiveStatus, MaxSendSize, MaxReceiveSize, UserPrincipalName, PrimarySmtpAddress, DisplayName, WhenCreated -ErrorAction Stop
    }
    else {
        $mailboxes = & $getMailboxCommand.Name -ResultSize Unlimited -ErrorAction Stop
    }
}
catch {
    Write-Host -ForegroundColor Yellow -BackgroundColor Red "`n[002] - Unable to retrieve mailboxes. $($_.Exception.Message)`n"
    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }
    exit 2
}

if (-not $mailboxes) {
    Write-Host -ForegroundColor Yellow 'No mailboxes returned from Exchange Online.'
    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }
    exit 0
}

Write-Host -ForegroundColor $processmessagecolor "Start checking mailboxes`n"

$results = New-Object System.Collections.Generic.List[object]
$compliantMailboxCount = 0

foreach ($mailbox in $mailboxes) {
    $mailboxCompliant = $true
    $upn = [string]$mailbox.UserPrincipalName
    $primarySmtp = [string]$mailbox.PrimarySmtpAddress

    if ($upn.Length -gt 60) {
        $upnDisplay = $upn.Substring(0, 60) + '...'
    }
    else {
        $upnDisplay = $upn
    }

    if ($primarySmtp.Length -gt 60) {
        $primarySmtpDisplay = $primarySmtp.Substring(0, 60) + '...'
    }
    else {
        $primarySmtpDisplay = $primarySmtp
    }

    Write-Host -ForegroundColor Yellow -BackgroundColor Black "Mailbox = $($mailbox.DisplayName) [$upnDisplay]"
    Write-Host -ForegroundColor Gray "  Primary SMTP address = $primarySmtpDisplay"
    Write-Host -ForegroundColor Gray "  Created = $($mailbox.WhenCreated)"

    if ($mailbox.AuditEnabled) {
        Write-Host -ForegroundColor $processmessagecolor "  Audit enabled = $($mailbox.AuditEnabled)"
    }
    else {
        $mailboxCompliant = $false
        Write-Host -ForegroundColor $errormessagecolor "  Audit enabled = $($mailbox.AuditEnabled)"
    }

    $auditLogAgeLimitDays = Get-SafeDaysValue -Value $mailbox.AuditLogAgeLimit
    if (($null -ne $auditLogAgeLimitDays) -and ($auditLogAgeLimitDays -gt $auditlogagelimitdefault)) {
        Write-Host -ForegroundColor $processmessagecolor "  Audit log limit (days) = $auditLogAgeLimitDays"
    }
    else {
        $mailboxCompliant = $false
        Write-Host -ForegroundColor $errormessagecolor "  Audit log limit (days) = $auditLogAgeLimitDays"
    }

    $retainDeletedItemsDays = Get-SafeDaysValue -Value $mailbox.RetainDeletedItemsFor
    if (($null -ne $retainDeletedItemsDays) -and ($retainDeletedItemsDays -ge $retaindeleteditemsmax)) {
        Write-Host -ForegroundColor $processmessagecolor "  Retain Deleted items for (days) = $retainDeletedItemsDays"
    }
    else {
        $mailboxCompliant = $false
        Write-Host -ForegroundColor $errormessagecolor "  Retain Deleted items for (days) = $retainDeletedItemsDays"
    }

    if (-not [string]::IsNullOrEmpty([string]$mailbox.ForwardingAddress)) {
        $mailboxCompliant = $false
        Write-Host -ForegroundColor $errormessagecolor "  Forwarding address = $($mailbox.ForwardingAddress)"
    }

    if (-not [string]::IsNullOrEmpty([string]$mailbox.ForwardingSmtpAddress)) {
        $mailboxCompliant = $false
        Write-Host -ForegroundColor $errormessagecolor "  Forwarding SMTP address = $($mailbox.ForwardingSmtpAddress)"
    }

    if ($mailbox.LitigationHoldEnabled) {
        Write-Host -ForegroundColor $processmessagecolor "  Litigation hold = $($mailbox.LitigationHoldEnabled)"
    }
    else {
        $mailboxCompliant = $false
        Write-Host -ForegroundColor $errormessagecolor "  Litigation hold = $($mailbox.LitigationHoldEnabled)"
    }

    if ([string]$mailbox.ArchiveStatus -eq 'Active') {
        Write-Host -ForegroundColor $processmessagecolor "  Archive status = $($mailbox.ArchiveStatus)"
    }
    else {
        $mailboxCompliant = $false
        Write-Host -ForegroundColor $errormessagecolor "  Archive status = $($mailbox.ArchiveStatus)"
    }

    Write-Host -ForegroundColor Gray "  Max send size = $($mailbox.MaxSendSize)"
    Write-Host -ForegroundColor Gray "  Max receive size = $($mailbox.MaxReceiveSize)"

    $extraMailbox = $null
    $casLookupReason = $null
    try {
        $extraMailbox = & $getCasMailboxCommand.Name -Identity $mailbox.UserPrincipalName -ErrorAction Stop
    }
    catch {
        $casLookupReason = $_.Exception.Message
    }

    if ($null -ne $extraMailbox) {
        if (-not $extraMailbox.PopEnabled) {
            Write-Host -ForegroundColor $processmessagecolor "  POP3 enabled = $($extraMailbox.PopEnabled)"
        }
        else {
            $mailboxCompliant = $false
            Write-Host -ForegroundColor $errormessagecolor "  POP3 enabled = $($extraMailbox.PopEnabled)"
        }

        if (-not $extraMailbox.ImapEnabled) {
            Write-Host -ForegroundColor $processmessagecolor "  IMAP enabled = $($extraMailbox.ImapEnabled)"
        }
        else {
            $mailboxCompliant = $false
            Write-Host -ForegroundColor $errormessagecolor "  IMAP enabled = $($extraMailbox.ImapEnabled)"
        }
    }
    else {
        $mailboxCompliant = $false
        if ([string]::IsNullOrWhiteSpace($casLookupReason)) {
            $casLookupReason = 'Unknown error'
        }
        $trimmedReason = if ($casLookupReason.Length -gt 100) { $casLookupReason.Substring(0, 100) + '...' } else { $casLookupReason }
        Write-Host -ForegroundColor $errormessagecolor "  Further mailbox details not found - $trimmedReason"
    }

    if ($mailboxCompliant) {
        $compliantMailboxCount++
    }

    $results.Add([PSCustomObject]@{
            DisplayName = $mailbox.DisplayName
            UserPrincipalName = $mailbox.UserPrincipalName
            PrimarySMTPAddress = $mailbox.PrimarySmtpAddress
            WhenCreated = $mailbox.WhenCreated
            AuditEnabled = $mailbox.AuditEnabled
            AuditLogAgeLimit = $mailbox.AuditLogAgeLimit
            AuditLogAgeLimitDays = $auditLogAgeLimitDays
            RetainDeletedItemsFor = $mailbox.RetainDeletedItemsFor
            RetainDeletedItemsDays = $retainDeletedItemsDays
            ForwardingAddress = $mailbox.ForwardingAddress
            ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
            LitigationHoldEnabled = $mailbox.LitigationHoldEnabled
            ArchiveStatus = $mailbox.ArchiveStatus
            PopEnabled = if ($null -ne $extraMailbox) { $extraMailbox.PopEnabled } else { $null }
            ImapEnabled = if ($null -ne $extraMailbox) { $extraMailbox.ImapEnabled } else { $null }
            MaxSendSize = $mailbox.MaxSendSize
            MaxReceiveSize = $mailbox.MaxReceiveSize
            MailboxCompliant = $mailboxCompliant
        })

    Write-Host
}

$totalMailboxes = $results.Count
$nonCompliantMailboxCount = $totalMailboxes - $compliantMailboxCount

Write-Host -ForegroundColor $processmessagecolor "Summary: $compliantMailboxCount compliant, $nonCompliantMailboxCount non-compliant, $totalMailboxes total"

if ($csv) {
    $csvPath = "..\o365-mx-check$(Get-Date -Format yyyyMMddHHmmss).csv"
    Write-Host -ForegroundColor $processmessagecolor "Writing all output to file $csvPath in parent directory"
    $results | Export-Csv -Path $csvPath -NoTypeInformation
}
else {
    Write-Host -ForegroundColor $processmessagecolor 'No CSV created'
}

Write-Host -ForegroundColor $processmessagecolor 'Finish checking mailboxes'
Write-Host
Write-Host -ForegroundColor $systemmessagecolor "Script completed`n"

if ($transcriptStarted) {
    Stop-Transcript | Out-Null
}