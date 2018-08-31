## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to check and report the status of mailbox in the tenant

## Source - https://github.com/directorcia/Office365/blob/master/o365-mx-check.ps1

## Prerequisites = 1
## 1. Connected to Exchange Online

## Variables
$auditlogagelimitdefault = 90
$retaindeleteditemsfordefault = 14
$systemmessagecolor = "cyan"

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started"

## set-executionpolicy remotesigned
## May be required once to allow ability to runs scripts in PowerShell

write-host -ForegroundColor $systemmessagecolor "Getting Mailboxes"
$mailboxes=get-mailbox -ResultSize unlimited
write-host -ForegroundColor $systemmessagecolor "Start checking mailboxes"
write-host
foreach ($mailbox in $mailboxes){
    write-host -foregroundcolor yellow -BackgroundColor Black "Mailbox =",$mailbox.displayname

## all mailboxes should have auditing enabled

    if ($mailbox.auditenabled) {
        write-host -foregroundcolor green "  Audit enabled =",$mailbox.AuditEnabled
    } else {
        write-host -foregroundcolor red "  Audit enabled =",$mailbox.AuditEnabled
    }

## audit log limit for mailboxes shoudl be extended from default

    if ([timespan]::parse($mailbox.auditlogagelimit) -gt $auditlogagelimitdefault) {
        write-host -foregroundcolor green "  Audit login limit (days)",$mailbox.Auditlogagelimit
    } else {
        write-host -foregroundcolor red "  Audit login limit (days)",$mailbox.Auditlogagelimit
    }

## all mailboxes should have their retained deleted item retention period extended to 30 days

    if ([timespan]::parse($mailbox.retaindeleteditemsfor) -gt $retaindeleteditemsfordefault) {
        write-host -foregroundcolor green "  Retain Deleted items for (days) =",$mailbox.retaindeleteditemsfor
    } else {
        write-host -foregroundcolor red "  Retain Deleted items for (days) =",$mailbox.retaindeleteditemsfor
    }

## mailboxes should not be forwarding to oteh email addresses

    if ($mailbox.forwardingaddress -ne $null){
        write-host -foregroundcolor red "  Forwarding address =",$mailbox.forwardingaddress
    }
    if ($mailbox.forwardingsmtpaddress -ne $null){
        write-host -foregroundcolor red "  Forwarding SMTP address =",$mailbox.forwardingsmtpaddress
    }

## mailboxes should have litigation hold enabled

    if ($mailbox.LitigationHoldEnabled) {
        write-host -foregroundcolor green "  Litigation hold =",$mailbox.litigationholdenabled
    } else {
        write-host -foregroundcolor red "  Litigation hold =",$mailbox.litigationholdenabled
    }

## mailboxes should have archive enabled

    if ($mailbox.archivestatus -eq "active") {
        Write-host -foregroundcolor Green "  Archive status =",$mailbox.archivestatus
    } else {
        Write-host -foregroundcolor red "  Archive status =",$mailbox.archivestatus
    }

## report mailbox maximum send and receive sizes

    Write-host "  Max send size =",$mailbox.maxsendsize
    write-host "  Max receive size =",$mailbox.maxreceivesize
    $extramailbox=get-casmailbox -Identity $mailbox.displayname

## mailboxes should not have POP3 enabled

    if (-not $extramailbox.popenabled) {
        write-host -foregroundcolor green "  POP3 enabled =",$extramailbox.popenabled
    } else {
        write-host -foregroundcolor red "  POP3 enabled =",$extramailbox.popenabled
    }

## mailboxes should not have IMAP enabled

    if (-not $extramailbox.ImapEnabled) {
        write-host -foregroundcolor green "  IMAP enabled =",$extramailbox.imapenabled
    } else {
        write-host -foregroundcolor red "  IMAP enabled =",$extramailbox.imapenabled
    }
    write-host
}
write-host -ForegroundColor $systemmessagecolor "Finish checking mailboxes"
write-host
write-host -ForegroundColor $systemmessagecolor "Finish script"
