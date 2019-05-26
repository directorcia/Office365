<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description
Script designed to check and report the status of mailbox in the tenant

Source - https://github.com/directorcia/Office365/blob/master/o365-mx-check.ps1

Prerequisites = 1
1. Connected to Exchange Online

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$auditlogagelimitdefault = 90
$retaindeleteditemsfordefault = 14
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

write-host -ForegroundColor $processmessagecolor "Getting Mailboxes"
$mailboxes=get-mailbox -ResultSize unlimited
write-host -ForegroundColor $processmessagecolor "Start checking mailboxes"
write-host
foreach ($mailbox in $mailboxes){
    write-host -foregroundcolor yellow -BackgroundColor Black "Mailbox =",$mailbox.displayname

## all mailboxes should have auditing enabled

    if ($mailbox.auditenabled) {
        write-host -foregroundcolor $processmessagecolor "  Audit enabled =",$mailbox.AuditEnabled
    } else {
        write-host -foregroundcolor $errormessagecolor "  Audit enabled =",$mailbox.AuditEnabled
    }

## audit log limit for mailboxes shoudl be extended from default

    if ([timespan]::parse($mailbox.auditlogagelimit).days -gt $auditlogagelimitdefault) {
        write-host -foregroundcolor $processmessagecolor "  Audit login limit (days)",$mailbox.Auditlogagelimit
    } else {
        write-host -foregroundcolor $errormessagecolor "  Audit login limit (days)",$mailbox.Auditlogagelimit
    }

## all mailboxes should have their retained deleted item retention period extended to 30 days

    if ([timespan]::parse($mailbox.retaindeleteditemsfor).days -gt $retaindeleteditemsfordefault) {
        write-host -foregroundcolor $processmessagecolor "  Retain Deleted items for (days) =",$mailbox.retaindeleteditemsfor
    } else {
        write-host -foregroundcolor $errormessagecolor "  Retain Deleted items for (days) =",$mailbox.retaindeleteditemsfor
    }

## mailboxes should not be forwarding to other email addresses

    if ($mailbox.forwardingaddress -ne $null){
        write-host -foregroundcolor $errormessagecolor "  Forwarding address =",$mailbox.forwardingaddress
    }
    if ($mailbox.forwardingsmtpaddress -ne $null){
        write-host -foregroundcolor $errormessagecolor "  Forwarding SMTP address =",$mailbox.forwardingsmtpaddress
    }

## mailboxes should have litigation hold enabled

    if ($mailbox.LitigationHoldEnabled) {
        write-host -foregroundcolor $processmessagecolor "  Litigation hold =",$mailbox.litigationholdenabled
    } else {
        write-host -foregroundcolor $errormessagecolor "  Litigation hold =",$mailbox.litigationholdenabled
    }

## mailboxes should have archive enabled

    if ($mailbox.archivestatus -eq "active") {
        Write-host -foregroundcolor $processmessagecolor "  Archive status =",$mailbox.archivestatus
    } else {
        Write-host -foregroundcolor $errormessagecolor "  Archive status =",$mailbox.archivestatus
    }

## report mailbox maximum send and receive sizes

    Write-host "  Max send size =",$mailbox.maxsendsize
    write-host "  Max receive size =",$mailbox.maxreceivesize
    $extramailbox=get-casmailbox -Identity $mailbox.displayname

## mailboxes should not have POP3 enabled

    if (-not $extramailbox.popenabled) {
        write-host -foregroundcolor $processmessagecolor "  POP3 enabled =",$extramailbox.popenabled
    } else {
        write-host -foregroundcolor $errormessagecolor "  POP3 enabled =",$extramailbox.popenabled
    }

## mailboxes should not have IMAP enabled

    if (-not $extramailbox.ImapEnabled) {
        write-host -foregroundcolor $processmessagecolor "  IMAP enabled =",$extramailbox.imapenabled
    } else {
        write-host -foregroundcolor $errormessagecolor "  IMAP enabled =",$extramailbox.imapenabled
    }
    write-host
}
write-host -ForegroundColor $processmessagecolor "Finish checking mailboxes"
write-host
write-host -foregroundcolor $systemmessagecolor "Script completed`n"
