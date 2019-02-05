## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to check which email boxes have forwarding options set

## Source - https://github.com/directorcia/Office365/blob/master/o365-exo-fwd-chk.ps1

## Prerequisites = 1
## 1. Ensure connection to Exchange Online has already been completed

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warnmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

## Results
## Green - no forwarding enabled and no forwarding address present
## Yellow - forwarding disabled but forwarding address present
## Red - forwarding enabled

write-host -foregroundcolor $processmessagecolor "`nCheck Exchange Forwards"

foreach ($mailbox in $mailboxes) {
    if ($mailbox.DeliverToMailboxAndForward) { ## if email forwarding is active
        Write-host
        Write-host "**********" -foregroundcolor $errormessagecolor
        Write-Host "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress) - Forwarding = $($mailbox.delivertomailboxandforward)" -foregroundColor $errormessagecolor
        Write-host "Forwarding address = $($mailbox.forwardingsmtpaddress)" -foregroundColor $errormessagecolor
        Write-host "**********" -foregroundcolor $errormessagecolor
        write-host
    }
    else {
        if ($mailbox.forwardingsmtpaddress){ ## if email forward email address has been set
            Write-host
            Write-host "**********" -foregroundcolor $warnmessagecolor
            Write-Host "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress) - Forwarding = $($mailbox.delivertomailboxandforward)" -foregroundColor $warnmessagecolor
            Write-host "Forwarding address = $($mailbox.forwardingsmtpaddress)" -foregroundColor $warnmessagecolor
            Write-host "**********" -foregroundcolor $warnmessagecolor
            write-host
        }
        else {
            Write-Host "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress) - Forwarding = $($mailbox.delivertomailboxandforward)" -foregroundColor Gray
        }
    }
}

write-host -foregroundcolor $processmessagecolor "`nCheck Outlook Rule Forwards"

foreach ($mailbox in $mailboxes)
{
  Write-Host -foregroundcolor gray "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress)"
  $rules = get-inboxrule -mailbox $mailbox.identity
  foreach ($rule in $rules)
    {
       If ($rule.enabled) {
        if ($rule.forwardto -or $rule.RedirectTo -or $rule.CopyToFolder -or $rule.DeleteMessage -or $rule.ForwardAsAttachmentTo -or $rule.SendTextMessageNotificationTo) { write-host "`nSuspect Enabled Rule name -",$rule.name }
        If ($rule.forwardto) { write-host -ForegroundColor $errormessagecolor "Forward to:",$rule.forwardto }
        If ($rule.RedirectTo) { write-host -ForegroundColor $errormessagecolor "Redirect to:",$rule.redirectto }
        If ($rule.CopyToFolder) { write-host -ForegroundColor $errormessagecolor "Copy to folder:",$rule.copytofolder }
        if ($rule.DeleteMessage) { write-host -ForegroundColor $errormessagecolor "Delete message:", $rule.deletemessage }
        if ($rule.ForwardAsAttachmentTo) { write-host -ForegroundColor $errormessagecolor "Forward as attachment to:",$rule.forwardasattachmentto}
        if ($rule.SendTextMessageNotificationTo) { write-host -ForegroundColor $errormessagecolor "Sent TXT msg to:",$rule.sendtextmessagenotificationto }
        }
        else {
        if ($rule.forwardto -or $rule.RedirectTo -or $rule.CopyToFolder -or $rule.DeleteMessage -or $rule.ForwardAsAttachmentTo -or $rule.SendTextMessageNotificationTo) { write-host "`nSuspect Disabled Rule name -",$rule.name }
        If ($rule.forwardto) { write-host -ForegroundColor $warnmessagecolor "Forward to:",$rule.forwardto }
        If ($rule.RedirectTo) { write-host -ForegroundColor $warnmessagecolor "Redirect to:",$rule.redirectto }
        If ($rule.CopyToFolder) { write-host -ForegroundColor $warnmessagecolor "Copy to folder:",$rule.copytofolder }
        if ($rule.DeleteMessage) { write-host -ForegroundColor $warnmessagecolor "Delete message:", $rule.deletemessage }
        if ($rule.ForwardAsAttachmentTo) { write-host -ForegroundColor $warnmessagecolor "Forward as attachment to:",$rule.forwardasattachmentto}
        if ($rule.SendTextMessageNotificationTo) { write-host -ForegroundColor $warnmessagecolor "Sent TXT msg to:",$rule.sendtextmessagenotificationto }
        }
    }
}

write-host -foregroundcolor $processmessagecolor "`nCheck Sweep Rules"

foreach ($mailbox in $mailboxes)
{
    Write-Host -foregroundcolor gray "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress)"
  $rules = get-sweeprule -mailbox $mailbox.identity
  foreach ($rule in $rules) {
    if ($rule.enabled) { ## if email forwarding is active
        Write-host
        Write-host "**********" -foregroundcolor $errormessagecolor
        Write-Host -foregroundcolor $errormessagecolor "Sweep rules enabled for $($mailbox.displayname) - $($mailbox.primarysmtpaddress)"
        Write-host -foregroundColor $errormessagecolor "Name = ",$rule.name
        Write-host -foregroundColor $errormessagecolor "Source Folder = ",$rule.sourcefolder 
        write-host -foregroundColor $errormessagecolor "Destination folder = ",$rule.destinationfolder
        Write-host -foregroundColor $errormessagecolor "Keep for days = ",$rule.keepfordays
        Write-host "**********" -foregroundcolor $errormessagecolor
        write-host
        }
    }
}

write-host -foregroundcolor $systemmessagecolor "`nScript complete`n"