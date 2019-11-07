<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Check which email boxes have forwarding options set.
Will check mailbox forwarding, rules set by Outlook client and Sweep setting

Source - https://github.com/directorcia/Office365/blob/master/o365-exo-fwd-chk.ps1

Prerequisites = 1
1. Ensure connection to Exchange Online has already been completed

More scripts available by joining http://www.ciaopspatron.com

#>

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
write-host -foregroundcolor $processmessagecolor "Get all mailbox details - Start`n"
$mailboxes = Get-Mailbox -ResultSize Unlimited
write-host -foregroundcolor $processmessagecolor "Get all mailbox details - Finish`n"

## Results
## Green - no forwarding enabled and no forwarding address present
## Yellow - forwarding disabled but forwarding address present
## Red - forwarding enabled

write-host -foregroundcolor $processmessagecolor "Check Mailbox Forwards - Start`n"

foreach ($mailbox in $mailboxes) {
    Write-Host -foregroundColor Gray "Mailbox forwards for $($mailbox.displayname) - $($mailbox.primarysmtpaddress)"
    if ($mailbox.DeliverToMailboxAndForward) { ## if email forwarding is active
        Write-Host -foregroundColor $errormessagecolor "`nMailbox forwards for $($mailbox.displayname) - $($mailbox.primarysmtpaddress) - Forwarding = $($mailbox.delivertomailboxandforward)" 
        Write-host -foregroundColor $errormessagecolor "Forwarding address = $($mailbox.forwardingsmtpaddress)`n" 
    }
    else {
        if ($mailbox.forwardingsmtpaddress){ ## if email forward email address has been set
            Write-Host -foregroundColor $warnmessagecolor "`nMailbox forwards for $($mailbox.displayname) - $($mailbox.primarysmtpaddress) - Forwarding = $($mailbox.delivertomailboxandforward)"
            Write-host -foregroundColor $warnmessagecolor "Forwarding address = $($mailbox.forwardingsmtpaddress)`n"
        }
    }
}
write-host -foregroundcolor $processmessagecolor "`nCheck Exchange Forwards - Finish`n"
write-host -foregroundcolor $processmessagecolor "Check Outlook Rule Forwards - Start`n"

foreach ($mailbox in $mailboxes)
{
  Write-Host -foregroundcolor gray "Outlook forwards for $($mailbox.displayname) - $($mailbox.primarysmtpaddress)"
  $rules = get-inboxrule -mailbox $mailbox.userprincipalname
  foreach ($rule in $rules)
    {
       If ($rule.enabled) {
        if ($rule.forwardto -or $rule.RedirectTo -or $rule.CopyToFolder -or $rule.DeleteMessage -or $rule.ForwardAsAttachmentTo -or $rule.SendTextMessageNotificationTo) { write-host -ForegroundColor $warnmessagecolor "`nSuspect Enabled Rule name -",$rule.name }
        If ($rule.forwardto) { write-host -ForegroundColor $errormessagecolor "Forward to:",$rule.forwardto,"`n" }
        If ($rule.RedirectTo) { write-host -ForegroundColor $errormessagecolor "Redirect to:",$rule.redirectto,"`n" }
        If ($rule.CopyToFolder) { write-host -ForegroundColor $errormessagecolor "Copy to folder:",$rule.copytofolder,"`n" }
        if ($rule.DeleteMessage) { write-host -ForegroundColor $errormessagecolor "Delete message:", $rule.deletemessage,"`n" }
        if ($rule.ForwardAsAttachmentTo) { write-host -ForegroundColor $errormessagecolor "Forward as attachment to:",$rule.forwardasattachmentto, "`n"}
        if ($rule.SendTextMessageNotificationTo) { write-host -ForegroundColor $errormessagecolor "Sent TXT msg to:",$rule.sendtextmessagenotificationto, "`n" }
        }
        else {
        if ($rule.forwardto -or $rule.RedirectTo -or $rule.CopyToFolder -or $rule.DeleteMessage -or $rule.ForwardAsAttachmentTo -or $rule.SendTextMessageNotificationTo) { write-host -ForegroundColor $warnmessagecolor "`nSuspect Disabled Rule name -",$rule.name }
        If ($rule.forwardto) { write-host -ForegroundColor $warnmessagecolor "Forward to:",$rule.forwardto,"`n" }
        If ($rule.RedirectTo) { write-host -ForegroundColor $warnmessagecolor "Redirect to:",$rule.redirectto,"`n" }
        If ($rule.CopyToFolder) { write-host -ForegroundColor $warnmessagecolor "Copy to folder:",$rule.copytofolder,"`n" }
        if ($rule.DeleteMessage) { write-host -ForegroundColor $warnmessagecolor "Delete message:", $rule.deletemessage,"`n" }
        if ($rule.ForwardAsAttachmentTo) { write-host -ForegroundColor $warnmessagecolor "Forward as attachment to:",$rule.forwardasattachmentto,"`n"}
        if ($rule.SendTextMessageNotificationTo) { write-host -ForegroundColor $warnmessagecolor "Sent TXT msg to:",$rule.sendtextmessagenotificationto,"`n" }
        }
    }
}
write-host -foregroundcolor $processmessagecolor "`nCheck Outlook Rule Forwards - Finish`n"
write-host -foregroundcolor $processmessagecolor "Check Sweep Rules - Start`n"

foreach ($mailbox in $mailboxes)
{
  Write-Host -foregroundcolor gray "Sweep forwards for $($mailbox.displayname) - $($mailbox.primarysmtpaddress)"
  $rules = get-sweeprule -mailbox $mailbox.userprincipalname
  foreach ($rule in $rules) {
    if ($rule.enabled) { ## if Sweep is active 
        Write-Host -foregroundcolor $errormessagecolor "`nSweep rules enabled for $($mailbox.displayname) - $($mailbox.primarysmtpaddress)"
        Write-host -foregroundColor $errormessagecolor "Name = ",$rule.name
        Write-host -foregroundColor $errormessagecolor "Source Folder = ",$rule.sourcefolder 
        write-host -foregroundColor $errormessagecolor "Destination folder = ",$rule.destinationfolder
        Write-host -foregroundColor $errormessagecolor "Keep for days = ",$rule.keepfordays"`n"
        }
    }
}
write-host -foregroundcolor $processmessagecolor "`nCheck Sweep Rules - Finish`n"
write-host -foregroundcolor $systemmessagecolor "Script complete`n"