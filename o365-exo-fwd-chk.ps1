## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to check which email boxes have forwarding options set

## Prerequisites = 1
## 1. Ensure connection to Exchange Online has already been completed

Clear-Host

write-host -foregroundcolor Cyan "`nScript started"

## Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

## Results
## Green - no forwarding enabled and no forwarding address present
## Yellow - forwarding disabled but forwarding address present
## Red - forwarding enabled

write-host -foregroundcolor Cyan "`nCheck Exchange Forwards"
 
foreach ($mailbox in $mailboxes) {
    if ($mailbox.DeliverToMailboxAndForward) { ## if email forwarding is active
        Write-host
        Write-host "**********" -foregroundcolor red        
        Write-Host "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress) - Forwarding = $($mailbox.delivertomailboxandforward)" -foregroundColor Red
        Write-host "Forwarding address = $($mailbox.forwardingsmtpaddress)" -foregroundColor Red
        Write-host "**********" -foregroundcolor red
        write-host
    }
    else {
        if ($mailbox.forwardingsmtpaddress){ ## if email forward email address has been set
            Write-host
            Write-host "**********" -foregroundcolor yellow        
            Write-Host "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress) - Forwarding = $($mailbox.delivertomailboxandforward)" -foregroundColor yellow
            Write-host "Forwarding address = $($mailbox.forwardingsmtpaddress)" -foregroundColor yellow
            Write-host "**********" -foregroundcolor yellow
            write-host
        }
        else {
            Write-Host "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress) - Forwarding = $($mailbox.delivertomailboxandforward)" -foregroundColor Green
        }
    }
}

write-host -foregroundcolor Cyan "`nCheck Outlook Rule Forwards"

foreach ($mailbox in $mailboxes) 
{
    Write-Host -foregroundcolor gray "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress)"
  $rules = get-inboxrule -mailbox $mailbox.identity 
  foreach ($rule in $rules)
    {
       If ($rule.enabled) {
        if ($rule.forwardto -or $rule.RedirectTo -or $rule.CopyToFolder -or $rule.DeleteMessage -or $rule.ForwardAsAttachmentTo -or $rule.SendTextMessageNotificationTo) { write-host "`nSuspect Enabled Rule name -",$rule.name }      
        If ($rule.forwardto) { write-host -ForegroundColor red "Forward to:",$rule.forwardto }
        If ($rule.RedirectTo) { write-host -ForegroundColor red "Redirect to:",$rule.redirectto }
        If ($rule.CopyToFolder) { write-host -ForegroundColor red "Copy to folder:",$rule.copytofolder }
        if ($rule.DeleteMessage) { write-host -ForegroundColor Red "Delete message:", $rule.deletemessage }
        if ($rule.ForwardAsAttachmentTo) { write-host -ForegroundColor Red "Forward as attachment to:",$rule.forwardasattachmentto}
        if ($rule.SendTextMessageNotificationTo) { write-host -ForegroundColor Red "Sent TXT msg to:",$rule.sendtextmessagenotificationto }
        }
        else {
        if ($rule.forwardto -or $rule.RedirectTo -or $rule.CopyToFolder -or $rule.DeleteMessage -or $rule.ForwardAsAttachmentTo -or $rule.SendTextMessageNotificationTo) { write-host "`nSuspect Disabled Rule name -",$rule.name }      
        If ($rule.forwardto) { write-host -ForegroundColor Yellow "Forward to:",$rule.forwardto }
        If ($rule.RedirectTo) { write-host -ForegroundColor Yellow "Redirect to:",$rule.redirectto }
        If ($rule.CopyToFolder) { write-host -ForegroundColor Yellow "Copy to folder:",$rule.copytofolder }
        if ($rule.DeleteMessage) { write-host -ForegroundColor Yellow "Delete message:", $rule.deletemessage }
        if ($rule.ForwardAsAttachmentTo) { write-host -ForegroundColor Yellow "Forward as attachment to:",$rule.forwardasattachmentto}
        if ($rule.SendTextMessageNotificationTo) { write-host -ForegroundColor Yellow "Sent TXT msg to:",$rule.sendtextmessagenotificationto }
        }
    }
}

write-host -foregroundcolor Cyan "`nScript complete"
