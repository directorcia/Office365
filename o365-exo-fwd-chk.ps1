## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to check which email boxes have forwarding options set

## Prerequisites = 1
## 1. Ensure connection to Exchange Online has already been completed

Clear-Host

write-host -foregroundcolor green "Script started"

## Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

## Results
## Green - no forwarding enabled and no forwarding address present
## Yellow - forwarding disabled but forwarding address present
## Red - forwarding enabled
 
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
write-host -foregroundcolor green "Script complete"
