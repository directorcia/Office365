## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Source = https://github.com/directorcia/Office365/blob/master/o365-exo-addins.ps1

## Description
## Script designed to check which which add ins are present on each mailbox

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

write-host -foregroundcolor Cyan "`nCheck Mailbox Add ins"

foreach ($mailbox in $mailboxes) {
    write-host "Mailbox =",$mailbox.primarysmtpaddress
    get-app -mailbox $mailbox.primarysmtpaddress | Select-Object displayname,enabled,appversion | Format-Table
}

write-host -foregroundcolor Cyan "`nScript complete"