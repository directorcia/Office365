## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Source = https://github.com/directorcia/Office365/blob/master/o365-exo-addins.ps1

## Description
## Script designed to check which which add ins are present on each mailbox

## Prerequisites = 1
## 1. Ensure connection to Exchange Online has already been completed

## Variables
$systemmessagecolor = "cyan"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "`nScript started"

## Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

write-host -foregroundcolor $systemmessagecolor "`nCheck Mailbox Add ins"

foreach ($mailbox in $mailboxes) {
    write-host "Mailbox =",$mailbox.primarysmtpaddress
    get-app -mailbox $mailbox.primarysmtpaddress | Select-Object displayname,enabled,appversion | Format-Table
}

write-host -foregroundcolor $systemmessagecolor "`nScript complete"