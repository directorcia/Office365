<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source = https://github.com/directorcia/Office365/blob/master/o365-exo-addins.ps1

Description - Check which which add ins are present on each mailbox

Prerequisites = 1
1. Ensure connection to Exchange Online has already been completed

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

write-host -foregroundcolor $processmessagecolor "`nCheck Mailbox Add ins"

foreach ($mailbox in $mailboxes) {
    write-host "Mailbox =",$mailbox.primarysmtpaddress
    get-app -mailbox $mailbox.userprincipalname | Select-Object displayname, providername, enabled, appversion | Format-Table
}

write-host -foregroundcolor $systemmessagecolor "Script Completed`n"