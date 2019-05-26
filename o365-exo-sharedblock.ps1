<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report and potentially disable interactive logins to shared mailboxes

## Source - https://github.com/directorcia/Office365/blob/master/o365-exo-sharedblock.ps1

Prerequisites = 2
1. Connected to Exchange Online
2. Connect to Azure AD

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$secure = $false         ## $true = shared mailbox login will be automatically disabled, $false = report only
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

write-host -ForegroundColor $processmessagecolor "Getting shared mailboxes"
$Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited
write-host -ForegroundColor $processmessagecolor "Start checking shared mailboxes"
write-host
foreach ($mailbox in $mailboxes) {
    $accountdetails=get-azureaduser -objectid $mailbox.userprincipalname        ## Get the Azure AD account connected to shared mailbox
    If ($accountdetails.accountenabled){                                        ## if that login is enabled
        Write-host -foregroundcolor $errormessagecolor $mailbox.displayname,"["$mailbox.userprincipalname"] - Direct Login ="$accountdetails.accountenabled
        If ($secure) {                                                          ## if the secure variable is true disable login to shared mailbox
            Set-AzureADUser -ObjectID $mailbox.userprincipalname -AccountEnabled $false     ## disable shared mailbox account
            $accountdetails=get-azureaduser -objectid $mailbox.userprincipalname            ## Get the Azure AD account connected to shared mailbox again
            write-host -ForegroundColor $processmessagecolor "*** SECURED"$mailbox.displayname,"["$mailbox.userprincipalname"] - Direct Login ="$accountdetails.accountenabled
        }
    } else {
        Write-host -foregroundcolor $processmessagecolor $mailbox.displayname,"["$mailbox.userprincipalname"] - Direct Login ="$accountdetails.accountenabled
    }
}
write-host -ForegroundColor $processmessagecolor "`nFinish checking mailboxes"
write-host
write-host -foregroundcolor $systemmessagecolor "Script completed`n"