<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/o365-user-off.ps1

Description - Disable a user's access to Office 365 services

Prerequisites = 3
1. Ensure connected to Azure AD (AzureAD module)
2. Ensure connected to Exchange Online (ExchangeOnlineManagement module)
3. Ensure connected to SharePoint Online (Microsoft.Online.SharePoint.PowerShell module)

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script start`n"

$useremail=read-host -prompt 'Enter user email address'
try {   ## See whether input matches a user in Azure AD
    $user= get-azureaduser -objectid $useremail -erroraction stop
    write-host -ForegroundColor $processmessagecolor "Reset user's password on local AD if synced or in cloud"
    Read-Host -prompt "Press any key to continue" 
}
catch
{       # if there is no match provide warning and terminate script
    write-host $useremail,"doesn't appear to be a valid AD user account" -ForegroundColor Red
    return
}

write-host -foregroundcolor $processmessagecolor "Found",$user.displayname

## Disable account to block user logins
Set-AzureADUser -objectid $user.ObjectId -AccountEnabled $false
write-host -foregroundcolor $processmessagecolor "Disabled login"

## Invalidates all the refresh tokens used to obtain new access tokens for Office 365 applications by setting their expiry to the current date and time.
## When a user authenticates to connect to an Office 365 application, they create a session with that application.
## The session receives an access token and a refresh token from Azure Active Directory.
## An Office 365 access token is valid for an hour (the period can be changed if needed).
## When that period elapses, an automatic reauthentication process commences to obtain a new access token to allow the session to continue
Revoke-AzureADUserAllRefreshToken -ObjectId $user.ObjectId
write-host -foregroundcolor $processmessagecolor "Revoked Token"

## The ActiveSyncEnabled parameter enables or disables Exchange ActiveSync for the mailbox.
Set-CASMailbox $useremail -ActiveSyncEnabled $false
write-host -foregroundcolor $processmessagecolor "ActiveSync disabled"

## The OWAEnabled parameter enables or disables access to the mailbox by using Outlook on the web
Set-casmailbox $useremail -owaenabled $false
write-host -foregroundcolor $processmessagecolor "OWA disabled"

## TheActiveSyncAllowedDeviceIDs parameter specifies one or more Exchange ActiveSync device IDs that are allowed to synchronize with the mailbox.
## Setting this to $NULL clears the list of device IDs
Set-casmailbox $useremail -activesyncalloweddeviceids $null
write-host -foregroundcolor $processmessagecolor "Removed allowed ActiveSync devices"

## The MAPIEnabled parameter enables or disables access to the mailbox by using MAPI clients (for example, Outlook).
Set-casmailbox $useremail -mapienabled $false
write-host -foregroundcolor $processmessagecolor "Disabled MAPI"

## The OWAforDevicesEnabled parameter enables or disables access to the mailbox by using Outlook on the web for devices.
Set-casmailbox $useremail -OWAforDevicesEnabled $false
write-host -foregroundcolor $processmessagecolor "Disabled MAPI fo devices"

## The PopEnabled parameter enables or disables access to the mailbox by using POP3 clients.
Set-casmailbox $useremail -popenabled $false
write-host -foregroundcolor $processmessagecolor "Disabled POP"

## The ImapEnabled parameter enables or disables access to the mailbox by using IMAP4 clients.
Set-casmailbox $useremail -imapenabled $false
write-host -foregroundcolor $processmessagecolor "Disabled IMAP"

## The UniversalOutlookEnabled parameter enables or disables access to the mailbox by using Mail and Calendar
Set-casmailbox $useremail -universaloutlookenabled $false
write-host -foregroundcolor $processmessagecolor "Disabled Outlook"

## User will be signed out of browser, desktop and mobile applications accessing Office 365 resources across all devices.
## It can take up to an hour to sign out from all devices.
Revoke-SPOUserSession -user $useremail -Confirm:$false

write-host -foregroundcolor $systemmessagecolor "Script completed`n"