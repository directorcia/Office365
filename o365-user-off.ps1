## CIAOPS
## Script provided as is. Use at own risk. No guarantees or warranty provided.

## Description
## Script designed to disable a user's access to Office 365 services

## Prerequisites = 1
## 1. Ensure connected to Office 365 Azure AD
## 2. Ensure connected to Exchange Online
## 3. Ensure connected to SharePoint Online

Clear-Host

write-host -foregroundcolor Cyan "Script start"

## set-executionpolicy remotesigned
## May be required once to allow ability to runs scripts in PowerShell

$useremail=read-host -prompt 'Enter user email address'
try {   ## See whether input matches a user in Azure AD
    $user= get-azureaduser -objectid $useremail -erroraction stop
}
catch
{       # if there is no match provide warning and terminate script
    write-host $useremail,"doesn't appear to be a valid AD user account" -ForegroundColor Red
    return
}

write-host -foregroundcolor green "Found",$user.displayname

## Disable account to block user logins
Set-AzureADUser -objectid $user.ObjectId -AccountEnabled $false
write-host -foregroundcolor green "Disabled login"

## Invalidates all the refresh tokens used to obtain new access tokens for Office 365 applications by setting their expiry to the current date and time. 
## When a user authenticates to connect to an Office 365 application, they create a session with that application. 
## The session receives an access token and a refresh token from Azure Active Directory. 
## An Office 365 access token is valid for an hour (the period can be changed if needed). 
## When that period elapses, an automatic reauthentication process commences to obtain a new access token to allow the session to continue
Revoke-AzureADUserAllRefreshToken -ObjectId $user.ObjectId
write-host -foregroundcolor green "Revoked Token"

## The ActiveSyncEnabled parameter enables or disables Exchange ActiveSync for the mailbox.
Set-CASMailbox $useremail -ActiveSyncEnabled $false
write-host -foregroundcolor green "ActiveSync disabled"

## The OWAEnabled parameter enables or disables access to the mailbox by using Outlook on the web
Set-casmailbox $useremail -owaenabled $false
write-host -foregroundcolor green "OWA disabled"

## TheActiveSyncAllowedDeviceIDs parameter specifies one or more Exchange ActiveSync device IDs that are allowed to synchronize with the mailbox.
## Setting this to $NULL clears the list of device IDs
Set-casmailbox $useremail -activesyncalloweddeviceids $null
write-host -foregroundcolor green "Removed allowed ActiveSync devices"

## The MAPIEnabled parameter enables or disables access to the mailbox by using MAPI clients (for example, Outlook).
Set-casmailbox $useremail -mapienabled $false
write-host -foregroundcolor green "Disabled MAPI"

## The OWAforDevicesEnabled parameter enables or disables access to the mailbox by using Outlook on the web for devices.
Set-casmailbox $useremail -OWAforDevicesEnabled $false
write-host -foregroundcolor green "Disabled MAPI fo devices"

## The PopEnabled parameter enables or disables access to the mailbox by using POP3 clients.
Set-casmailbox $useremail -popenabled $false
write-host -foregroundcolor green "Disabled POP"

## The ImapEnabled parameter enables or disables access to the mailbox by using IMAP4 clients.
Set-casmailbox $useremail -imapenabled $false
write-host -foregroundcolor green "Disabled IMAP"

## The UniversalOutlookEnabled parameter enables or disables access to the mailbox by using Mail and Calendar
Set-casmailbox $useremail -universaloutlookenabled $false
write-host -foregroundcolor green "Disabled Outlook"

## User will be signed out of browser, desktop and mobile applications accessing Office 365 resources across all devices.
## It can take up to an hour to sign out from all devices.
Revoke-SPOUserSession -user $useremail -Confirm:$false

write-host -foregroundcolor cyan "Ending script"