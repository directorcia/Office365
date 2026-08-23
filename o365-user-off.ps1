<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/o365-user-off.ps1

Description - Disable a user's access to Office 365 services

Prerequisites = 3
1. Microsoft Graph (Microsoft.Graph.Users module) - replaces the retired AzureAD module
2. Exchange Online (ExchangeOnlineManagement module)
3. SharePoint Online (Microsoft.Online.SharePoint.PowerShell module) - optional, only for SPO session revocation

More scripts available by joining http://www.ciaopspatron.com

Improvements over original:
- Replaced retired AzureAD cmdlets with Microsoft Graph equivalents
- Added a param() block so no interactive prompt is needed when parameters are supplied
- Verifies module availability and active connections before starting (connects if missing)
- Per-step error handling so one failure doesn't abort the whole offboarding
- Combined all Set-CASMailbox changes into a single splatted call
- Added -WhatIf/-Confirm support and an end-of-run summary table
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory, Position = 0, HelpMessage = 'Enter user email address / UPN')]
    [ValidateNotNullOrEmpty()]
    [string]$UserPrincipalName,

    [Parameter(HelpMessage = 'e.g. https://contoso-admin.sharepoint.com')]
    [ValidateNotNullOrEmpty()]
    [string]$TenantAdminUrl
)

$systemmessagecolor  = 'cyan'
$processmessagecolor = 'green'

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

$results = [System.Collections.Generic.List[psobject]]::new()
function Add-StepResult {
    param(
        [string]$Step,
        [ValidateSet('Success', 'Failed', 'Skipped')][string]$Status,
        [string]$Detail = ''
    )
    $results.Add([pscustomobject]@{ Step = $Step; Status = $Status; Detail = $Detail })
    $color = if ($Status -eq 'Failed') { 'red' } else { $processmessagecolor }
    Write-Host -ForegroundColor $color "$Step - $Status $(if ($Detail) { "($Detail)" })"
}

Clear-Host
Write-Host -ForegroundColor $systemmessagecolor "Script start`n"

## --- Check prerequisites and connections --------------------------------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    throw 'Microsoft.Graph.Users module not found. Install it with: Install-Module Microsoft.Graph.Users -Scope CurrentUser'
}
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw 'ExchangeOnlineManagement module not found. Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
}

if (-not (Get-MgContext)) {
    Write-Host -ForegroundColor $processmessagecolor 'Connecting to Microsoft Graph...'
    Connect-MgGraph -Scopes 'User.ReadWrite.All' -NoWelcome
}
if (-not (Get-ConnectionInformation)) {
    Write-Host -ForegroundColor $processmessagecolor 'Connecting to Exchange Online...'
    Connect-ExchangeOnline -ShowBanner:$false
}

$spoAvailable = $false
if ($TenantAdminUrl) {
    if (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell) {
        try {
            Connect-SPOService -Url $TenantAdminUrl -ErrorAction Stop
            $spoAvailable = $true
        }
        catch {
            Write-Warning "Could not connect to SharePoint Online: $($_.Exception.Message). Session revocation will be skipped."
        }
    }
    else {
        Write-Warning 'Microsoft.Online.SharePoint.PowerShell module not found. Session revocation will be skipped.'
    }
}
else {
    Write-Host -ForegroundColor $systemmessagecolor 'No -TenantAdminUrl supplied; skipping SharePoint Online session revocation.'
}

## --- Validate user ------------------------------------------------------------
try {
    $user = Get-MgUser -UserId $UserPrincipalName -Property 'Id,DisplayName,UserPrincipalName,AccountEnabled' -ErrorAction Stop
}
catch {
    Write-Host "$UserPrincipalName doesn't appear to be a valid Entra ID user account" -ForegroundColor Red
    Add-StepResult -Step 'User lookup' -Status Failed -Detail $_.Exception.Message
    return
}
Write-Host -ForegroundColor $processmessagecolor "Found $($user.DisplayName) ($($user.UserPrincipalName))"
Write-Host -ForegroundColor $processmessagecolor 'Reminder: reset the user''s password in on-prem AD (if synced) or in the cloud (if cloud-only)'
Read-Host -Prompt 'Press Enter to continue'

## --- 1. Disable account to block user logins ----------------------------------
try {
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Disable account')) {
        Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop
        Add-StepResult -Step 'Disable login' -Status Success
    }
}
catch { Add-StepResult -Step 'Disable login' -Status Failed -Detail $_.Exception.Message }

## --- 2. Revoke all refresh tokens / sign-in sessions --------------------------
## Invalidates all refresh tokens used to obtain new access tokens for Office 365 applications.
## An access token is valid for an hour; revoking sessions forces re-authentication on expiry.
try {
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Revoke all sign-in sessions and refresh tokens')) {
        $null = Revoke-MgUserSignInSession -UserId $user.Id -ErrorAction Stop
        Add-StepResult -Step 'Revoke tokens' -Status Success -Detail 'may take up to an hour to take full effect'
    }
}
catch { Add-StepResult -Step 'Revoke tokens' -Status Failed -Detail $_.Exception.Message }

## --- 3. Disable all mailbox client access protocols ---------------------------
## ActiveSync, OWA, MAPI (Outlook), OWA for Devices, POP3, IMAP4 and Universal Outlook (Mail and Calendar).
## Also clears the list of allowed ActiveSync device IDs.
$mailboxParams = @{
    Identity                   = $UserPrincipalName
    ActiveSyncEnabled          = $false
    OWAEnabled                 = $false
    OWAforDevicesEnabled       = $false
    MAPIEnabled                = $false
    PopEnabled                 = $false
    ImapEnabled                = $false
    UniversalOutlookEnabled    = $false
    ActiveSyncAllowedDeviceIDs = $null
    ErrorAction                = 'Stop'
}
try {
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Disable all mailbox protocols (EAS, OWA, MAPI, POP, IMAP, Universal Outlook)')) {
        Set-CASMailbox @mailboxParams
        Add-StepResult -Step 'Disable mailbox protocols' -Status Success
    }
}
catch { Add-StepResult -Step 'Disable mailbox protocols' -Status Failed -Detail $_.Exception.Message }

## --- 4. Sign user out of SharePoint Online / OneDrive -------------------------
## Signs the user out of browser, desktop and mobile applications across all devices (can take up to an hour).
if ($spoAvailable) {
    try {
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Revoke SharePoint Online sessions')) {
            Revoke-SPOUserSession -User $UserPrincipalName -Confirm:$false -ErrorAction Stop
            Add-StepResult -Step 'Revoke SPO sessions' -Status Success
        }
    }
    catch { Add-StepResult -Step 'Revoke SPO sessions' -Status Failed -Detail $_.Exception.Message }
}
else {
    Add-StepResult -Step 'Revoke SPO sessions' -Status Skipped -Detail 'not connected to SharePoint Online'
}

## --- Summary ------------------------------------------------------------------
Write-Host ''
$results | Format-Table -AutoSize
if ($results.Status -contains 'Failed') {
    Write-Warning 'One or more steps failed - review the summary above.'
}
Write-Host -ForegroundColor $systemmessagecolor "Script completed`n"
