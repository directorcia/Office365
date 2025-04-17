<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Check which email boxes have forwarding options set.
Will check mailbox forwarding, rules set by Outlook client and Sweep setting

Source - https://github.com/directorcia/Office365/blob/master/o365-exo-fwd-chk.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Report-email-forwards

Prerequisites = 1
1. Ensure connection to Exchange Online has already been completed

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warnmessagecolor = "yellow"

## Parameters
param(
    [string]$LogFile, # Log file for output
    [switch]$VerboseOutput = $false                # Enable verbose output
)

## Fixing the invalid assignment expression error
# The error might occur if the $LogFile variable is being used incorrectly elsewhere in the script.
# Ensuring that $LogFile is properly initialized and used as a string path.

# Correcting the initialization of $LogFile to use a full path for clarity
$LogFile = "..\o365-exo-fwd-chk-log.txt"

# Ensuring all references to $LogFile are valid and consistent throughout the script.

## Functions
function Log-Message {
    param (
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host -ForegroundColor $Color $Message
    Add-Content -Path $LogFile -Value $Message
}

## Start Script
Clear-Host
Log-Message "Script started`n" $systemmessagecolor

try {
    # Ensure the user is connected to Exchange Online
    if (-not (Get-Module -Name ExchangeOnlineManagement)) {
        throw "The term 'Get-Mailbox' is not recognized as a name of a cmdlet. Please ensure you are connected to Exchange Online first. Use the 'Connect-ExchangeOnline' cmdlet to establish a connection."
    }

    ## Get all mailboxes
    Log-Message "[INFO] Get all mailbox details - Start" $processmessagecolor
    $mailboxes = Get-Mailbox -ResultSize Unlimited
    Log-Message "[INFO] Get all mailbox details - Finish`n" $processmessagecolor

    ## Check Mailbox Forwards
    Log-Message "Check Mailbox Forwards - Start`n" $processmessagecolor

    foreach ($mailbox in $mailboxes) {
        $shortenedName = $mailbox.displayname.Substring(0, [Math]::Min(40, $mailbox.displayname.Length))
        $shortenedUPN = $mailbox.UserPrincipalName.Substring(0, [Math]::Min(60, $mailbox.UserPrincipalName.Length))
        Log-Message "Mailbox forwards for $shortenedName - $shortenedUPN" "Gray"
        if ($mailbox.DeliverToMailboxAndForward) {
            Log-Message "    Forwarding enabled for $shortenedName - Forwarding = $($mailbox.delivertomailboxandforward)" $errormessagecolor
            Log-Message "    Forwarding address = $($mailbox.forwardingsmtpaddress)" $errormessagecolor
        } elseif ($mailbox.forwardingsmtpaddress) {
            Log-Message "    Forwarding address set but disabled for $shortenedName - Forwarding = $($mailbox.delivertomailboxandforward)" $warnmessagecolor
            Log-Message "    Forwarding address = $($mailbox.forwardingsmtpaddress)" $warnmessagecolor
        }
    }

    Log-Message "`nCheck Mailbox Forwards - Finish`n" $processmessagecolor

    ## Check Outlook Rule Forwards
    Log-Message "Check Outlook Rule Forwards - Start`n" $processmessagecolor

    foreach ($mailbox in $mailboxes) {
        $shortenedName = $mailbox.displayname.Substring(0, [Math]::Min(40, $mailbox.displayname.Length))
        $shortenedUPN = $mailbox.UserPrincipalName.Substring(0, [Math]::Min(60, $mailbox.UserPrincipalName.Length))
        Log-Message "Outlook forwards for $shortenedName - $shortenedUPN" "Gray"
        try {
            $rules = Get-InboxRule -Mailbox $mailbox.UserPrincipalName
            foreach ($rule in $rules) {
                if ($rule.Enabled) {
                    if ($rule.ForwardTo) { Log-Message "    Forward to: $($rule.ForwardTo -join ', ')" $errormessagecolor }
                    if ($rule.RedirectTo) { Log-Message "    Redirect to: $($rule.RedirectTo -join ', ')" $errormessagecolor }
                }
            }
        } catch {
            Log-Message "    Error retrieving rules for $shortenedname $_" $errormessagecolor
        }
    }

    Log-Message "`nCheck Outlook Rule Forwards - Finish`n" $processmessagecolor

    ## Check Sweep Rules
    Log-Message "Check Sweep Rules - Start`n" $processmessagecolor

    foreach ($mailbox in $mailboxes) {
        $shortenedName = $mailbox.displayname.Substring(0, [Math]::Min(40, $mailbox.displayname.Length))
        $shortenedUPN = $mailbox.UserPrincipalName.Substring(0, [Math]::Min(60, $mailbox.UserPrincipalName.Length))
        Log-Message "Sweep forwards for $shortenedName - $shortenedUPN" "Gray"
        try {
            $rules = Get-SweepRule -Mailbox $mailbox.UserPrincipalName
            foreach ($rule in $rules) {
                if ($rule.Enabled) {
                    Log-Message "    Sweep rule enabled for $shortenedName" $errormessagecolor
                    Log-Message "    Name = $($rule.Name)" $errormessagecolor
                    Log-Message "    Source Folder = $($rule.SourceFolder)" $errormessagecolor
                    Log-Message "    Destination Folder = $($rule.DestinationFolder)" $errormessagecolor
                }
            }
        } catch {
            Log-Message "    Error retrieving sweep rules for $shortenedName $_" $errormessagecolor
        }
    }

    Log-Message "`nCheck Sweep Rules - Finish`n" $processmessagecolor

} catch {
    Log-Message "An error occurred: $_" $errormessagecolor
} finally {
    Log-Message "Script complete`n" $systemmessagecolor
}