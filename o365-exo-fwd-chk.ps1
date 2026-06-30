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

## Parameters
[CmdletBinding()]
param(
    [string]$LogFile = (Join-Path -Path $PSScriptRoot -ChildPath "o365-exo-fwd-chk-log.txt"),
    [switch]$VerboseOutput = $false                # Enable verbose output
)

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warnmessagecolor = "yellow"

$script:LogFile = $LogFile

## Functions
function Write-LogMessage {
    param (
        [string]$Message,
        [string]$Color = "White"
    )

    $timestampedMessage = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host -ForegroundColor $Color $Message
    Add-Content -Path $script:LogFile -Value $timestampedMessage
}

function Get-ShortText {
    param(
        [AllowNull()]
        [string]$Text,
        [int]$MaxLength
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return "<blank>"
    }

    return $Text.Substring(0, [Math]::Min($MaxLength, $Text.Length))
}

## Start Script
Clear-Host

$logDirectory = Split-Path -Path $script:LogFile -Parent
if ($logDirectory -and -not (Test-Path -Path $logDirectory)) {
    New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
}

"" | Set-Content -Path $script:LogFile
Write-LogMessage "Script started`n" $systemmessagecolor

try {
    if (-not (Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue)) {
        throw "Exchange Online cmdlets are not available. Connect first by running Connect-ExchangeOnline."
    }

    ## Get all mailboxes
    Write-LogMessage "[INFO] Get all mailbox details - Start" $processmessagecolor
    $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
    $mailboxCount = @($mailboxes).Count
    Write-LogMessage "[INFO] Retrieved $mailboxCount mailbox entries" $processmessagecolor
    Write-LogMessage "[INFO] Get all mailbox details - Finish`n" $processmessagecolor

    $mailboxForwardEnabledCount = 0
    $mailboxForwardDisabledCount = 0
    $inboxForwardRuleCount = 0
    $inboxRedirectRuleCount = 0
    $inboxRuleErrorCount = 0
    $sweepRuleCount = 0
    $sweepRuleErrorCount = 0

    ## Check Mailbox Forwards
    Write-LogMessage "Check Mailbox Forwards - Start`n" $processmessagecolor

    foreach ($mailbox in $mailboxes) {
        $shortenedName = Get-ShortText -Text $mailbox.DisplayName -MaxLength 40
        $shortenedUPN = Get-ShortText -Text $mailbox.UserPrincipalName -MaxLength 60
        if ($VerboseOutput) { Write-LogMessage "Mailbox forwards for $shortenedName - $shortenedUPN" "Gray" }

        if ($mailbox.DeliverToMailboxAndForward) {
            $mailboxForwardEnabledCount++
            Write-LogMessage "    Forwarding enabled for $shortenedName - Forwarding = $($mailbox.delivertomailboxandforward)" $errormessagecolor
            Write-LogMessage "    Forwarding address = $($mailbox.forwardingsmtpaddress)" $errormessagecolor
        } elseif ($mailbox.forwardingsmtpaddress) {
            $mailboxForwardDisabledCount++
            Write-LogMessage "    Forwarding address set but disabled for $shortenedName - Forwarding = $($mailbox.delivertomailboxandforward)" $warnmessagecolor
            Write-LogMessage "    Forwarding address = $($mailbox.forwardingsmtpaddress)" $warnmessagecolor
        }
    }

    Write-LogMessage "`nCheck Mailbox Forwards - Finish`n" $processmessagecolor

    ## Check Outlook Rule Forwards
    Write-LogMessage "Check Outlook Rule Forwards - Start`n" $processmessagecolor

    foreach ($mailbox in $mailboxes) {
        $shortenedName = Get-ShortText -Text $mailbox.DisplayName -MaxLength 40
        $shortenedUPN = Get-ShortText -Text $mailbox.UserPrincipalName -MaxLength 60
        if ($VerboseOutput) { Write-LogMessage "Outlook forwards for $shortenedName - $shortenedUPN" "Gray" }

        try {
            $rules = Get-InboxRule -Mailbox $mailbox.UserPrincipalName -ErrorAction Stop
            foreach ($rule in $rules) {
                if ($rule.Enabled) {
                    if ($rule.ForwardTo) {
                        $inboxForwardRuleCount++
                        Write-LogMessage "    Forward to: $($rule.ForwardTo -join ', ')" $errormessagecolor
                    }
                    if ($rule.RedirectTo) {
                        $inboxRedirectRuleCount++
                        Write-LogMessage "    Redirect to: $($rule.RedirectTo -join ', ')" $errormessagecolor
                    }
                }
            }
        } catch {
            $inboxRuleErrorCount++
            Write-LogMessage "    Error retrieving rules for ${shortenedName}: $($_.Exception.Message)" $errormessagecolor
        }
    }

    Write-LogMessage "`nCheck Outlook Rule Forwards - Finish`n" $processmessagecolor

    ## Check Sweep Rules
    Write-LogMessage "Check Sweep Rules - Start`n" $processmessagecolor

    foreach ($mailbox in $mailboxes) {
        $shortenedName = Get-ShortText -Text $mailbox.DisplayName -MaxLength 40
        $shortenedUPN = Get-ShortText -Text $mailbox.UserPrincipalName -MaxLength 60
        if ($VerboseOutput) { Write-LogMessage "Sweep forwards for $shortenedName - $shortenedUPN" "Gray" }

        try {
            $rules = Get-SweepRule -Mailbox $mailbox.UserPrincipalName -ErrorAction Stop
            foreach ($rule in $rules) {
                if ($rule.Enabled) {
                    $sweepRuleCount++
                    Write-LogMessage "    Sweep rule enabled for $shortenedName" $errormessagecolor
                    Write-LogMessage "    Name = $($rule.Name)" $errormessagecolor
                    Write-LogMessage "    Source Folder = $($rule.SourceFolder)" $errormessagecolor
                    Write-LogMessage "    Destination Folder = $($rule.DestinationFolder)" $errormessagecolor
                }
            }
        } catch {
            $sweepRuleErrorCount++
            Write-LogMessage "    Error retrieving sweep rules for ${shortenedName}: $($_.Exception.Message)" $errormessagecolor
        }
    }

    Write-LogMessage "`nCheck Sweep Rules - Finish`n" $processmessagecolor

    Write-LogMessage "Summary" $systemmessagecolor
    Write-LogMessage "    Mailbox forwarding enabled = $mailboxForwardEnabledCount" $systemmessagecolor
    Write-LogMessage "    Mailbox forwarding address set but disabled = $mailboxForwardDisabledCount" $systemmessagecolor
    Write-LogMessage "    Inbox rules with ForwardTo = $inboxForwardRuleCount" $systemmessagecolor
    Write-LogMessage "    Inbox rules with RedirectTo = $inboxRedirectRuleCount" $systemmessagecolor
    Write-LogMessage "    Inbox rule retrieval errors = $inboxRuleErrorCount" $warnmessagecolor
    Write-LogMessage "    Sweep rules enabled = $sweepRuleCount" $systemmessagecolor
    Write-LogMessage "    Sweep rule retrieval errors = $sweepRuleErrorCount" $warnmessagecolor

} catch {
    Write-LogMessage "An error occurred: $($_.Exception.Message)" $errormessagecolor
} finally {
    Write-LogMessage "Script complete" $systemmessagecolor
    Write-LogMessage "Log file: $script:LogFile`n" $processmessagecolor
}
