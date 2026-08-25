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
    [string]$CsvFile,                              # Optional path to export findings as CSV
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

$fatalError = $false
try {
    if (-not (Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue)) {
        throw "Exchange Online cmdlets are not available. Connect first by running Connect-ExchangeOnline."
    }

    # Check optional cmdlets once rather than failing on every mailbox
    $canCheckInboxRules = [bool](Get-Command -Name Get-InboxRule -ErrorAction SilentlyContinue)
    $canCheckSweepRules = [bool](Get-Command -Name Get-SweepRule -ErrorAction SilentlyContinue)
    if (-not $canCheckInboxRules) { Write-LogMessage "[WARN] Get-InboxRule not available - inbox rule checks will be skipped" $warnmessagecolor }
    if (-not $canCheckSweepRules) { Write-LogMessage "[WARN] Get-SweepRule not available - sweep rule checks will be skipped" $warnmessagecolor }

    ## Get all mailboxes
    Write-LogMessage "[INFO] Get all mailbox details - Start" $processmessagecolor
    if (Get-Command -Name Get-EXOMailbox -ErrorAction SilentlyContinue) {
        # REST-based cmdlet with a minimal property set is much faster in large tenants
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName, UserPrincipalName, DeliverToMailboxAndForward, ForwardingSmtpAddress, ForwardingAddress -ErrorAction Stop
    }
    else {
        $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
    }
    $mailboxCount = @($mailboxes).Count
    Write-LogMessage "[INFO] Retrieved $mailboxCount mailbox entries" $processmessagecolor
    Write-LogMessage "[INFO] Get all mailbox details - Finish`n" $processmessagecolor

    $mailboxForwardKeepCopyCount = 0
    $mailboxForwardOnlyCount = 0
    $inboxForwardRuleCount = 0
    $inboxForwardAttachmentRuleCount = 0
    $inboxRedirectRuleCount = 0
    $inboxRuleErrorCount = 0
    $sweepRuleCount = 0
    $sweepRuleErrorCount = 0
    $findings = [System.Collections.Generic.List[object]]::new()

    ## Check each mailbox for mailbox-level forwards, inbox rules and sweep rules
    Write-LogMessage "Check mailbox forwards, inbox rules and sweep rules - Start`n" $processmessagecolor

    $index = 0
    foreach ($mailbox in $mailboxes) {
        $index++
        $shortenedName = Get-ShortText -Text $mailbox.DisplayName -MaxLength 40
        $shortenedUPN = Get-ShortText -Text $mailbox.UserPrincipalName -MaxLength 60
        $percentComplete = if ($mailboxCount -gt 0) { ($index / $mailboxCount) * 100 } else { 100 }
        Write-Progress -Activity "Checking mailboxes for forwards" -Status "$index of $mailboxCount - $shortenedUPN" -PercentComplete $percentComplete
        if ($VerboseOutput) { Write-LogMessage "Checking $shortenedName - $shortenedUPN" "Gray" }

        ## Mailbox-level forwarding. ForwardingSmtpAddress (user-set) and ForwardingAddress (admin-set) can both be present
        $forwardTargets = @()
        if ($mailbox.ForwardingSmtpAddress) { $forwardTargets += "$($mailbox.ForwardingSmtpAddress)" -replace '^smtp:', '' }
        if ($mailbox.ForwardingAddress) { $forwardTargets += "$($mailbox.ForwardingAddress)" }
        if ($forwardTargets) {
            $forwardTarget = $forwardTargets -join ', '
            if ($mailbox.DeliverToMailboxAndForward) {
                $mailboxForwardKeepCopyCount++
                $detail = "Forwards and keeps a copy in mailbox"
            } else {
                $mailboxForwardOnlyCount++
                $detail = "Forwards without keeping a copy in mailbox"
            }
            Write-LogMessage "    Mailbox forwarding set for $shortenedName - $detail" $errormessagecolor
            Write-LogMessage "    Forwarding address = $forwardTarget" $errormessagecolor
            $findings.Add([pscustomobject]@{
                Mailbox     = $mailbox.DisplayName
                UPN         = $mailbox.UserPrincipalName
                FindingType = "Mailbox forwarding"
                RuleName    = ""
                Target      = "$forwardTarget"
                Detail      = $detail
            })
        }

        ## Inbox rules set via Outlook / OWA
        try {
            if (-not $canCheckInboxRules) { throw "skip" }
            $rules = Get-InboxRule -Mailbox $mailbox.UserPrincipalName -ErrorAction Stop
            foreach ($rule in $rules) {
                if (-not $rule.Enabled) { continue }
                if ($rule.ForwardTo) {
                    $inboxForwardRuleCount++
                    Write-LogMessage "    Inbox rule '$($rule.Name)' on $shortenedName - Forward to: $($rule.ForwardTo -join ', ')" $errormessagecolor
                    $findings.Add([pscustomobject]@{
                        Mailbox     = $mailbox.DisplayName
                        UPN         = $mailbox.UserPrincipalName
                        FindingType = "Inbox rule forward"
                        RuleName    = $rule.Name
                        Target      = ($rule.ForwardTo -join ', ')
                        Detail      = "Rule forwards a copy"
                    })
                }
                if ($rule.ForwardAsAttachmentTo) {
                    $inboxForwardAttachmentRuleCount++
                    Write-LogMessage "    Inbox rule '$($rule.Name)' on $shortenedName - Forward as attachment to: $($rule.ForwardAsAttachmentTo -join ', ')" $errormessagecolor
                    $findings.Add([pscustomobject]@{
                        Mailbox     = $mailbox.DisplayName
                        UPN         = $mailbox.UserPrincipalName
                        FindingType = "Inbox rule forward as attachment"
                        RuleName    = $rule.Name
                        Target      = ($rule.ForwardAsAttachmentTo -join ', ')
                        Detail      = "Rule forwards message as attachment"
                    })
                }
                if ($rule.RedirectTo) {
                    $inboxRedirectRuleCount++
                    Write-LogMessage "    Inbox rule '$($rule.Name)' on $shortenedName - Redirect to: $($rule.RedirectTo -join ', ')" $errormessagecolor
                    $findings.Add([pscustomobject]@{
                        Mailbox     = $mailbox.DisplayName
                        UPN         = $mailbox.UserPrincipalName
                        FindingType = "Inbox rule redirect"
                        RuleName    = $rule.Name
                        Target      = ($rule.RedirectTo -join ', ')
                        Detail      = "Rule redirects message"
                    })
                }
            }
        } catch {
            if ($_.Exception.Message -ne "skip") {
                $inboxRuleErrorCount++
                Write-LogMessage "    Error retrieving rules for ${shortenedName}: $($_.Exception.Message)" $errormessagecolor
            }
        }

        ## Sweep rules set via OWA
        try {
            if (-not $canCheckSweepRules) { throw "skip" }
            $rules = Get-SweepRule -Mailbox $mailbox.UserPrincipalName -ErrorAction Stop
            foreach ($rule in $rules) {
                if (-not $rule.Enabled) { continue }
                $sweepRuleCount++
                Write-LogMessage "    Sweep rule '$($rule.Name)' enabled for $shortenedName" $errormessagecolor
                Write-LogMessage "    Source Folder = $($rule.SourceFolder)" $errormessagecolor
                Write-LogMessage "    Destination Folder = $($rule.DestinationFolder)" $errormessagecolor
                $findings.Add([pscustomobject]@{
                    Mailbox     = $mailbox.DisplayName
                    UPN         = $mailbox.UserPrincipalName
                    FindingType = "Sweep rule"
                    RuleName    = $rule.Name
                    Target      = "$($rule.SourceFolder) -> $($rule.DestinationFolder)"
                    Detail      = "Sweep rule enabled"
                })
            }
        } catch {
            if ($_.Exception.Message -ne "skip") {
                $sweepRuleErrorCount++
                Write-LogMessage "    Error retrieving sweep rules for ${shortenedName}: $($_.Exception.Message)" $errormessagecolor
            }
        }
    }

    Write-Progress -Activity "Checking mailboxes for forwards" -Completed
    Write-LogMessage "`nCheck mailbox forwards, inbox rules and sweep rules - Finish`n" $processmessagecolor

    Write-LogMessage "Summary" $systemmessagecolor
    Write-LogMessage "    Mailbox forwarding (forward and keep copy) = $mailboxForwardKeepCopyCount" $systemmessagecolor
    Write-LogMessage "    Mailbox forwarding (forward only, no copy) = $mailboxForwardOnlyCount" $systemmessagecolor
    Write-LogMessage "    Inbox rules with ForwardTo = $inboxForwardRuleCount" $systemmessagecolor
    Write-LogMessage "    Inbox rules with ForwardAsAttachmentTo = $inboxForwardAttachmentRuleCount" $systemmessagecolor
    Write-LogMessage "    Inbox rules with RedirectTo = $inboxRedirectRuleCount" $systemmessagecolor
    Write-LogMessage "    Inbox rule retrieval errors = $inboxRuleErrorCount" $warnmessagecolor
    Write-LogMessage "    Sweep rules enabled = $sweepRuleCount" $systemmessagecolor
    Write-LogMessage "    Sweep rule retrieval errors = $sweepRuleErrorCount" $warnmessagecolor
    Write-LogMessage "    Total findings = $($findings.Count)" $systemmessagecolor

    if ($CsvFile) {
        try {
            $csvDirectory = Split-Path -Path $CsvFile -Parent
            if ($csvDirectory -and -not (Test-Path -Path $csvDirectory)) {
                New-Item -Path $csvDirectory -ItemType Directory -Force | Out-Null
            }
            $findings | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8
            Write-LogMessage "`nFindings exported to $CsvFile" $processmessagecolor
        } catch {
            Write-LogMessage "`nFailed to export CSV to ${CsvFile}: $($_.Exception.Message)" $errormessagecolor
        }
    }

} catch {
    $fatalError = $true
    Write-LogMessage "An error occurred: $($_.Exception.Message)" $errormessagecolor
} finally {
    Write-LogMessage "Script complete" $systemmessagecolor
    Write-LogMessage "Log file: $script:LogFile`n" $processmessagecolor
}

if ($fatalError) { exit 1 }
