param(                         ## if no parameters used then don't output to CSV
    [switch]$csv = $false      ## if -csv parameter used then write to CSV to parent directory
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description
Script designed to check and report the status of mailboxes in the tenant

Source - https://github.com/directorcia/Office365/blob/master/o365-mx-check.ps1

Notes - You can extend many of the default limits at no additional cost

Prerequisites = 1
1. Connected to Exchange Online. Recommended script = https://github.com/directorcia/Office365/blob/master/o365-connect-exov2.ps1

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$auditlogagelimitdefault = 90                   ## default days for mailbox audit log. You can extend beyond this for free
$retaindeleteditemsmax = 30              ## default days for deleted items retention. You can extend beyond this for free
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$version = "2.00"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

start-transcript "..\o365-mx-check.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run

Clear-host

write-host -foregroundcolor $systemmessagecolor "Script started. Version = $version`n"
write-host -foregroundcolor cyan -backgroundcolor DarkBlue ">>>>>> Created by www.ciaops.com <<<<<<`n"
write-host "--- Script to display mailbox settings ---`n"

if ((get-module -listavailable -name ExchangeOnlineManagement) -or (get-module -listavailable -name msonline)) {    ## Has the Exchange Online PowerShell module been loaded?
    write-host -ForegroundColor $processmessagecolor "Exchange Online PowerShell found"
}
else {              ## If Exchange Online PowerShell module not found
    write-host -ForegroundColor yellow -backgroundcolor Red "`n[001] - Exchange Online PowerShell module not installed. Please install and re-run script`n"
    write-host -ForegroundColor yellow -backgroundcolor red "Exception message:",$_.Exception.Message,"`n"
    Stop-Transcript                 ## Terminate transcription
    exit 1                          ## Terminate script
}

write-host -ForegroundColor $processmessagecolor "Getting Mailboxes"

try {
    $mailboxes=get-mailbox -ResultSize unlimited
}
catch {
    write-host -ForegroundColor yellow -backgroundcolor Red "`n[002] - Exchange Online PowerShell module not installed. Please install and re-run script`n"
    write-host -ForegroundColor yellow -backgroundcolor red "Exception message:",$_.Exception.Message,"`n"
    Stop-Transcript                 ## Terminate transcription
    exit 2                          ## Terminate script  
}

write-host -ForegroundColor $processmessagecolor "Start checking mailboxes`n"

$results = @()
foreach ($mailbox in $mailboxes) {
    <#          ----- Truncate Long email addresses ----- #>
    if ($mailbox.userprincipalname.length -gt 60) {
        $upn = $mailbox.userprincipalname.substring(0,60)+"..."
    }
    else {
        $upn = $mailbox.userprincipalname
    }
    if ($mailbox.PrimarySMTPAddress.length -gt 60) {
        $primsmtp = $mailbox.PrimarySMTPAddress.substring(0,60)+"..."
    }
    else {
        $primsmtp = $mailbox.PrimarySMTPAddress
    }
    write-host -foregroundcolor yellow -BackgroundColor Black "Mailbox =",$mailbox.displayname,"[",$upn,"]"
    write-host -foregroundcolor Gray "  Primary SMTP address =",$primsmtp
    write-host -foregroundcolor Gray "  Created =",$mailbox.whencreated

## Mailboxes should have auditing enabled

    if ($mailbox.auditenabled) {
        write-host -foregroundcolor $processmessagecolor "  Audit enabled =",$mailbox.AuditEnabled
    } else {
        write-host -foregroundcolor $errormessagecolor "  Audit enabled =",$mailbox.AuditEnabled
    }

## Audit log limit for mailboxes should be extended from default

    if ([timespan]::parse($mailbox.auditlogagelimit).days -gt $auditlogagelimitdefault) {
        write-host -foregroundcolor $processmessagecolor "  Audit log limit (days) =",($mailbox.Auditlogagelimit).split('.')[0]
    } else {
        write-host -foregroundcolor $errormessagecolor "  Audit log limit (days) =",($mailbox.Auditlogagelimit).split('.')[0]
    }

## Mailboxes should have their retained deleted item retention period extended to 30 days

    if ([timespan]::parse($mailbox.retaindeleteditemsfor).days -ge $retaindeleteditemsmax) {
        write-host -foregroundcolor $processmessagecolor "  Retain Deleted items for (days) =",($mailbox.retaindeleteditemsfor).split('.')[0]
    } else {
        write-host -foregroundcolor $errormessagecolor "  Retain Deleted items for (days) =",($mailbox.retaindeleteditemsfor).split('.')[0]
    }

## Mailboxes should not be forwarding to other email addresses or mailboxes

    if (-not [string]::IsNullOrEmpty($mailbox.forwardingaddress)){
        write-host -foregroundcolor $errormessagecolor "  Forwarding address =",$mailbox.forwardingaddress
    }
    if (-not [string]::IsNullOrEmpty($mailbox.forwardingsmtpaddress)){
        write-host -foregroundcolor $errormessagecolor "  Forwarding SMTP address =",$mailbox.forwardingsmtpaddress
    }

## Mailboxes should have litigation hold enabled

    if ($mailbox.LitigationHoldEnabled) {
        write-host -foregroundcolor $processmessagecolor "  Litigation hold =",$mailbox.litigationholdenabled
    } else {
        write-host -foregroundcolor $errormessagecolor "  Litigation hold =",$mailbox.litigationholdenabled
    }

## Mailboxes should have archive enabled

    if ($mailbox.archivestatus -eq "active") {
        Write-host -foregroundcolor $processmessagecolor "  Archive status =",$mailbox.archivestatus
    } else {
        Write-host -foregroundcolor $errormessagecolor "  Archive status =",$mailbox.archivestatus
    }

## Report mailbox maximum send and receive sizes

    Write-host -foregroundcolor Gray "  Max send size =",$mailbox.maxsendsize
    write-host -foregroundcolor Gray "  Max receive size =",$mailbox.maxreceivesize
    try {               ## See if further details about mailbox are available
        $extramailbox=get-casmailbox -Identity $mailbox.userprincipalname -erroraction stop
    }
    catch {
        $extramailbox = $null
        $reason = $_.Exception.Message
    }    

## Mailboxes should not have POP3 enabled

    if (-not [string]::IsNullOrEmpty($extramailbox)){
        if (-not $extramailbox.popenabled) {
            write-host -foregroundcolor $processmessagecolor "  POP3 enabled =",$extramailbox.popenabled
        } else {
            write-host -foregroundcolor $errormessagecolor "  POP3 enabled =",$extramailbox.popenabled
        }

    ## Mailboxes should not have IMAP enabled

        if (-not $extramailbox.ImapEnabled) {
            write-host -foregroundcolor $processmessagecolor "  IMAP enabled =",$extramailbox.imapenabled
        } else {
            write-host -foregroundcolor $errormessagecolor "  IMAP enabled =",$extramailbox.imapenabled
        }
    }
    else {
        write-host -ForegroundColor $errormessagecolor "  Further mailbox details not found -",$reason.substring(0,60),"..."
    }
    $results+=[PSCustomObject]@{
        Displayname = $mailbox.Displayname;
        Userprincipalname = $mailbox.userprincipalname;
        PrimarySMTPAddress = $mailbox.PrimarySMTPAddress;
        WhenCreated = $mailbox.whencreated;
        AuditEnabled = $mailbox.AuditEnabled;
        Auditlogagelimit = $mailbox.Auditlogagelimit;
        Retaindeleteditemsfor = $mailbox.Retaindeleteditemsfor;
        Forwardingaddress = $mailbox.forwardingaddress;
        Forwardingsmtpaddress = $mailbox.forwardingsmtpaddress;
        LitigationHoldEnabled = $mailbox.LitigationHoldEnabled;
        Archivestatus = $mailbox.Archivestatus;
        PopEnabled = $extramailbox.Popenabled;
        IMAPEnabled = $extramailbox.IMAPEnabled;
        MaxSendSize = $mailbox.MaxSendSize;
        MaxReceiveSize = $mailbox.MaxReceiveSize
    }
    write-host 
}
if ($csv) {                                                     ## If CSV paramter set
    write-host -foregroundcolor $processmessagecolor "Writing all output to file o365-mx-check$(get-date -f yyyyMMddHHmmss).csv in parent directory" 
    $results | export-csv -path "..\o365-mx-check$(get-date -f yyyyMMddHHmmss).csv" -NoTypeInformation ## Export array results to CSV file
}
else {                                                          ## If CSV parameter not set
    write-host -foregroundcolor $processmessagecolor "No CSV created"
}
write-host -ForegroundColor $processmessagecolor "Finish checking mailboxes"
write-host
write-host -foregroundcolor $systemmessagecolor "Script completed`n"