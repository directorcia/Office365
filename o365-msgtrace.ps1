<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Display a trace of all emails sent in recent hours to screen

Source - https://github.com/directorcia/Office365/blob/master/o365-msgtrace.ps1

Prerequisites = 1
1. Ensure connection to Exchange Online has already been completed

More scripts available by joining http://www.ciaopspatron.com

#>

param(
    [int]$hours = 48,           ## Number of prior hours to check (max 240 = 10 days)
    [string]$ExportPath = ""    ## Optional CSV export path (leave blank to skip)
)

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warnmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

if ($hours -lt 1 -or $hours -gt 240) {
    write-host -foregroundcolor $errormessagecolor "Hours must be between 1 and 240 (Get-MessageTraceV2 limit is 10 days per query)`n"
    exit 1
}

$dateEnd = Get-Date
$dateStart = $dateEnd.AddHours(-$hours)

write-host -foregroundcolor $processmessagecolor "Retrieving message trace for last ${hours} hours ($($dateStart.ToString('yyyy-MM-dd HH:mm')) to $($dateEnd.ToString('yyyy-MM-dd HH:mm')))..."

## Get-MessageTraceV2 does not support page-number pagination.
## To retrieve subsequent pages, pass the last record's RecipientAddress and Received
## values as -StartingRecipientAddress and -EndDate respectively.
$allResults = [System.Collections.Generic.List[object]]::new()
$resultSize = 5000
$startingRecipient = $null
$currentEndDate = $dateEnd
$page = 1

try {
    do {
        write-host -foregroundcolor $processmessagecolor "  Fetching page ${page}..."

        $params = @{
            StartDate  = $dateStart
            EndDate    = $currentEndDate
            ResultSize = $resultSize
        }
        if ($startingRecipient) {
            $params['StartingRecipientAddress'] = $startingRecipient
        }

        $batch = Get-MessageTraceV2 @params |
            Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageID, MessageTraceID

        if ($batch) {
            $allResults.AddRange(@($batch))
            ## Prepare cursor values for next page
            $lastRecord = $batch | Select-Object -Last 1
            $startingRecipient = $lastRecord.RecipientAddress
            $currentEndDate = $lastRecord.Received
        }
        $page++
    } while ($batch -and $batch.Count -eq $resultSize)
}
catch {
    write-host -foregroundcolor $errormessagecolor "Error retrieving message trace: $($_.Exception.Message)`n"
    exit 1
}

$count = $allResults.Count
write-host -foregroundcolor $processmessagecolor "Retrieved ${count} message(s)`n"

if ($count -eq 0) {
    write-host -foregroundcolor $warnmessagecolor "No messages found in the specified time range`n"
}
else {
    ## Summary by status
    write-host -foregroundcolor $systemmessagecolor "Status breakdown:"
    $allResults | Group-Object Status | Sort-Object Count -Descending |
        ForEach-Object { write-host -foregroundcolor $processmessagecolor "  $($_.Name): $($_.Count)" }

    ## Top 5 senders
    write-host -foregroundcolor $systemmessagecolor "`nTop senders:"
    $allResults | Group-Object SenderAddress | Sort-Object Count -Descending | Select-Object -First 5 |
        ForEach-Object { write-host -foregroundcolor $processmessagecolor "  $($_.Name): $($_.Count)" }

    ## Top 5 recipients
    write-host -foregroundcolor $systemmessagecolor "`nTop recipients:"
    $allResults | Group-Object RecipientAddress | Sort-Object Count -Descending | Select-Object -First 5 |
        ForEach-Object { write-host -foregroundcolor $processmessagecolor "  $($_.Name): $($_.Count)" }

    write-host ""

    if ($ExportPath) {
        try {
            $allResults | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            write-host -foregroundcolor $processmessagecolor "Results exported to: $ExportPath`n"
        }
        catch {
            write-host -foregroundcolor $errormessagecolor "Failed to export CSV: $($_.Exception.Message)`n"
        }
    }

    $allResults | Out-GridView -Title "Message Trace - Last ${hours} hours (${count} messages)"
}

write-host -foregroundcolor $systemmessagecolor "Script completed`n"