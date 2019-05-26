<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Dsplay a trace of all emails sent in recent hours to screen

Source - https://github.com/directorcia/Office365/blob/master/o365-msgtrace.ps1

Prerequisites = 1
1. Ensure connection to Exchange Online has already been completed

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warnmessagecolor = "yellow"
$hours = 48     ## Number of prior hours to check 

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

$dateEnd = get-date                         ## get current time
$dateStart = $dateEnd.AddHours(-$hours)     ## get current time less last $hours
$results = Get-MessageTrace -StartDate $dateStart -EndDate $dateEnd | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageID, MessageTraceID
$results | out-gridview

write-host -foregroundcolor $systemmessagecolor "Script completed`n"