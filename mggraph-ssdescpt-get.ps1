<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Get all items from tenant secure score

Source - https://github.com/directorcia/Office365/blob/master/mggraph-ssdescpt-get.ps1

Prerequisites = 1
1. Install MSGRAph module - https://www.powershellgallery.com/packages/Microsoft.Graph/

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

start-transcript "..\mggraph-ssdescpt-get $(get-date -f yyyyMMddHHmmss).txt"       ## write output file to parent directory

write-host -foregroundcolor $systemmessagecolor "Script started`n"

# Specify the URI to call and method
$uri = "https://graph.microsoft.com/beta/security/securescores"
$method = "GET"

write-host -foregroundcolor $processmessagecolor "Run Graph API Query"
# Run Graph API query 
$query = Invoke-MgGraphRequest -Uri $URI -method GET -ErrorAction Stop

$names = $query.value[0].controlscores

$item = 0
write-host -foregroundcolor $processmessagecolor "Display results`n"
foreach ($control in $names) {
    $item++
    write-host -foregroundcolor green -BackgroundColor Black "`n*** Item", $item, "***"
    write-host "Control Category     : ", $control.controlCategory
    write-host "Control Name         : ", $control.controlName
    write-host "Control Score        : ", $control.Score
    write-host "Control Description  : ", $control.Description
    write-host "Control On           : ", $control.on
    write-host "Implementation status: ", $control.implementationstatus
    write-host "Score in percentage  : ", $control.scoreinpercentage
    write-host "Last synced          : ", $control.lastsynced
    write-host "`n"
}
write-host -foregroundcolor $systemmessagecolor "`nScript Completed`n"

stop-transcript