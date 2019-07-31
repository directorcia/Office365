<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/o365-spo-getusage.ps1

Description - Show SharePoint and ODFB site storage usage from largest to smallest

Prerequisites = 1
1. Ensure SharePoint online PowerShell module installed or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$highlightmessagecolor = "yellow"
$sectionmessagecolor = "white"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

$sposites=get-sposite -IncludePersonalSite $false -limit all | Sort-Object StorageUsageCurrent -Descending          ## get all non-ODFB sites
Write-host -foregroundcolor $sectionmessagecolor "*** Current SharePoint Site Usage ***`n"
foreach ($sposite in $sposites) {                           ## loop through all of these sites
    $mbsize=$sposite.StorageUsageCurrent                    ## save total size to a variable to be formatted later
    write-host -foregroundcolor $highlightmessagecolor $sposite.title,"=",$mbsize.tostring('N0'),"MB"
    write-host -foregroundcolor $processmessagecolor $sposite.url
    write-host
}
$sposites=get-sposite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/" | Sort-Object StorageUsageCurrent -Descending
Write-host -foregroundcolor $sectionmessagecolor "*** Current ODFB Site Usage ***`n"
foreach ($sposite in $sposites) {
    $mbsize=$sposite.StorageUsageCurrent
    write-host -foregroundcolor $highlightmessagecolor $sposite.title,"=",$mbsize.tostring('N0'),"MB"
    write-host -foregroundcolor $processmessagecolor $sposite.url
    write-host
}
write-host -foregroundcolor $systemmessagecolor "Script completed`n"