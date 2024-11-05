param(                           ## if no parameter used then login without MFA and use interactive mode
    [switch]$prompt = $false     ## if -prompt parameter used prompt user for input
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Get all items from tenant secure score

Source - https://github.com/directorcia/Office365/blob/master/mggraph-ssdescpt-get.ps1

Prerequisites = 1
1. Install MSGraph module - https://www.powershellgallery.com/packages/Microsoft.Graph/

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host
if ($prompt) {
    write-host -foregroundcolor $processmessagecolor "Script activity logged at ..\mggraph-ssdescpt-get.txt"
    start-transcript "..\mggraph-ssdescpt-get.txt"
}

write-host -foregroundcolor $systemmessagecolor "Script started`n"

# Check if the Microsoft Graph module is installed
try {
    $context = get-mgcontext -ErrorAction Stop
}
catch {
    write-host -foregroundcolor $errormessagecolor "Not connected to Microsoft Graph. Please connect to Microsoft Graph first using connect-mggraph`n"
    if ($prompt) {stop-transcript}
    exit
}
if (-not $context) {
    write-host -foregroundcolor $errormessagecolor "Not connected to Microsoft Graph. Please connect to Microsoft Graph first using connect-mggraph`n"
    if ($prompt) {stop-transcript}
    exit
}

write-host -foregroundcolor $processmessagecolor "Connected to Microsoft Graph"
write-host "  - Connected account =",$context.Account,"`n"
if ($prompt) { pause }

# Specify the URI to call and method
$uri = "https://graph.microsoft.com/beta/security/securescores"
$method = "GET"

write-host -foregroundcolor $processmessagecolor "Run Graph API Query"
# Run Graph API query 
$query = Invoke-MgGraphRequest -Uri $URI -method $method -ErrorAction Stop

$names = $query.value[0].controlscores          # get the most current secure score results

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
    if ($prompt) { pause }
}
write-host -foregroundcolor $systemmessagecolor "`nScript Completed`n"

if ($prompt) {stop-transcript}