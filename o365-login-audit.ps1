param(                         ## if no parameters used then get logs for 1 days, output successful and unsuccessful logins and don't output to CSV
    [switch]$days = $false,    ## if -days parameter prompt for number of days
    [switch]$fail = $false,    ## if -fail parameter only show failed logins
    [switch]$csv = $false      ## if -csv parameter used then write to CSV to parent directory
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on user logins from Office 365 Unified Audit logs

Notes - 
1. That the Office 365 Unified audit los are NOT immediate. Information may take a while to actually end up in there.
2. That the unified logs generally only record 'interactive' logins not app token refresh. This may explain why you see more login entries in Azure AD Signs reports

Source - https://github.com/directorcia/Office365/blob/master/o365-login-audit.ps1

Prerequisites = 1
1. Connected to Exchange Online - Recommended script = https://github.com/directorcia/Office365/blob/master/o365-connect-exov2.ps1
2. Need to have Unified Audit logs enabled prior - https://blog.ciaops.com/2018/02/01/enable-activity-auditing-in-office-365/

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$version="2.00"
$sesid = get-random                                     ## get random session number 
$Results = @()                                          ## initialise array                 
$displays = @()                                         ## initailise array
$AuditOutput = @()                                      ## initialise array
$currentoutput = @()                                    ## initialise array
$strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName           ## determine current local timezone
$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)    ## for Timezone calculations
## Valid record types = 
## AzureActiveDirectory, AzureActiveDirectoryAccountLogon,AzureActiveDirectoryStsLogon, ComplianceDLPExchange
## ComplianceDLPSharePoint, CRM, DataCenterSecurityCmdlet, Discovery, ExchangeAdmin, ExchangeAggregatedOperation
## ExchangeItem, ExchangeItemGroup, MicrosoftTeams, MicrosoftTeamsAddOns, MicrosoftTeamsSettingsOperation, OneDrive
## PowerBIAudit, SecurityComplianceCenterEOPCmdlet, SharePoint, SharePointFileOperation, SharePointSharingOperation
## SkypeForBusinessCmdlets, SkypeForBusinessPSTNUsage, SkypeForBusinessUsersBlocked, Sway, ThreatIntelligence, Yammer
$recordtype = "azureactivedirectorystslogon"

## Office 365 Management Activity API schema
## Valid record types = https://docs.microsoft.com/en-us/office365/securitycompliance/search-the-audit-log-in-security-and-compliance?redirectSourcePath=%252farticle%252f0d4d0f35-390b-4518-800e-0c7ec95e946c#audited-activities
## Operation types = "<value1>","<value2>","<value3>"
$operation="userloginfailed","userloggedin" ## use this line to report all logins
##$operation="userloginfailed" ## use this line to report failed logins
##$operation="userloggedin" ## use this line to report successful logins

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

start-transcript "..\o365-login-audit.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run

write-host -foregroundcolor $systemmessagecolor "Script started. Version = $version`n"
write-host -foregroundcolor cyan -backgroundcolor DarkBlue ">>>>>> Created by www.ciaops.com <<<<<<`n"
write-host "--- Script to display user logins from Unified Audit log ---`n"

write-host -foregroundcolor $processmessagecolor "Calculate day range"
if (-not $days) {                                           ## If days parameter not specified 
    $numberdays = 1                                         ## Set search to one day
}
else {                                                      ## If days parameter specified
    do{
        $numberdays = read-host -Prompt "`nEnter total number of previous days to check from now"       ## Prompt for number of days to check
    } Until ((-not [string]::isnullorempty($numberdays)) -and ($numberdays -match "^\d+$"))             ## Keep prompting until not blank and numeric
    Write-Host
}
$numberdaysint = [int]::Parse($numberdays)                          ## Convert string to integer
$startdatelocal = (get-date).adddays(-$numberdaysint)               ## Starting date for audit log search Local. Default = 1 day ago
$startdate = $startdatelocal.touniversaltime()                      ## Convert local start time to UTC
$enddate = (get-date).touniversaltime()                             ## Ending date for audit log search UTC. Default = current time
$diff = New-TimeSpan -Start $startdate -End $enddate                ## Determine the difference between start and finish dates
write-host -ForegroundColor $processmessagecolor "Total number of previous days to check from now:",([int]$diff.TotalDays)

if ((get-module -listavailable -name ExchangeOnlineManagement) -or (get-module -listavailable -name msonline)) {    ## Has the Exchange Online PowerShell module been loaded?
    write-host -ForegroundColor $processmessagecolor "Exchange Online PowerShell found"
}
else {              ## If Exchange Online PowerShell module not found
    write-host -ForegroundColor yellow -backgroundcolor Red "`n[001] - Exchange Online PowerShell module not installed. Please install and re-run script`n"
    write-host -ForegroundColor yellow -backgroundcolor red "Exception message:",$_.Exception.Message,"`n"
    Stop-Transcript                 ## Terminate transcription
    exit 1                          ## Terminate script
}

# Search the defined date(s), SessionId + SessionCommand in combination with the loop will return and append 5000 object per iteration until all objects are returned (minimum limit is 50k objects)
$count = 1
do {
    write-host -foregroundcolor $processmessagecolor "Getting unified audit logs page $count - Please wait"
    try {
        $currentOutput = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -recordtype $recordtype -operations $operation -SessionId $sesid -SessionCommand ReturnLargeSet -resultsize 5000
    }
    catch {
        write-host -ForegroundColor yellow -backgroundcolor red "`n[002] - Search Unified Log error. Typically not connected to Exchange Online. Please connect and re-run script`n"
        write-host -ForegroundColor yellow -backgroundcolor red "Exception message:",$_.Exception.Message,"`n"
        Stop-Transcript                 ## Terminate transcription
        exit 2                          ## Terminate script     
    }
    $AuditOutput += $currentoutput      ## Build total results array
    ++$count                            ## Increament page count
} until ($currentoutput.count -eq 0)    ## Until there are no more logs in range to get
    
# Select and expand the nested object (AuditData) as it holds relevant reporting data. Convert output format from default JSON to enable export to csv
$ConvertedOutput = $AuditOutput | Select-Object -ExpandProperty AuditData | sort-object creationtime |  ConvertFrom-Json

foreach ($Entry in $convertedoutput)            ## Loop through all result entries
{  
    $return = "" | select-object Creationtime,Localtime,UserId,Operation,ClientIP
    $return.Creationtime = $Entry.CreationTime
    $return.localtime = [System.TimeZoneInfo]::ConvertTimeFromUtc($Entry.Creationtime, $TZ)         ## Convert entry to local time
    $return.clientip = $Entry.ClientIP
    $return.UserId = $Entry.UserId               
    $return.Operation = $Entry.Operation         
    $Results += $return                         ## Build results array
}

$displays = $results | sort-object -descending localtime        ## Sort result array in reverse chronological order
if ($csv) {                                                     ## If CSV paramter set
    write-host -foregroundcolor $processmessagecolor "Writing all output to file o365-login-audit$(get-date -f yyyyMMddHHmmss).csv in parent directory"
    $displays | Select-Object LocalTime, ClientIP, Operation, UserId | export-csv -path "..\o365-login-audit$(get-date -f yyyyMMddHHmmss).csv" -NoTypeInformation ## Export array results to CSV file
}
else {                                                          ## If CSV parameter not set
    write-host -foregroundcolor $processmessagecolor "No CSV created"
}
write-host
write-host -foregroundcolor white "Local Time`t`t Client IP`t`t Operation`t`t Login"            ## Merely an indication of the headings
write-host -foregroundcolor White "----------`t`t ---------`t`t ---------`t`t -----"            ## Not possible to align for every run option
foreach ($display in $displays){
    if (($display.clientip).length -lt 14) {        ## Determine total lenght of first field
        $gap = "`t`t"                               ## If a shorter field add two tabs in output
    }
    else {
        $gap = "`t"
    }
    if ($display.operation -eq "userloginfailed"){          ## Report failed logins in red
        write-host -ForegroundColor red $display.localtime,"`t",$display.clientip,$gap,$display.operation,"`t",$display.userid
    } elseif (-not $fail) {                                                ## Report successful logins in $processmessagecolor
        write-host -foregroundcolor $processmessagecolor $display.localtime,"`t",$display.clientip,$gap,$display.operation,"`t`t",$display.userid
    }
}
write-host
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"
Stop-Transcript | Out-Null          ## End transscript
