<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on user logins from Office 365 Audit logs

Source - https://github.com/directorcia/Office365/blob/master/o365-login-audit.ps1

Prerequisites = 1
1. Connected to Exchange Online

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$startdate = "9/5/2018"     ## Starting date for audit log search MM/DD/YYYY
$enddate = "9/11/2018"      ## Ending date for audit log search MM/DD/YYYY
$sesid="0"                  ## change this if you want to re-reun the script multiple times in a single session
$Results = @()              ## where the ultimate results end up
$strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName           ## determine current local timezone
$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)    ## for Timezone calculations
$AuditOutput = 1            ## Set variable value to trigger loop below (can be anything)
$convertedoutput=""
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

write-host -foregroundcolor $systemmessagecolor "Script started`n"

# Loop will run until $AuditOutput returns null which equals that no more event objects exists from the specified date
while ($AuditOutput) {
    # Search the defined date(s), SessionId + SessionCommand in combination with the loop will return and append 100 object per iteration until all objects are returned (minimum limit is 50k objects)
    write-host -foregroundcolor $processmessagecolor "Searching Audit logs. Please wait"
    $AuditOutput = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -recordtype $recordtype -operations $operation -SessionId $sesid -SessionCommand ReturnLargeSet
    # Select and expand the nested object (AuditData) as it holds relevant reporting data. Convert output format from default JSON to enable export to csv
    $ConvertedOutput = $AuditOutput | Select-Object -ExpandProperty AuditData | ConvertFrom-Json
    # Export results exluding type information. Append rather than overwrite if the file exist in destination folder
    ## $ConvertedOutput | Select-Object CreationTime,UserId,Operation,ClientIP 
    foreach ($Entry in $convertedoutput)
    {  
        $return = "" | select-object Creationtime,Localtime,UserId,Operation,clientip
        $data = $Entry | Select-Object Creationtime,userid,operation,clientip
        $return.Creationtime = $data.CreationTime
        $return.localtime = [System.TimeZoneInfo]::ConvertTimeFromUtc($data.Creationtime, $TZ)
        $return.clientip = $data.ClientIP
        $return.UserId = $data.UserId
        #Obtain the UserAgent string from inside the
        $return.Operation = $data.Operation
        #Returns the data to outside of the loop
        $Results += $return
    }
}
write-host
write-host -foregroundcolor white "Local Time            Client IP       Operation"
write-host -foregroundcolor White "----------            ---------       ---------"
foreach ($result in $results){
    if ($result.operation -eq "userloginfailed"){
        ## Report failed logins in red
        write-host -ForegroundColor red $result.localtime,$result.clientip,$result.operation,$result.userid
    } else {
        ## Report successful logins in green
        write-host -foregroundcolor green $result.localtime,$result.clientip,$result.operation,$result.userid
    }
}

write-host
write-host -foregroundcolor $systemmessagecolor "Script Completed`n"
