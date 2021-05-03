param(                        
    [switch]$debug = $false    ## if -debug parameter don't prompt for input
)<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Report on Azure Sentinel rules
Source - https://github.com/directorcia/Office365/blob/master/az-sentinel-ruleget.ps1

Prerequisites = 2
1. Azure AZ.SecurityInsights Module installed
2. Connect to Azure tenant - https://github.com/directorcia/office365/blob/master/az-connect-si.ps1

More scripts available at www.ciaopspatron.com

#>

Function Get-LocalTime($UTCTime)
{
$strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
$LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
Return ($LocalTime)
}

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

if ($debug) {
    write-host "Script activity logged at ..\az-sentinel-ruleget.txt"
    start-transcript "..\az-sentinel-ruleget.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

Clear-Host
write-host -foregroundcolor cyan -backgroundcolor DarkBlue ">>>>>> Copyright www.ciaops.com <<<<<<`n"
write-host "--- Report on Azure Sentinel rules ---"

write-host -foregroundcolor $processmessagecolor "`nCheck for AZ.Securityinsights PowerShell module"
if (get-module -listavailable -name Az.Securityinsights) {
    ## Has the AZ PowerShell module been loaded?
    write-host -foregroundcolor yellow -backgroundcolor darkGreen "Azure AZ.SecurityInsights PowerShell Module found"
}
else {
    write-host -foregroundcolor yellow -backgroundcolor red "Azure AZ.Securityinsights PowerShell Module not installed. Please install and re-run script`n"
    write-host "You can install the Azure AZ.Securityinsights Powershell module by:`n"
    write-host "    1. Launching an elevated Powershell console then,"
    write-host "    2. Running the command,'install-module -name az.securityinsights -allowclobber'.`n"
    Write-Host -ForegroundColor $systemmessagecolor "`nScript Finished"
    if ($debug) {
        Stop-Transcript | Out-Null
    }
    Pause                                                                               ## Pause to view error on screen
    exit 0                                                                              ## Terminate script 
}
Try {
    import-module -name Az.Securityinsights | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[001] - Failed to import AZ.Securityinsights module - ", $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null
    }
    exit 1
}

write-host -ForegroundColor $processmessagecolor "Getting Workspaces in tenant"
$ws = Get-AzOperationalInsightsWorkspace |  select-object name, resourcegroupname,tags | Out-GridView -PassThru -title "Select the Workspace to use"
write-host -ForegroundColor $processmessagecolor "Getting all available valid rule templates"
$templatestotalraw = Get-AzSentinelAlertRuleTemplate -ResourceGroupName $ws.resourcegroupname -WorkspaceName $ws.Name
write-host -ForegroundColor $processmessagecolor "Getting all available valid rules in use"
$templatesinuseraw = get-azsentinelalertrule -ResourceGroupName $ws.resourcegroupname -WorkspaceName $ws.Name
write-host -ForegroundColor Gray -backgroundcolor blue "`n--- All available valid rule templates in date order ---"
$templatestotal = @()
foreach ($template in $templatestotalraw) {
    if ($template.kind -ne "Error") {
        $templatestotal +=[PSCustomObject]@{
            'Name' = $template.displayname;
            'Created' = (get-localtime($template.createddateutc)).tostring("dd-MMM-yyyy")
        }
    }
}
$templatestotal | Select-Object Name,Created | sort-object Created -Descending | format-table
write-host -ForegroundColor Gray -backgroundcolor blue "--- All valid rules currently in use in date order ---"
$templatesinuse = @()
foreach ($template in $templatesinuseraw) {
    if ($template.kind -ne "Error") {
        $templatesinuse +=[PSCustomObject]@{
            'Name' = $template.displayname;
            'LastModified' = (get-localtime($template.lastmodifiedutc)).tostring("dd-MMM-yyyy")
        }
    }
}
Write-Output $templatesinuse |  Select-Object Name,LastModified | sort-object LastModified -Descending | format-table

write-host -ForegroundColor Gray -backgroundcolor blue "`n--- Total valid rule status ---`n"
$templatesnotinuseraw = @()
foreach ($template in $templatestotalraw) {
    if ($templatesinuseraw.displayname -match $template.displayname ) {
        write-host -ForegroundColor $processmessagecolor "[Active]",$template.displayname
    } else {
        write-host -ForegroundColor $warningmessagecolor "[Inactive]",$template.displayname
        $templatesnotinuseraw += $template
    }
}
write-host -ForegroundColor Gray -backgroundcolor blue "`n--- Valid rules not in use in date order ---"
$templatesnotinuse = @()
foreach ($template in $templatesnotinuseraw) {
    if ($template.kind -ne "Error") {
        $templatesnotinuse +=[PSCustomObject]@{
            'Name' = $template.displayname;
            'Created' = (get-localtime($template.createddateutc)).tostring("dd-MMM-yyyy")
        }
    }
}
$templatesnotinuse |  Select-Object Name,Created | sort-object Created -Descending | format-table
write-host -ForegroundColor Gray -backgroundcolor blue "`n--- Sentinel Rule Summary ---`n"
write-host "Total templates =",$templatestotalraw.Count
write-host "Total templates without errors =",($templatestotalraw | where-object {$_.kind -ne "Error"}).Count
write-host "Newest template date =",($templatestotal | sort-object Created -Descending | select-object -first 1).Created
write-host "`nTotal templates in use =",$templatesinuseraw.Count
write-host "Total templates in use without errors =",($templatesinuseraw | where-object {$_.kind -ne "Error"}).Count
write-host "`nTotal templates not in use =",$templatesnotinuseraw.Count
write-host "Total templates not in use without errors =",($templatesnotinuseraw | where-object {$_.kind -ne "Error"}).Count

write-host -foregroundcolor $systemmessagecolor "`nScript Completed`n"
if ($debug) {
    Stop-Transcript | Out-Null
}