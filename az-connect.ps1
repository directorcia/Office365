<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Script designed to login to Azure resources
Source - https://github.com/directorcia/office365/blob/master/az-connect.ps1

Prerequisites = 1
1. Ensure az module installed or updated

ensure that install-az msonline has been run
ensure that update-az msonline has been run to get latest module

Allow custom scripts to run just for this instance
set-executionpolicy -executionpolicy bypass -scope currentuser -force
#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warnmessagecolor = "yellow"
$version = "2.00"

start-transcript "..\az-connect.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run

Clear-host

write-host -foregroundcolor $systemmessagecolor "Script started. Version = $version`n"
write-host -foregroundcolor cyan -backgroundcolor DarkBlue ">>>>>> Created by www.ciaops.com <<<<<<`n"
write-host "--- Script to connect to an Azure tenant ---`n"

if (get-module -listavailable -name az.accounts) {    ## Has the Azure PowerShell module been loaded?
    write-host -ForegroundColor $processmessagecolor "Azure PowerShell module found"
}
else {              ## If Azure PowerShell module not found
    write-host -ForegroundColor $warnmessagecolor -backgroundcolor $errormessagecolor "`n[001] - Azure PowerShell module not installed. Please install and re-run script`n"
    write-host -ForegroundColor $warnmessagecolor -backgroundcolor $errormessagecolor "Exception message:",$_.Exception.Message,"`n"
    Stop-Transcript                 ## Terminate transcription
    exit 1                          ## Terminate script
}

write-host -ForegroundColor $processmessagecolor "[Start] - Connect to Azure tenant"
try {
    connect-AzAccount -warningaction "SilentlyContinue" | Out-Null
}
catch {
    write-host -ForegroundColor $warnmessagecolor -backgroundcolor $errormessagecolor "`n[002] - Unable to connect to Azure tenant`n"
    write-host -ForegroundColor $warnmessagecolor -backgroundcolor $errormessagecolor "Exception message:",$_.Exception.Message,"`n"
    Stop-Transcript                 ## Terminate transcription
    exit 2                          ## Terminate script  
}
## Select desired Azure subscription from list of subscriptions
Get-AzSubscription -warningaction "SilentlyContinue" | Out-GridView -PassThru -title "Select the Azure subscription to use" | Select-AzSubscription | Out-Null
write-host -ForegroundColor $processmessagecolor "[Finish] - Connect to Azure tenant"

write-host -foregroundcolor $systemmessagecolor "`nScript finished`n"