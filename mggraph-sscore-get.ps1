param(                         
    [switch]$debug = $false     ## if -debug parameter capture log information
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Connect to the Microsoft Graph and retrieve all the items in Secure Score for the last 90 days
Source - https://github.com/directorcia/Office365/blob/master/mggraph-sscore-get.ps1

Prerequisites = 1
1. Microsoft Graph module installed - https://www.powershellgallery.com/packages/Microsoft.Graph/

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor ="red"
$warningmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host
write-host -foregroundcolor $systemmessagecolor "Secure Score report script started`n"
if ($debug) {
    write-host -foregroundcolor $processmessagecolor "Script activity logged at ..\mggraph-sscore-get.txt"
    start-transcript "..\mggraph-sscore-get.txt"
}

<#  ----- [Start] Graph PowerShell module check -----   #>
if (get-module -listavailable -name Microsoft.Graph.Authentication) {    ## Has the Graph import module been installed?
    write-host -ForegroundColor $processmessagecolor "Graph authentication module found"
}
else {
    write-host -ForegroundColor $warningmessagecolor -backgroundcolor $errormessagecolor "[001] - Graph PowerShell module not installed. Please install and re-run script - ", $_.Exception.Message
    if ($debug) {
        Stop-Transcript                 ## Terminate transcription
    }
    exit 1                              ## Terminate script
}

$scopes = "SecurityEvents.Read.All"
write-host -foregroundcolor $processmessagecolor "Settings Scopes = $scopes"
write-host -foregroundcolor $processmessagecolor "Connect to Microsoft Graph"
Connect-MgGraph -scopes $scopes -NoWelcome | Out-Null
$graphcontext = Get-MgContext
write-host "  - Connected account =",$graphcontext.Account
write-host "  - Tenant ID =", $graphcontext.TenantId

write-host -foregroundcolor $processmessagecolor "Make Graph Request"
$uri = "https://graph.microsoft.com/beta/security/securescores"
$request = Invoke-MgGraphRequest -Uri $URI -method GET -ErrorAction Stop

write-host "`nSecure score for last 90 days"
write-host "-----------------------------`n"
foreach ($item in $request.value) {
    $sspercent=($item.currentscore/$item.maxscore)
    $formattedDate = $item.createdDateTime.ToString("dd-MM-yyyy")
    write-host -foregroundcolor white -BackgroundColor Blue "$formattedDate Score =",$item.currentscore, "of",$item.maxscore,"["$sspercent.tostring("P")"]`n"
}
Write-Host -ForegroundColor $systemmessagecolor "Script Finished`n"
if ($debug) {
    Stop-Transcript | Out-Null   
}
