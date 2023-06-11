param(                        
    [switch]$debug = $false,     ## if -debug parameter don't prompt for input
    [switch]$noupdate = $false,   ## if -noupdate used then module will not be checked for more recent version
    [switch]$noprompt = $false   ## if -noprompt parameter used prompt user for input
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description
Script designed to check and report the status of mailboxes in the tenant

Source - https://github.com/directorcia/Office365/blob/master/c.ps1

Notes - 

Prerequisites - 0

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

if ($debug) {
    write-host "Script activity logged at ..\c.txt"
    start-transcript "..\c.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

Clear-host

write-host -foregroundcolor $systemmessagecolor "Microsoft Cloud connections menu script started"
write-host -foregroundcolor cyan -backgroundcolor DarkBlue ">>>>>> Created by www.ciaops.com <<<<<<`n"
write-host "--- Script to connect to Microsoft Cloud services ---`n"
if (-not $debug) {
    Write-host -foregroundcolor $warningmessagecolor "    * use the -debug parameter on the command line to create an execution log file for this script"
}
if (-not $noupdate) {
    write-host -foregroundcolor $warningmessagecolor  "    * use the -noupdate parameter on the command line to prevent checking for latest module version"
}
if (-not $noprompt) {
    write-host -foregroundcolor $warningmessagecolor  "    * use the -noprompt parameter on the command line present no prompts"
}
#Region Modules
$scripts = @()
$scripts += [PSCustomObject]@{
    Name = "o365-connect-tms.ps1";
    Service = "Teams";
    Module = "MicrosoftTeams"    
}
$scripts += [PSCustomObject]@{
    Name = "o365-connect-spo.ps1";
    Service = "SharePoint Online"; 
    Module = "Microsoft.Online.SharePoint.PowerShell"   
}
$scripts += [PSCustomObject]@{
    Name = "o365-connect-sac.ps1";
    Service = "Security and Compliance";
    Module = "MSOnline"    
}
$scripts += [PSCustomObject]@{
    Name = "o365-connect-mfa-s4b.ps1";
    Service = "Skype for Business/CSTeams";
    Module = "skypeonlineconnector"
}
$scripts += [PSCustomObject]@{
    Name = "o365-connect-exo.ps1";
    Service = "Exchange Online";
    Module ="ExchangeOnlineManagement"    
}
$scripts += [PSCustomObject]@{
    Name = "o365-connect-ctrldply.ps1";
    Service = "Central Add-in deployment";
    Module = "";    
}
$scripts += [PSCustomObject]@{
    Name = "o365-connect-aip.ps1";
    Service = "Azure Information Protection";
    Module = "Aipservice"    
}
$scripts += [PSCustomObject]@{
    Name = "o365-connect-aad.ps1";
    Service = "Azure AD";
    Module = "AzureAD"    
}
$scripts += [PSCustomObject]@{
    Name = "o365-connect.ps1";
    Service = "MS Online";  
    Module = "MSOnline"  
}
$scripts += [PSCustomObject]@{
    Name = "Az-connect.ps1";
    Service = "Azure";  
    Module = "Az.Accounts"  
}
$scripts += [PSCustomObject]@{
    Name = "Intune-connect.ps1";
    Service = "Intune";  
    Module = "Microsoft.Graph.Intune"  
}
$scripts += [PSCustomObject]@{
    Name = "msgraph-connect.ps1";
    Service = "Graph (old)";  
    Module = "MSGraph"  
}
$scripts += [PSCustomObject]@{
    Name = "mggraph-connect.ps1";
    Service = "Graph (new)";  
    Module = "MGGraph"  
}
$scripts += [PSCustomObject]@{
    Name = "o365-connect-mfa-ctldply.ps1";
    Service = "Add-ins";  
    Module = "O365CentralizedAddInDeployment"  
}
$scripts += [PSCustomObject]@{
    Name = "az-connect-si.ps1";
    Service = "Azure Security Insights";  
    Module = "az.securityinsights"  
}
$scripts += [PSCustomObject]@{
    Name = "o365-connect-pnp.ps1";
    Service = "SharePoint Online PNP";  
    Module = "pnp.powershell"  
}
$scripts += [PSCustomObject]@{
    Name = "o365-connect-msc.ps1";
    Service = "Microsoft 365 Commerce";  
    Module = "mscommerce"  
}
$scripts += [PSCustomObject]@{
    Name = "o365-connect-pa.ps1";
    Service = "PowerApps";  
    Module = "Microsoft.PowerApps.PowerShell"  
}
#EndRegion Modules

if (-not $noprompt) {
    try {
        $results = $scripts | select-object service | Sort-Object Service | Out-GridView -PassThru -title "Select services to connect to (Multiple selections permitted) "
    }
    catch {
        write-host -ForegroundColor yellow -backgroundcolor $errormessagecolor "`n[001] - Error getting options`n"
        if ($debug) {
            Stop-Transcript | Out-Null      ## Terminate transcription
        }
        exit 1                          ## Terminate script
    }
}
else {
    $results = $scripts
}

foreach ($result in $results) {
    foreach ($script in $scripts) {
        if ($result.service -eq $script.service) {
            $run=".\"+$script.Name
            if (-not [string]::isnullorempty($script.module)) {             ## If a PowerShell module is required to be installed?
                if (get-module -listavailable -name $script.module) {       ## Has the Online PowerShell module been loaded?
                    write-host -ForegroundColor $processmessagecolor $script.module,"module found"
                }
                else {
                    write-host -ForegroundColor yellow -backgroundcolor $errormessagecolor "`n[002] - Online PowerShell module",$script.module,"not installed. Please install and re-run script`n"
                    if ($debug) {
                        Stop-Transcript | Out-Null      ## Terminate transcription
                    }
                    exit 2                          ## Terminate script
                }
            }
            <# Test for script in current location #>
            if (-not (test-path -path $run)) {
                write-host -ForegroundColor yellow -backgroundcolor $errormessagecolor "`n[003] -",$script.name,"script not found in current directory - Please ensure exists first`n"
                if ($debug) {
                    Stop-Transcript | Out-Null      ## Terminate transcription
                }
                exit 3                          ## Terminate script
            }
            else {
                write-host -ForegroundColor $processmessagecolor $script.name,"script found in current directory`n"
            }
            if ($noupdate) {
                & $run -noupdate          ## Run script
            }
            else {
                & $run
            }
        }
    }
}

write-host -foregroundcolor $systemmessagecolor "`nMicrosoft Cloud connections menu script finished`n"
if ($debug) {
    Stop-Transcript | Out-Null
}