param(                         
    [switch]$noprompt = $false,   ## if -noprompt used then user will not be asked for any input
    [switch]$noupdate = $false,   ## if -noupdate used then module will not be checked for more recent version
    [switch]$debug = $false       ## if -debug create a log file
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the Office 365 Centralized Deployment for add ins

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-ctrldply.ps1

Prerequisites = 1
1. Ensure powershell cmdlest for Centralized deployment installed or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host
if ($debug) {
    write-host "Script activity logged at ..\o365-connect-ctrldply.txt"
    start-transcript "..\o365-connect-ctrldply.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

write-host -foregroundcolor $systemmessagecolor "Centralized Deployment Connection script started`n"
write-host -ForegroundColor $processmessagecolor "Prompt =",(-not $noprompt)

if (get-module -listavailable -name O365CentralizedAddInDeployment) {    ## Has the Centralized Deployment PowerShell module been installed?
    write-host -ForegroundColor $processmessagecolor "Centralized Deployment PowerShell module installed"
}
else {
    write-host -ForegroundColor $warningmessagecolor -backgroundcolor $errormessagecolor "[001] - Centralized Deployment PowerShell module not installed`n"
    if (-not $noprompt) {
        do {
            $response = read-host -Prompt "`nDo you wish to install the Centralized Deployment PowerShell module (Y/N)?"
        } until (-not [string]::isnullorempty($response))
        if ($result -eq 'Y' -or $result -eq 'y') {
            write-host -foregroundcolor $processmessagecolor "Installing Centralized Deployment PowerShell module - Administration escalation required"
            Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name O365CentralizedAddInDeployment -Force -confirm:$false" -wait -WindowStyle Hidden
            write-host -foregroundcolor $processmessagecolor "Azure Centralized Deployment PowerShell module installed"
        }
        else {
            write-host -foregroundcolor $processmessagecolor "Terminating script"
            if ($debug) {
                Stop-Transcript | Out-Null                 ## Terminate transcription
            }
            exit 1                          ## Terminate script
        }
    }
    else {
        write-host -foregroundcolor $processmessagecolor "Installing Centralized Deployment PowerShell module - Administration escalation required"
        Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name O365CentralizedAddInDeployment -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "Centralized Deployment PowerShell module installed"    
    }
}

if (-not $noupdate) {
    write-host -foregroundcolor $processmessagecolor "Check whether newer version of Centralized Deployment PowerShell module is available"
    #get version of the module (selects the first if there are more versions installed)
    $version = (Get-InstalledModule -name O365CentralizedAddInDeployment) | Sort-Object Version -Descending  | Select-Object Version -First 1
    #get version of the module in psgallery
    $psgalleryversion = Find-Module -Name O365CentralizedAddInDeployment | Sort-Object Version -Descending | Select-Object Version -First 1
    #convert to string for comparison
    $stringver = $version | Select-Object @{n='ModuleVersion'; e={$_.Version -as [string]}}
    $a = $stringver | Select-Object Moduleversion -ExpandProperty Moduleversion
    #convert to string for comparison
    $onlinever = $psgalleryversion | Select-Object @{n='OnlineVersion'; e={$_.Version -as [string]}}
    $b = $onlinever | Select-Object OnlineVersion -ExpandProperty OnlineVersion
    #version compare
    if ([version]"$a" -ge [version]"$b") {
        Write-Host -foregroundcolor $processmessagecolor "Local module $a greater or equal to Gallery module $b"
        write-host -foregroundcolor $processmessagecolor "No update required"
    }
    else {
        Write-Host -foregroundcolor $warningmessagecolor "Local module $a lower version than Gallery module $b"
        write-host -foregroundcolor $warningmessagecolor "Update recommended"
        if (-not $noprompt) {
            do {
                $response = read-host -Prompt "`nDo you wish to update the Centralized Deployment PowerShell module (Y/N)?"
            } until (-not [string]::isnullorempty($response))
            if ($result -eq 'Y' -or $result -eq 'y') {
                write-host -foregroundcolor $processmessagecolor "Updating Centralized Deployment PowerShell module - Administration escalation required"
                Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name O365CentralizedAddInDeployment -Force -confirm:$false" -wait -WindowStyle Hidden
                write-host -foregroundcolor $processmessagecolor "Centralized Deployment PowerShell module - updated"
            }
            else {
                write-host -foregroundcolor $processmessagecolor "Centralized Deployment PowerShell module - not updated"
            }
        }
        else {
        write-host -foregroundcolor $processmessagecolor "Updating Centralized Deployment PowerShell module - Administration escalation required" 
        Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name O365CentralizedAddInDeployment -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "Centralized Deployment PowerShell module - updated"
        }
    }
}

write-host -foregroundcolor $processmessagecolor "Centralized Deployment PowerShell module loading"
Try {
    Import-Module O365CentralizedAddInDeployment | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[002] - Unable to load Centralized Deployment PowerShell module`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                ## Terminate transcription
    }
    exit 2
}
write-host -foregroundcolor $processmessagecolor "Centralized Deployment PowerShell module loaded"

## Connect to Centralized Deployment
write-host -foregroundcolor $processmessagecolor "Connecting to Centralized Deployment"
try {
    $result = Connect-OrganizationAddInService | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[003] - Unable to connect to Centralized Deployment`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                 ## Terminate transcription
    }
    exit 3 
}

write-host -foregroundcolor $processmessagecolor "Connected to Centralized Deployment`n"
write-host -foregroundcolor $systemmessagecolor "Centralized Deployment Connection script finished`n"
if ($debug) {
    Stop-Transcript | Out-Null
}