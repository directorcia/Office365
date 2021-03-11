param(                         
    [switch]$noprompt = $false,   ## if -noprompt used then user will not be asked for any input
    [switch]$debug = $false       ## if -debug create a log file
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into Microsoft Teams

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-tms.ps1

Prerequisites = 1
1. Ensure Micosoft Teams Module is install or updated

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$warningmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
## set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

if ($debug) {
    start-transcript "..\o365-connect-tms.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}
write-host -foregroundcolor $systemmessagecolor "Microsoft Teams connection script started`n"
write-host -ForegroundColor $processmessagecolor "Debug = ",$debug
write-host -ForegroundColor $processmessagecolor "Prompt = ",(-not $noprompt)

if (get-module -listavailable -name MicrosoftTeams) {    ## Has the Teams PowerShell module been installed?
    write-host -ForegroundColor $processmessagecolor "Teams PowerShell module installed"
}
else {
    write-host -ForegroundColor $warningmessagecolor -backgroundcolor $errormessagecolor "[001] - Teams PowerShell module not installed`n"
    if (-not $noprompt) {
        do {
            $response = read-host -Prompt "`nDo you wish to install the Microsoft Teams PowerShell module (Y/N)?"
        } until (-not [string]::isnullorempty($response))
        if ($result -eq 'Y' -or $result -eq 'y') {
            write-host -foregroundcolor $processmessagecolor "Installing Microsoft Teams PowerShell module - Administration escalation required"
            Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name MicrosoftTeams -Force -confirm:$false" -wait -WindowStyle Hidden
            write-host -foregroundcolor $processmessagecolor "Microsoft Teams PowerShell module installed"
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
        write-host -foregroundcolor $processmessagecolor "Installing Microsoft Teams PowerShell module - Administration escalation required"
        Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name MicrosoftTeams -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "Microsoft Teams PowerShell module installed"    
    }
    
}
try {
    $result = import-module MicrosoftTeams
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[002] - Unable to load Microsoft Teams PowerShell module`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                ## Terminate transcription
    }
    exit 2
}

write-host -foregroundcolor $processmessagecolor "Microsoft Teams PowerShell module loaded"

## Connect to Microsoft Teams service
try {
    $result=Connect-MicrosoftTeams
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[003] - Unable to connect to Microsoft Teams`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                 ## Terminate transcription
    }
    exit 3
}
write-host -foregroundcolor $processmessagecolor "Now connected to Microsoft Teams`n"
write-host -foregroundcolor $systemmessagecolor "Microsoft Teams connection script finished`n"
if ($debug) {
    Stop-Transcript | Out-Null
}