param(                         
    [switch]$noprompt = $false,   ## if -noprompt used then user will not be asked for any input
    [switch]$noupdate = $false,   ## if -noupdate used then module will not be checked for more recent version
    [switch]$debug = $false       ## if -debug create a log file
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.
Description - Log to Exchange Online using the V2 module
Reference - https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/exchange-online-powershell-v2/exchange-online-powershell-v2?view=exchange-ps
Source - https://github.com/directorcia/Office365/blob/master/o365-connect-exo.ps1
Prerequisites = 1
1. Ensure Exchange Online module is installed
More scripts available by joining http://www.ciaopspatron.com
#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

## If you have running scripts that don't have a certificate, run this command once to disable that level of security
##  set-executionpolicy -executionpolicy bypass -scope currentuser -force

Clear-Host

## Enforce higher version of TLS
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ($debug) {
    write-host "Script activity logged at ..\o365-connect-exo.txt"
    start-transcript "..\o365-connect-exo.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}

write-host -foregroundcolor $systemmessagecolor "Exchange Online Connection script started`n"
write-host -ForegroundColor $processmessagecolor "Prompt =",(-not $noprompt)

if (get-module -listavailable -name ExchangeOnlineManagement) {    ## Has the Exchange Online PowerShell module been installed?
    write-host -ForegroundColor $processmessagecolor "Exchange Online PowerShell module installed"
}
else {
    write-host -ForegroundColor $warningmessagecolor -backgroundcolor $errormessagecolor "[001] - Exchange Online PowerShell module not installed`n"
    if (-not $noprompt) {
        do {
            $response = read-host -Prompt "`nDo you wish to install the Exchange Online PowerShell module (Y/N)?"
        } until ($response -match '^[YyNn]$')
        if ($response -eq 'Y') {
            write-host -foregroundcolor $processmessagecolor "Installing PowerShellGet module - Administration escalation required"
            Start-Process powershell -Verb runAs -ArgumentList "Install-Module PowershellGet -Force" -wait -WindowStyle Hidden       
            write-host -foregroundcolor $processmessagecolor "Installing Exchange Online PowerShell module - Administration escalation required"
            Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name ExchangeOnlineManagement -Force -confirm:$false" -wait -WindowStyle Hidden
            write-host -foregroundcolor $processmessagecolor "Exchange Online PowerShell module installed"
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
        write-host -foregroundcolor $processmessagecolor "Installing PowerShellGet module - Administration escalation required"
        Start-Process powershell -Verb runAs -ArgumentList "Install-Module PowershellGet -Force" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "Installing Exchange Online PowerShell module - Administration escalation required"
        Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name ExchangeOnlineManagement -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "Exchange Online PowerShell module installed"    
    }
}
if (-not $noupdate) {
    write-host -foregroundcolor $processmessagecolor "Check whether newer version of Exchange Online PowerShell module is available"
    try {
        $localversion = (Get-InstalledModule -Name ExchangeOnlineManagement -ErrorAction Stop | Sort-Object Version -Descending | Select-Object -First 1).Version
        $galleryversion = (Find-Module -Name ExchangeOnlineManagement -ErrorAction Stop | Sort-Object Version -Descending | Select-Object -First 1).Version
    }
    catch {
        Write-Host -ForegroundColor $warningmessagecolor "Unable to determine module versions ($($_.Exception.Message)). Skipping update check"
        $localversion = $null
    }
    if ($null -eq $localversion) {
        ## version check skipped
    }
    elseif ([version]$localversion -ge [version]$galleryversion) {
        Write-Host -foregroundcolor $processmessagecolor "Local module $localversion greater or equal to Gallery module $galleryversion"
        write-host -foregroundcolor $processmessagecolor "No update required"
    }
    else {
        Write-Host -foregroundcolor $warningmessagecolor "Local module $localversion lower version than Gallery module $galleryversion"
        write-host -foregroundcolor $warningmessagecolor "Update recommended"
        if (-not $noprompt) {
            do {
                $response = read-host -Prompt "`nDo you wish to update the Exchange Online PowerShell module (Y/N)?"
            } until ($response -match '^[YyNn]$')
            if ($response -eq 'Y') {
                write-host -foregroundcolor $processmessagecolor "Updating Exchange Online PowerShell module - Administration escalation required"
                Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name ExchangeOnlineManagement -Force -confirm:$false" -wait -WindowStyle Hidden
                write-host -foregroundcolor $processmessagecolor "Exchange Online PowerShell module - updated"
            }
            else {
                write-host -foregroundcolor $processmessagecolor "Exchange Online PowerShell module - not updated"
            }
        }
        else {
        write-host -foregroundcolor $processmessagecolor "Updating Exchange Online PowerShell module - Administration escalation required" 
        Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name ExchangeOnlineManagement -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "Exchange Online PowerShell module - updated"
        }
    }
}
write-host -foregroundcolor $processmessagecolor "Exchange Online PowerShell module loading"
Try {
    Import-Module ExchangeOnlineManagement | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[002] - Unable to load Exchange Online PowerShell module`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                ## Terminate transcription
    }
    exit 2
}
write-host -foregroundcolor $processmessagecolor "Exchange Online PowerShell module loaded"

## Connect to Exchange Online service
write-host -foregroundcolor $processmessagecolor "Connecting to Exchange Online"
$connectparams = (Get-Command Connect-ExchangeOnline).Parameters
$connected = $false
try {
    Connect-ExchangeOnline -ShowProgress:$false -ShowBanner:$false -ErrorAction Stop | Out-Null
    $connected = $true
}
catch {
    Write-Host -ForegroundColor $warningmessagecolor "Primary connect failed ($($_.Exception.Message))"
}
if (-not $connected -and $connectparams.ContainsKey('UseWebLogin')) {      ## UseWebLogin removed in module v3+
    Write-Host -ForegroundColor $warningmessagecolor "Retrying with web login..."
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Connect-ExchangeOnline -UseWebLogin -ShowBanner:$false -ErrorAction Stop | Out-Null
        $connected = $true
    }
    catch {
        Write-Host -ForegroundColor $warningmessagecolor "Web login failed ($($_.Exception.Message))"
    }
}
if (-not $connected -and $connectparams.ContainsKey('Device') -and $PSVersionTable.PSVersion.Major -ge 7) {    ## Device code requires PowerShell 7+
    Write-Host -ForegroundColor $warningmessagecolor "Retrying with device code..."
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop | Out-Null
        $connected = $true
    }
    catch {
        Write-Host -ForegroundColor $warningmessagecolor "Device code login failed ($($_.Exception.Message))"
    }
}
if (-not $connected) {
    Write-Host -ForegroundColor $errormessagecolor "[003] - Unable to connect to Exchange Online`n"
    if ($debug) {
        Stop-Transcript | Out-Null                 ## Terminate transcription
    }
    exit 3
}

write-host -foregroundcolor $processmessagecolor "Connected to Exchange Online`n"
write-host -foregroundcolor $systemmessagecolor "Exchange Online Connection script finished`n"
if ($debug) {
    Stop-Transcript | Out-Null
}
