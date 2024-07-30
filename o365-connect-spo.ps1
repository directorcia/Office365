param(                         
    [switch]$noprompt = $false,   ## if -noprompt used then user will not be asked for any input
    [switch]$noupdate = $false,   ## if -noupdate used then module will not be checked for more recent version
    [switch]$debug = $false       ## if -debug create a log file
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the Office 365 admin portal and the SharePoint Online portal

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-spo.ps1

Prerequisites = 2
1. Ensure Graph module installed or updated
2. Ensure SharePoint online PowerShell module installed or updated

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
    start-transcript "..\o365-connect-spo.txt" | Out-Null
    write-host -foregroundcolor $processmessagecolor "[Info] = Script activity logged at ..\o365-connect-spo.txt`n"
}
else {
    write-host -foregroundcolor $processmessagecolor "[Info] = Debug mode disabled`n"
}
write-host -foregroundcolor $systemmessagecolor "SharePoint Online Connection script started`n"
if ($prompt) {
    write-host -foregroundcolor $processmessagecolor "[Info] = Prompt mode enabled"
}
else {
    write-host -foregroundcolor $processmessagecolor "[Info] = Prompt mode disabled"
}
write-host -foregroundcolor $processmessagecolor "[Info] = Checking PowerShell version"
$ps = $PSVersionTable.PSVersion
Write-host -foregroundcolor $processmessagecolor "- Detected supported PowerShell version: $($ps.Major).$($ps.Minor)"

# Microsoft Online Module
if (get-module -listavailable -name Microsoft.Graph.Identity.DirectoryManagement) {    ## Has the Microsoft Online PowerShell module been installed?
    write-host -ForegroundColor $processmessagecolor "Microsoft Graph PowerShell module installed"
}
else {
    write-host -ForegroundColor $warningmessagecolor -backgroundcolor $errormessagecolor "[001] - Microsoft Graph PowerShell module not installed`n"
    if (-not $noprompt) {
        do {
            $response = read-host -Prompt "`nDo you wish to install the Microsoft Graph PowerShell module (Y/N)?"
        } until (-not [string]::isnullorempty($response))
        if ($result -eq 'Y' -or $result -eq 'y') {
            write-host -foregroundcolor $processmessagecolor "Installing Microsoft Graph PowerShell module - Administration escalation required"
            Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force -confirm:$false" -wait -WindowStyle Hidden
            write-host -foregroundcolor $processmessagecolor "Microsoft Graph PowerShell module installed"
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
        write-host -foregroundcolor $processmessagecolor "Installing Microsoft Graph module - Administration escalation required"
        Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "Microsoft Graph module installed"    
    }
}
if (-not $noupdate) {
    write-host -foregroundcolor $processmessagecolor "Check whether newer version of Microsoft Graph PowerShell module is available"
    #get version of the module (selects the first if there are more versions installed)
    $version = (Get-InstalledModule -name Microsoft.Graph.Identity.DirectoryManagement) | Sort-Object Version -Descending  | Select-Object Version -First 1
    #get version of the module in psgallery
    $psgalleryversion = Find-Module -Name Microsoft.Graph.Identity.DirectoryManagement | Sort-Object Version -Descending | Select-Object Version -First 1
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
                $response = read-host -Prompt "`nDo you wish to update the Microsoft Graph PowerShell module (Y/N)?"
            } until (-not [string]::isnullorempty($response))
            if ($response -eq 'Y' -or $response -eq 'y') {
                write-host -foregroundcolor $processmessagecolor "Updating Microsoft Graph PowerShell module - Administration escalation required"
                Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force -confirm:$false" -wait -WindowStyle Hidden
                write-host -foregroundcolor $processmessagecolor "Microsoft Graph PowerShell module - updated"
            }
            else {
                write-host -foregroundcolor $processmessagecolor "Microsoft Graph PowerShell module - not updated"
            }
        }
        else {
        write-host -foregroundcolor $processmessagecolor "Microsoft Graph PowerShell module - Administration escalation required" 
        Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "Microsoft Graph PowerShell module - updated"
        }
    }
}

write-host -foregroundcolor $processmessagecolor "Microsoft Graph PowerShell module loading"
Try {
    
    Import-Module Microsoft.Graph.Identity.DirectoryManagement | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[002] - Unable to load Microsoft Graph PowerShell module`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                ## Terminate transcription
    }
    exit 2
}
write-host -foregroundcolor $processmessagecolor "Microsoft Graph PowerShell module loaded"

## Connect to Office 365 admin service
write-host -foregroundcolor $processmessagecolor "Connecting to Microsoft 365 Admin service"
try {
    Connect-MGGraph -Scopes "Domain.Read.All" | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[003] - Unable to connect to Microsoft Graph`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                ## Terminate transcription
    }
    exit 3
}
write-host -foregroundcolor $processmessagecolor "Connected to Microsoft 365 Admin service"

## Auto detect SharePoint Online admin domain
write-host -foregroundcolor $processmessagecolor "Determining SharePoint Administration URL"
$domains = Get-MGDomain
foreach ($domain in $domains.id) {                     ## loop through all these domains
    if ($domain.contains('onmicrosoft')) {     ## find the onmicrosoft.com domain
        $onname = $domain.split(".")           ## split the onmicrosoft.com domain when found at the period. Will produce an array that contains each string as an element
        $tenantname = $onname[0]                    ## the first string in this array is the name of the tenant
    }                                               ## end of find the on.microsoft.com domain
}                                                   ## end of the domain checking look
$tenanturl = "https://" + $tenantname + "-admin.sharepoint.com"
Write-host -ForegroundColor $processmessagecolor "SharePoint admin URL =", $tenanturl

# SharePoint Online module
if (get-module -listavailable -name microsoft.online.sharepoint.powershell) {    ## Has the SharePOint Online PowerShell module been installed?
    write-host -ForegroundColor $processmessagecolor "SharePoint Online PowerShell module installed"
}
else {
    write-host -ForegroundColor $warningmessagecolor -backgroundcolor $errormessagecolor "[004] - SharePoint Online PowerShell module not installed`n"
    if (-not $noprompt) {
        do {
            $response = read-host -Prompt "`nDo you wish to install the SharePoint Online PowerShell module (Y/N)?"
        } until (-not [string]::isnullorempty($response))
        if ($result -eq 'Y' -or $result -eq 'y') {
            write-host -foregroundcolor $processmessagecolor "Installing SharePoint Online PowerShell module - Administration escalation required"
            Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name microsoft.online.sharepoint.powershell -Force -confirm:$false" -wait -WindowStyle Hidden
            write-host -foregroundcolor $processmessagecolor "SharePoint Online Online PowerShell module installed"
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
        write-host -foregroundcolor $processmessagecolor "Installing SharePoint Online module - Administration escalation required"
        Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name microsoft.online.sharepoint.powershell -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "SharePoint Online module installed"    
    }
}

if (-not $noupdate) {
    write-host -foregroundcolor $processmessagecolor "Check whether newer version of SharePoint Online PowerShell module is available"
    #get version of the module (selects the first if there are more versions installed)
    $version = (Get-InstalledModule -name microsoft.online.sharepoint.powershell) | Sort-Object Version -Descending  | Select-Object Version -First 1
    #get version of the module in psgallery
    $psgalleryversion = Find-Module -Name microsoft.online.sharepoint.powershell | Sort-Object Version -Descending | Select-Object Version -First 1
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
                $response = read-host -Prompt "`nDo you wish to update the SharePoint Online PowerShell module (Y/N)?"
            } until (-not [string]::isnullorempty($response))
            if ($response -eq 'Y' -or $response -eq 'y') {
                write-host -foregroundcolor $processmessagecolor "Updating SharePoint Online PowerShell module - Administration escalation required"
                Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name microsoft.online.sharepoint.powershell -Force -confirm:$false" -wait -WindowStyle Hidden
                write-host -foregroundcolor $processmessagecolor "SharePoint Online PowerShell module - updated"
            }
            else {
                write-host -foregroundcolor $processmessagecolor "SharePoint Online PowerShell module - not updated"
            }
        }
        else {
        write-host -foregroundcolor $processmessagecolor "Updating SharePoint Online PowerShell module - Administration escalation required" 
        Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name microsoft.online.sharepoint.powershell -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "SharePoint Online PowerShell module - updated"
        }
    }
}

# Import SharePoint Online module
write-host -foregroundcolor $processmessagecolor "SharePoint Online PowerShell module loading"
Try {
    if ($ps.Major -lt 6) {
        $result = Import-Module microsoft.online.sharepoint.powershell -disablenamechecking
    }
    else {
        write-host -foregroundcolor $processmessagecolor "[Info] = Using compatibility mode`n"
        $result = Import-Module microsoft.online.sharepoint.powershell -disablenamechecking -UseWindowsPowerShell
    }
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[005] - Unable to load SharePoint Online PowerShell module`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                ## Terminate transcription
    }
    exit 5
}
write-host -foregroundcolor $processmessagecolor "SharePoint Online PowerShell module loaded"

# Connect to SharePoint Online Service
write-host -foregroundcolor $processmessagecolor "Connecting to SharePoint Online"
Try {
    connect-sposervice -url $tenanturl | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[006] - Unable to connect to SharePoint Online`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                ## Terminate transcription
    }
    exit 6
}
write-host -foregroundcolor $processmessagecolor "Connected to SharePoint Online`n"

write-host -foregroundcolor $systemmessagecolor "SharePoint Online Connection script finished`n"
if ($debug) {
    Stop-Transcript | Out-Null
}