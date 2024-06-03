param(                         
    [switch]$prompt = $false, ## if -noprompt used then user will not be asked for any input
    [switch]$noupdate = $false, ## if -noupdate used then module will not be checked for more recent version
    [switch]$debug = $false       ## if -debug create a log file
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the SharePoint Online with PnP

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-pnp.ps1

Prerequisites = 3
1. Ensure pnp.powershell module is installed and updated
2. Ensure Microsoft.Graph module is installed and updated
3. Newerversions of the pnp.powershell module require PowerShell V7 or above

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
    write-host "Script activity logged at ..\o365-connect-pnp.txt"
    start-transcript "..\o365-connect-pnp.txt" | Out-Null                                        ## Log file created in parent directory that is overwritten on each run
}
write-host -foregroundcolor $systemmessagecolor "SharePoint Online PNP Connection script started`n"
write-host -ForegroundColor $processmessagecolor "Prompt =", ($prompt)
write-host -ForegroundColor $processmessagecolor "Debug =", ($debug)
write-host -ForegroundColor $processmessagecolor "Update =", (-not $noupdate)

$ps = $PSVersionTable.PSVersion
if ($ps.Major -lt 7) {
    write-host -foregroundcolor $errormessagecolor "`nThis script requires PowerShell version 7 or above`n"
    if ($debug) {
        Stop-Transcript | Out-Null                 ## Terminate transcription
    }
    exit 1
}
Write-host -foregroundcolor $processmessagecolor "`nDetected supported PowerShell version: $($ps.Major).$($ps.Minor)"

# Microsoft Online Module
if (get-module -listavailable -name Microsoft.Graph.Identity.DirectoryManagement) {
    ## Has the Microsoft Graph module been installed?
    write-host -ForegroundColor $processmessagecolor "Microsoft Graph Identity Directory Management module installed"
}
else {
    write-host -ForegroundColor $warningmessagecolor -backgroundcolor $errormessagecolor "[001] - Microsoft Graph Identity Directory Management module not installed`n"
    if (prompt) {
        do {
            $response = read-host -Prompt "`nDo you wish to install the Microsoft Graph Identity Directory Management module (Y/N)?"
        } until (-not [string]::isnullorempty($response))
        if ($result -eq 'Y' -or $result -eq 'y') {
            write-host -foregroundcolor $processmessagecolor "Installing Microsoft Graph Identity Directory Management module - Administration escalation required"
            Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force -confirm:$false" -wait -WindowStyle Hidden
            write-host -foregroundcolor $processmessagecolor "Microsoft Graph Identity Directory Management module installed"
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
        write-host -foregroundcolor $processmessagecolor "Installing Microsoft Graph Identity Directory Management module - Administration escalation required"
        Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "Microsoft Graph Identity Directory Management module installed"    
    }
}
if (-not $noupdate) {
    write-host -foregroundcolor $processmessagecolor "Check whether newer version of Microsoft Graph Directory Management module is available"
    #get version of the module (selects the first if there are more versions installed)
    $version = (Get-InstalledModule -name Microsoft.Graph.Identity.DirectoryManagement) | Sort-Object Version -Descending  | Select-Object Version -First 1
    #get version of the module in psgallery
    $psgalleryversion = Find-Module -Name Microsoft.Graph.Identity.DirectoryManagement | Sort-Object Version -Descending | Select-Object Version -First 1
    #convert to string for comparison
    $stringver = $version | Select-Object @{n = 'ModuleVersion'; e = { $_.Version -as [string] } }
    $a = $stringver | Select-Object Moduleversion -ExpandProperty Moduleversion
    #convert to string for comparison
    $onlinever = $psgalleryversion | Select-Object @{n = 'OnlineVersion'; e = { $_.Version -as [string] } }
    $b = $onlinever | Select-Object OnlineVersion -ExpandProperty OnlineVersion
    #version compare
    if ([version]"$a" -ge [version]"$b") {
        Write-Host -foregroundcolor $processmessagecolor "Local module $a greater or equal to Gallery module $b"
        write-host -foregroundcolor $processmessagecolor "No update required"
    }
    else {
        Write-Host -foregroundcolor $warningmessagecolor "Local module $a lower version than Gallery module $b"
        write-host -foregroundcolor $warningmessagecolor "Update recommended"
        if ($prompt) {
            do {
                $response = read-host -Prompt "`nDo you wish to update the Microsoft Graph Identity Directory Management PowerShell module (Y/N)?"
            } until (-not [string]::isnullorempty($response))
            if ($result -eq 'Y' -or $result -eq 'y') {
                write-host -foregroundcolor $processmessagecolor "Updating Microsoft Graph Identity Directory Management PowerShell module - Administration escalation required"
                Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force -confirm:$false" -wait -WindowStyle Hidden
                write-host -foregroundcolor $processmessagecolor "Microsoft Graph Identity Directory Management PowerShell module - updated"
            }
            else {
                write-host -foregroundcolor $processmessagecolor "Microsoft Graph Identity Directory Management PowerShell module - not updated"
            }
        }
        else {
            write-host -foregroundcolor $processmessagecolor "Microsoft Graph Identity Directory Management PowerShell module - Administration escalation required" 
            Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force -confirm:$false" -wait -WindowStyle Hidden
            write-host -foregroundcolor $processmessagecolor "Microsoft Graph Identity Directory Management PowerShell module - updated"
        }
    }
}

write-host -foregroundcolor $processmessagecolor "Microsoft Graph Identity Directory Management PowerShell module loading"
Try {
    Import-Module Microsoft.Graph.Identity.DirectoryManagement | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[002] - Unable to load Microsoft Graph Identity Directory Management PowerShell module`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                ## Terminate transcription
    }
    exit 2
}
write-host -foregroundcolor $processmessagecolor "Microsoft Graph Identity Directory Management PowerShell module loaded"

# SharePoint Online module
if (get-module -listavailable -name Microsoft.Graph.Sites) {
    ## Has the SharePOint Online PowerShell module been installed?
    write-host -ForegroundColor $processmessagecolor "SharePoint Online Graph PowerShell module installed"
}
else {
    write-host -ForegroundColor $warningmessagecolor -backgroundcolor $errormessagecolor "[004] - SharePoint Online Graph PowerShell module not installed`n"
    if ($prompt) {
        do {
            $response = read-host -Prompt "`nDo you wish to install the SharePoint Online Graph PowerShell module (Y/N)?"
        } until (-not [string]::isnullorempty($response))
        if ($result -eq 'Y' -or $result -eq 'y') {
            write-host -foregroundcolor $processmessagecolor "Installing SharePoint Online Graph PowerShell module - Administration escalation required"
            Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name Microsoft.Graph.Sites -Force -confirm:$false" -wait -WindowStyle Hidden
            write-host -foregroundcolor $processmessagecolor "SharePoint Online Online Graph PowerShell module installed"
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
        write-host -foregroundcolor $processmessagecolor "Installing SharePoint Online Graph module - Administration escalation required"
        Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name Microsoft.Graph.Sites -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "SharePoint Online Graph module installed"    
    }
}

if (-not $noupdate) {
    write-host -foregroundcolor $processmessagecolor "Check whether newer version of SharePoint Online Graph PowerShell module is available"
    #get version of the module (selects the first if there are more versions installed)
    $version = (Get-InstalledModule -name Microsoft.Graph.Sites) | Sort-Object Version -Descending  | Select-Object Version -First 1
    #get version of the module in psgallery
    $psgalleryversion = Find-Module -Name Microsoft.Graph.Sites | Sort-Object Version -Descending | Select-Object Version -First 1
    #convert to string for comparison
    $stringver = $version | Select-Object @{n = 'ModuleVersion'; e = { $_.Version -as [string] } }
    $a = $stringver | Select-Object Moduleversion -ExpandProperty Moduleversion
    #convert to string for comparison
    $onlinever = $psgalleryversion | Select-Object @{n = 'OnlineVersion'; e = { $_.Version -as [string] } }
    $b = $onlinever | Select-Object OnlineVersion -ExpandProperty OnlineVersion
    #version compare
    if ([version]"$a" -ge [version]"$b") {
        Write-Host -foregroundcolor $processmessagecolor "Local module $a greater or equal to Gallery module $b"
        write-host -foregroundcolor $processmessagecolor "No update required"
    }
    else {
        Write-Host -foregroundcolor $warningmessagecolor "Local module $a lower version than Gallery module $b"
        write-host -foregroundcolor $warningmessagecolor "Update recommended"
        if ($prompt) {
            do {
                $response = read-host -Prompt "`nDo you wish to update the SharePoint Online Graph PowerShell module (Y/N)?"
            } until (-not [string]::isnullorempty($response))
            if ($result -eq 'Y' -or $result -eq 'y') {
                write-host -foregroundcolor $processmessagecolor "Updating SharePoint Online Graph PowerShell module - Administration escalation required"
                Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name Microsoft.Graph.Sites -Force -confirm:$false" -wait -WindowStyle Hidden
                write-host -foregroundcolor $processmessagecolor "SharePoint Online PowerShell Graph module - updated"
            }
            else {
                write-host -foregroundcolor $processmessagecolor "SharePoint Online PowerShell Graph module - not updated"
            }
        }
        else {
            write-host -foregroundcolor $processmessagecolor "Updating SharePoint Online Graph PowerShell module - Administration escalation required" 
            Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name Microsoft.Graph.Sites -Force -confirm:$false" -wait -WindowStyle Hidden
            write-host -foregroundcolor $processmessagecolor "SharePoint Online Graph PowerShell module - updated"
        }
    }
}

write-host -foregroundcolor $processmessagecolor "Microsoft Graph SharePoint Online PowerShell module loading"
Try {
    Import-Module Microsoft.Graph.Sites | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[002] - Unable to load Microsoft Graph SharePoint Online PowerShell module`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                ## Terminate transcription
    }
    exit 2
}
write-host -foregroundcolor $processmessagecolor "Microsoft Graph SharePoint Online PowerShell module loaded"

## Connect to Office 365 admin service
write-host -foregroundcolor $processmessagecolor "Connecting to Microsoft Graph"
try {
    Connect-MgGraph -nowelcome -Scopes "Sites.Read.All", "sites.ReadWrite.All,Domain.Read.All"
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[003] - Unable to connect to Microsoft Graph`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                ## Terminate transcription
    }
    exit 3
}
write-host -foregroundcolor $processmessagecolor "Connected to Microsoft Graph"

## Auto detect SharePoint Online admin domain
write-host -foregroundcolor $processmessagecolor "Determining SharePoint URL"
$domains = get-mgdomain                        ## get a list of all domains in tenant
foreach ($domain in $domains) {
    ## loop through all these domains
    if ($domain.id.contains('onmicrosoft')) {
        ## find the onmicrosoft.com domain
        $onname = $domain.id.split(".")           ## split the onmicrosoft.com domain when found at the period. Will produce an array that contains each string as an element
        $tenantname = $onname[0]                    ## the first string in this array is the name of the tenant
    }                                               ## end of find the on.microsoft.com domain
}                                                   ## end of the domain checking look
$tenanturl = "https://" + $tenantname + "-admin.sharepoint.com"
Write-host -ForegroundColor $processmessagecolor "SharePoint admin URL =", $tenanturl

# SharePoint PNP Online module
if (get-module -listavailable -name pnp.powershell) {
    ## Has the SharePoint Online PNP PowerShell module been installed?
    write-host -ForegroundColor $processmessagecolor "SharePoint Online PNP PowerShell module installed"
}
else {
    write-host -ForegroundColor $warningmessagecolor -backgroundcolor $errormessagecolor "[004] - SharePoint Online PNP PowerShell module not installed`n"
    if ($prompt) {
        do {
            $response = read-host -Prompt "`nDo you wish to install the SharePoint Online PNP PowerShell module (Y/N)?"
        } until (-not [string]::isnullorempty($response))
        if ($result -eq 'Y' -or $result -eq 'y') {
            write-host -foregroundcolor $processmessagecolor "Installing SharePoint Online PNP PowerShell module - Administration escalation required"
            Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name pnp.powershell -Force -confirm:$false" -wait -WindowStyle Hidden
            write-host -foregroundcolor $processmessagecolor "SharePoint Online Online PNP PowerShell module installed"
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
        write-host -foregroundcolor $processmessagecolor "Installing SharePoint Online PNP module - Administration escalation required"
        Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name pnp.powershell -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "SharePoint Online PNP module installed"    
    }
}

if (-not $noupdate) {
    write-host -foregroundcolor $processmessagecolor "Check whether newer version of SharePoint Online PNP PowerShell module is available"
    #get version of the module (selects the first if there are more versions installed)
    $version = (Get-InstalledModule -name pnp.powershell) | Sort-Object Version -Descending  | Select-Object Version -First 1
    #get version of the module in psgallery
    $psgalleryversion = Find-Module -Name pnp.powershell | Sort-Object Version -Descending | Select-Object Version -First 1
    #convert to string for comparison
    $stringver = $version | Select-Object @{n = 'ModuleVersion'; e = { $_.Version -as [string] } }
    $a = $stringver | Select-Object Moduleversion -ExpandProperty Moduleversion
    #convert to string for comparison
    $onlinever = $psgalleryversion | Select-Object @{n = 'OnlineVersion'; e = { $_.Version -as [string] } }
    $b = $onlinever | Select-Object OnlineVersion -ExpandProperty OnlineVersion
    #version compare
    if ([version]"$a" -ge [version]"$b") {
        Write-Host -foregroundcolor $processmessagecolor "Local module $a greater or equal to Gallery module $b"
        write-host -foregroundcolor $processmessagecolor "No update required"
    }
    else {
        Write-Host -foregroundcolor $warningmessagecolor "Local module $a lower version than Gallery module $b"
        write-host -foregroundcolor $warningmessagecolor "Update recommended"
        if ($prompt) {
            do {
                $response = read-host -Prompt "`nDo you wish to update the SharePoint Online PNP PowerShell module (Y/N)?"
            } until (-not [string]::isnullorempty($response))
            if ($result -eq 'Y' -or $result -eq 'y') {
                write-host -foregroundcolor $processmessagecolor "Updating SharePoint Online PNP PowerShell module - Administration escalation required"
                Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name pnp.powershell -Force -confirm:$false" -wait -WindowStyle Hidden
                write-host -foregroundcolor $processmessagecolor "SharePoint Online PNP PowerShell module - updated"
            }
            else {
                write-host -foregroundcolor $processmessagecolor "SharePoint Online PNP PowerShell module - not updated"
            }
        }
        else {
            write-host -foregroundcolor $processmessagecolor "Updating SharePoint Online PNP PowerShell module - Administration escalation required" 
            Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name pnp.powershell -Force -confirm:$false" -wait -WindowStyle Hidden
            write-host -foregroundcolor $processmessagecolor "SharePoint Online PNP PowerShell module - updated"
        }
    }
}

write-host -foregroundcolor $processmessagecolor "Get all SharePoint Online sites"
$sites = (Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/sites?search=$($tenantname)" -Method GET).value
$root = (Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/sites/root" -Method GET)
##$sites = get-mgsite -search $tenantname
$siteSummary = @()
$siteSummary += [pscustomobject]@{       
    Name         = $root.displayname
    weburl       = $root.weburl
}
Foreach ($site in $sites) {
    $siteSummary += [pscustomobject]@{       
        Name         = $site.name
        weburl       = $site.weburl
    }
}
write-host -foregroundcolor $processmessagecolor "Sites found =",$sitesummary.count
$result = $sitesummary | select-object Name, weburl | Sort-Object Name,weburl | Out-GridView -OutputMode Single -title "Select SharePoint site to connect to with PNP"
write-host -foregroundcolor $processmessagecolor "Selected SharePoint Online site =", $result.weburl

# Import SharePoint Online PNP module
write-host -foregroundcolor $processmessagecolor "SharePoint Online PNP PowerShell module loading"
Try {
    Import-Module pnp.powershell | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[005] - Unable to load SharePoint Online PNP PowerShell module. Try using PowerShell V7 or above`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                ## Terminate transcription
    }
    exit 5
}
write-host -foregroundcolor $processmessagecolor "SharePoint Online PNP PowerShell module loaded"

# Connect to SharePoint Online PNP Service
write-host -foregroundcolor $processmessagecolor "Connecting to SharePoint PNP Online"
Try {
    connect-pnponline -url $result.weburl -launchbrowser -devicelogin | Out-Null
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[006] - Unable to connect to SharePoint Online PNP`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null                ## Terminate transcription
    }
    exit 6
}
write-host -foregroundcolor $processmessagecolor "Connected to SharePoint Online PNP`n"

write-host -foregroundcolor $systemmessagecolor "SharePoint Online PNP Connection script finished`n"
if ($debug) {
    Stop-Transcript | Out-Null
}