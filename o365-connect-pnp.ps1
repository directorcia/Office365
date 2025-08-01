param(                         
    [switch]$prompt = $false, ## if -noprompt used then user will not be asked for any input
    [switch]$noupdate = $false, ## if -noupdate used then module will not be checked for more recent version
    [switch]$debug = $false,      ## if -debug create a log file
    [string]$ClientId = $null     ## Optional: Existing Azure AD App Client ID for PnP connection
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the SharePoint Online with PnP
✅ FIXED: Microsoft Graph PowerShell SDK v2.26 compatibility issue resolved
✅ SOLUTION: Automatic detection and downgrade to stable v2.25.0 or acceptance of v2.28.0+
✅ NEW: Support for existing Azure AD App Client ID (-ClientId parameter)

Parameters:
-prompt     : Set to $true to prompt for user input during installations
-noupdate   : Set to $true to skip module updates (faster execution)
-debug      : Set to $true to create detailed log file
-ClientId   : Specify existing Azure AD App Client ID for PnP connection (avoids creating new app)

Usage Examples:
.\o365-connect-pnp.ps1                                    # Basic connection (creates new app if needed)
.\o365-connect-pnp.ps1 -ClientId "12345678-1234-1234-1234-123456789012"  # Use existing app
.\o365-connect-pnp.ps1 -prompt $true -debug $true         # Interactive mode with logging
.\o365-connect-pnp.ps1 -noupdate -ClientId "your-app-id"  # Fast connection with existing app

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-pnp.ps1
Documentation - https://github.com/directorcia/Office365/wiki/SharePoint-Online-PnP-Connection-Script

Prerequisites = 3
1. Ensure pnp.powershell module is installed and updated
2. Ensure Microsoft.Graph module is installed and updated (compatibility auto-fixed)
3. Newer versions of the pnp.powershell module require PowerShell V7 or above
4. (Optional) Existing Azure AD App with appropriate SharePoint permissions

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
write-host -ForegroundColor $processmessagecolor "Existing ClientId =", $(if ($ClientId) { "Provided" } else { "Not provided (will create new)" })

$ps = $PSVersionTable.PSVersion
if ($ps.Major -lt 7) {
    write-host -foregroundcolor $errormessagecolor "`nThis script requires PowerShell version 7 or above`n"
    if ($debug) {
        Stop-Transcript | Out-Null                 ## Terminate transcription
    }
    exit 1
}
Write-host -foregroundcolor $processmessagecolor "`nDetected supported PowerShell version: $($ps.Major).$($ps.Minor)"

#region Microsoft Graph Version Compatibility Fix
write-host -foregroundcolor $systemmessagecolor "`n🔧 MICROSOFT GRAPH VERSION COMPATIBILITY CHECK"
write-host -foregroundcolor $processmessagecolor "Checking for Microsoft Graph PowerShell SDK version compatibility..."

# Check current Microsoft.Graph version
$currentGraphModule = Get-InstalledModule -Name "Microsoft.Graph" -ErrorAction SilentlyContinue
if ($currentGraphModule) {
    write-host -ForegroundColor $processmessagecolor "Current Microsoft.Graph version: $($currentGraphModule.Version)"
    
    # Version 2.26 has known AzureIdentityAccessTokenProvider constructor bug
    if ($currentGraphModule.Version -eq "2.26.0") {
        write-host -ForegroundColor $errormessagecolor -backgroundcolor $warningmessagecolor "`n🚨 CRITICAL: Microsoft Graph PowerShell v2.26.0 has a known compatibility issue!"
        write-host -foregroundcolor $warningmessagecolor "Error: AzureIdentityAccessTokenProvider constructor bug in v2.26.0"
        write-host -foregroundcolor $processmessagecolor "Applying fix: Downgrading to stable version 2.25.0..."
        
        # Uninstall ALL Microsoft Graph modules (problematic version 2.26)
        write-host -foregroundcolor $warningmessagecolor "Uninstalling ALL Microsoft Graph modules v2.26.0..."
        try {
            # Get all Microsoft Graph modules and uninstall them
            $allGraphModules = Get-InstalledModule Microsoft.Graph* -ErrorAction SilentlyContinue
            foreach ($module in $allGraphModules) {
                write-host -foregroundcolor $warningmessagecolor "Uninstalling $($module.Name) v$($module.Version)..."
                Uninstall-Module -Name $module.Name -AllVersions -Force -ErrorAction SilentlyContinue
            }
            
            # Install stable version 2.25.0
            write-host -foregroundcolor $processmessagecolor "Installing Microsoft Graph v2.25.0 (stable) - this will install all required modules..."
            Install-Module -Name Microsoft.Graph -RequiredVersion 2.25.0 -AllowClobber -Force -Scope CurrentUser
            write-host -foregroundcolor $processmessagecolor "✅ Microsoft Graph PowerShell SDK fixed with stable v2.25.0"
        } catch {
            write-host -ForegroundColor $errormessagecolor "Failed to fix Microsoft Graph version: $_"
            if ($debug) {
                Stop-Transcript | Out-Null
            }
            exit 1
        }
    } elseif ($currentGraphModule.Version -lt [version]"2.28.0" -and $currentGraphModule.Version -ne [version]"2.25.0") {
        write-host -foregroundcolor $warningmessagecolor "Microsoft Graph PowerShell version $($currentGraphModule.Version) may have compatibility issues."
        write-host -foregroundcolor $processmessagecolor "Installing recommended stable version 2.25.0..."
        
        try {
            # Check for version mismatches in sub-modules
            $graphCoreModule = Get-InstalledModule -Name "Microsoft.Graph.Core" -ErrorAction SilentlyContinue
            if ($graphCoreModule -and $graphCoreModule.Version -ne $currentGraphModule.Version) {
                write-host -foregroundcolor $warningmessagecolor "Version mismatch detected! Microsoft.Graph.Core: $($graphCoreModule.Version), Microsoft.Graph: $($currentGraphModule.Version)"
                write-host -foregroundcolor $processmessagecolor "Uninstalling all Microsoft Graph modules to fix version conflicts..."
                
                # Get all Microsoft Graph modules and uninstall them
                $allGraphModules = Get-InstalledModule Microsoft.Graph* -ErrorAction SilentlyContinue
                foreach ($module in $allGraphModules) {
                    write-host -foregroundcolor $warningmessagecolor "Uninstalling $($module.Name) v$($module.Version)..."
                    Uninstall-Module -Name $module.Name -AllVersions -Force -ErrorAction SilentlyContinue
                }
            }
            
            # Install specific stable version
            Install-Module -Name Microsoft.Graph -RequiredVersion 2.25.0 -AllowClobber -Force -Scope CurrentUser
            write-host -foregroundcolor $processmessagecolor "✅ Microsoft Graph PowerShell SDK set to stable v2.25.0"
        } catch {
            write-host -ForegroundColor $errormessagecolor "Failed to install stable Microsoft Graph version: $_"
            if ($debug) {
                Stop-Transcript | Out-Null
            }
            exit 1
        }
    } elseif ($currentGraphModule.Version -ge [version]"2.28.0") {
        # Check for version mismatches even in newer versions
        $graphCoreModule = Get-InstalledModule -Name "Microsoft.Graph.Core" -ErrorAction SilentlyContinue
        if ($graphCoreModule -and $graphCoreModule.Version -lt [version]"2.28.0") {
            write-host -foregroundcolor $warningmessagecolor "Version mismatch detected! Microsoft.Graph.Core: $($graphCoreModule.Version) is older than Microsoft.Graph: $($currentGraphModule.Version)"
            write-host -foregroundcolor $processmessagecolor "Fixing module version conflicts..."
            
            try {
                # Reinstall the main module to ensure all sub-modules are updated
                Install-Module -Name Microsoft.Graph -RequiredVersion $currentGraphModule.Version -AllowClobber -Force -Scope CurrentUser
                write-host -foregroundcolor $processmessagecolor "✅ Microsoft Graph module versions synchronized"
            } catch {
                write-host -ForegroundColor $errormessagecolor "Failed to fix version conflicts: $_"
            }
        } else {
            write-host -foregroundcolor $processmessagecolor "✅ Microsoft Graph PowerShell version $($currentGraphModule.Version) is compatible"
        }
    } else {
        write-host -foregroundcolor $processmessagecolor "✅ Microsoft Graph PowerShell version $($currentGraphModule.Version) is stable"
    }
} else {
    # Install stable version for new installations
    write-host -foregroundcolor $processmessagecolor "Installing Microsoft Graph PowerShell SDK v2.25.0 (stable)..."
    try {
        Install-Module -Name Microsoft.Graph -RequiredVersion 2.25.0 -AllowClobber -Force -Scope CurrentUser
        write-host -foregroundcolor $processmessagecolor "✅ Microsoft Graph PowerShell SDK v2.25.0 installed successfully"
    } catch {
        write-host -ForegroundColor $errormessagecolor "Failed to install Microsoft Graph PowerShell SDK: $_"
        if ($debug) {
            Stop-Transcript | Out-Null
        }
        exit 1
    }
}

# Verify the fix worked and check for version mismatches
write-host -foregroundcolor $systemmessagecolor "`n🔍 VERIFYING MICROSOFT GRAPH COMPATIBILITY FIX"
try {
    Import-Module Microsoft.Graph.Authentication -Force
    $installedVersion = Get-Module Microsoft.Graph.Authentication | Select-Object -ExpandProperty Version
    write-host -foregroundcolor $processmessagecolor "✅ Microsoft Graph Authentication module version: $installedVersion"
    
    # Check Microsoft.Graph.Core version for compatibility
    $coreModule = Get-InstalledModule -Name "Microsoft.Graph.Core" -ErrorAction SilentlyContinue
    if ($coreModule) {
        write-host -foregroundcolor $processmessagecolor "✅ Microsoft Graph Core module version: $($coreModule.Version)"
        
        # Verify versions are compatible
        if ($coreModule.Version -lt [version]"2.25.0") {
            write-host -foregroundcolor $errormessagecolor "❌ Microsoft Graph Core module version $($coreModule.Version) is too old and incompatible!"
            write-host -foregroundcolor $processmessagecolor "Attempting to fix Core module version..."
            
            try {
                # Force reinstall of the entire Graph module suite
                Install-Module -Name Microsoft.Graph -RequiredVersion 2.25.0 -AllowClobber -Force -Scope CurrentUser
                Import-Module Microsoft.Graph.Authentication -Force
                write-host -foregroundcolor $processmessagecolor "✅ Microsoft Graph Core module updated"
            } catch {
                write-host -ForegroundColor $errormessagecolor "❌ Failed to update Microsoft Graph Core module: $_"
                throw "Microsoft Graph Core module incompatibility detected"
            }
        }
    }
    
    write-host -foregroundcolor $processmessagecolor "✅ Assembly loading compatibility verified"
} catch {
    write-host -ForegroundColor $errormessagecolor "❌ Microsoft Graph Authentication module import failed: $_"
    if ($debug) {
        Stop-Transcript | Out-Null
    }
    exit 1
}
#endregion

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

# Modern PnP authentication requires app registration as of September 2024
write-host -foregroundcolor $warningmessagecolor "ℹ️  PnP PowerShell Note: As of September 9, 2024, PnP requires custom app registration"

# Check if user provided an existing ClientId
if (-not [string]::IsNullOrEmpty($ClientId)) {
    write-host -foregroundcolor $processmessagecolor "Using provided Azure AD App Client ID: $ClientId"
    
    Try {
        write-host -foregroundcolor $processmessagecolor "Connecting to PnP using existing Azure AD app..."
        connect-pnponline -url $result.weburl -Interactive -ClientId $ClientId
        write-host -foregroundcolor $processmessagecolor "✅ Successfully connected to SharePoint PnP Online with existing app!"
    }
    catch {
        Write-Host -ForegroundColor $errormessagecolor "[006] - Unable to connect to SharePoint Online PNP with provided Client ID`n"
        Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
        
        write-host -foregroundcolor $warningmessagecolor "🔧 TROUBLESHOOTING: Existing Azure AD App Connection Failed"
        write-host -foregroundcolor $processmessagecolor "Possible issues with Client ID: $ClientId"
        write-host -foregroundcolor $processmessagecolor "1. Verify the Client ID is correct"
        write-host -foregroundcolor $processmessagecolor "2. Ensure the app has appropriate SharePoint permissions"
        write-host -foregroundcolor $processmessagecolor "3. Check that redirect URI includes 'http://localhost'"
        write-host -foregroundcolor $processmessagecolor "4. Verify app is configured as 'Public client (mobile & desktop)'"
        
        if ($debug) {
            Stop-Transcript | Out-Null
        }
        exit 6
    }
} else {
    write-host -foregroundcolor $processmessagecolor "No existing Azure AD App specified. Creating new app registration..."

    Try {
        # Check if we have an active Microsoft Graph session for app registration
        $mgContext = Get-MgContext
        if ($mgContext) {
            write-host -foregroundcolor $processmessagecolor "Using existing Microsoft Graph session for PnP authentication"
            
            # Extract tenant information from Microsoft Graph context
            $tenantId = $mgContext.TenantId
            $tenantDomain = ""
            
            try {
                # Get tenant domain information
                $tenantInfo = Get-MgOrganization | Select-Object -First 1
                if ($tenantInfo.VerifiedDomains) {
                    $primaryDomain = $tenantInfo.VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1
                    if ($primaryDomain) {
                        $tenantDomain = $primaryDomain.Name
                    } else {
                        # Fallback to any onmicrosoft.com domain
                        $tenantDomain = ($tenantInfo.VerifiedDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" } | Select-Object -First 1).Name
                    }
                }
                
                write-host -foregroundcolor $processmessagecolor "Tenant Domain: $tenantDomain"
            } catch {
                write-host -foregroundcolor $warningmessagecolor "Could not retrieve tenant domain: $($_.Exception.Message)"
            }
            
            # Try to register or use existing PnP app
            write-host -foregroundcolor $processmessagecolor "Creating new PnP Entra ID app registration for interactive login..."
            try {
                # Generate a unique app name
                $appName = "CIAOPS-PnP-PowerShell-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
                
                if ($tenantDomain) {
                    write-host -foregroundcolor $processmessagecolor "Registering PnP app: $appName"
                    $appRegistration = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName $appName -Tenant $tenantDomain
                    
                    if ($appRegistration) {
                        $newClientId = $appRegistration.'AzureAppId/ClientId'
                        write-host -foregroundcolor $processmessagecolor "✅ PnP app registration created successfully!"
                        write-host -foregroundcolor $processmessagecolor "App Name: $appName"
                        write-host -foregroundcolor $processmessagecolor "Client ID: $newClientId"
                        
                        # Store the new Client ID for future use
                        write-host -foregroundcolor $systemmessagecolor "💡 TIP: Save this Client ID for future use with -ClientId parameter"
                        write-host -foregroundcolor $systemmessagecolor "Future usage: .\o365-connect-pnp.ps1 -ClientId $newClientId"
                        
                        # Now connect using the newly created app
                        write-host -foregroundcolor $processmessagecolor "Connecting to PnP using newly created app..."
                        connect-pnponline -url $result.weburl -Interactive -ClientId $newClientId
                        write-host -foregroundcolor $processmessagecolor "✅ Successfully connected to SharePoint PnP Online with new app!"
                    } else {
                        throw "App registration returned null"
                    }
                } else {
                    throw "Could not determine tenant domain"
                }
            }
            catch {
                write-host -foregroundcolor $warningmessagecolor "App registration failed: $($_.Exception.Message)"
                
                # Fallback: Try direct interactive connection (may fail with new PnP requirements)
                write-host -foregroundcolor $processmessagecolor "Attempting fallback interactive connection..."
                try {
                    connect-pnponline -url $result.weburl -Interactive
                    write-host -foregroundcolor $processmessagecolor "✅ Successfully connected to SharePoint PnP Online (fallback)"
                } catch {
                    throw "Both app registration and fallback connection failed: $($_.Exception.Message)"
                }
            }
        } else {
            # No Microsoft Graph context, provide manual instructions
            throw "No Microsoft Graph context found. Please ensure Microsoft Graph connection is established first."
        }
    }
    catch {
        Write-Host -ForegroundColor $errormessagecolor "[006] - Unable to connect to SharePoint Online PNP`n"
        Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
        
        # Provide comprehensive guidance for manual app registration
        write-host -foregroundcolor $systemmessagecolor "`n📋 MANUAL SOLUTION: Register your own Entra ID application for PnP PowerShell"
        write-host -foregroundcolor $processmessagecolor "Option 1 - Automated Registration (Recommended):"
        write-host -foregroundcolor $processmessagecolor "1. Ensure you have Microsoft Graph connected first"
        write-host -foregroundcolor $processmessagecolor "2. Run: Register-PnPEntraIDAppForInteractiveLogin -ApplicationName 'CIAOPS-PnP-App' -Tenant yourdomain.onmicrosoft.com"
        write-host -foregroundcolor $processmessagecolor "3. Note the Client ID returned"
        write-host -foregroundcolor $processmessagecolor "4. Re-run this script with: .\o365-connect-pnp.ps1 -ClientId <your-client-id>"
        
        write-host -foregroundcolor $processmessagecolor "`nOption 2 - Manual Azure Portal Registration:"
        write-host -foregroundcolor $processmessagecolor "1. Go to Azure Portal > Entra ID > App Registrations > New Registration"
        write-host -foregroundcolor $processmessagecolor "2. Name: 'CIAOPS PnP PowerShell'"
        write-host -foregroundcolor $processmessagecolor "3. Redirect URI: 'http://localhost' (Public client/native)"
        write-host -foregroundcolor $processmessagecolor "4. Copy the Application (client) ID"
        write-host -foregroundcolor $processmessagecolor "5. Re-run this script with: .\o365-connect-pnp.ps1 -ClientId <your-client-id>"
        
        write-host -foregroundcolor $processmessagecolor "`nDocumentation: https://pnp.github.io/powershell/articles/registerapplication.html"
        
        if ($debug) {
            Stop-Transcript | Out-Null
        }
        exit 6
    }
}
write-host -foregroundcolor $processmessagecolor "Connected to SharePoint Online PNP`n"

write-host -foregroundcolor $systemmessagecolor "SharePoint Online PNP Connection script finished`n"
if ($debug) {
    Stop-Transcript | Out-Null
}