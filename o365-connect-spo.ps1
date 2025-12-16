param(                         
    [switch]$noprompt = $false,   ## if -noprompt used then user will not be asked for any input
    [switch]$noupdate = $false,   ## if -noupdate used then module will not be checked for more recent version
    [switch]$debug = $false,      ## if -debug create a log file
    [switch]$UseDeviceCode = $false ## if -UseDeviceCode parameter use device code authentication
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Description - Log into the Office 365 admin portal and the SharePoint Online portal

Source - https://github.com/directorcia/Office365/blob/master/o365-connect-spo.ps1
Documentation - https://github.com/directorcia/Office365/wiki/SharePoint-Online-Connection-Script

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

$RequiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Identity.DirectoryManagement")

# Check if the assembly is already loaded in the .NET AppDomain to avoid conflicts
$LoadedAssembly = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq "Microsoft.Graph.Authentication" } | Select-Object -First 1
$TargetVersion = $null

if ($LoadedAssembly) {
    $LoadedVersion = $LoadedAssembly.GetName().Version
    # Map assembly version to module version (Major.Minor.Build)
    $TargetVersion = [Version]"$($LoadedVersion.Major).$($LoadedVersion.Minor).$($LoadedVersion.Build)"
    Write-Warning "Microsoft.Graph.Authentication assembly $TargetVersion is already loaded. Locking to this version."
}
else {
    # Attempt to clear existing Graph modules to prevent version conflicts
    Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue

    # Find the highest common version available for all required modules
    Write-Host -ForegroundColor $processmessagecolor "Resolving module versions..."
    $CommonVersions = $null
    foreach ($Name in $RequiredModules) {
        $Available = Get-Module -ListAvailable $Name
        if ($null -eq $Available) {
            Write-Host -ForegroundColor $errormessagecolor "[002] - Required module $Name is not installed.`n"
            if ($debug) {
                Stop-Transcript | Out-Null
            }
            exit 2
        }
        $Versions = $Available.Version
        if ($null -eq $CommonVersions) {
            $CommonVersions = $Versions
        } else {
            $CommonVersions = $CommonVersions | Where-Object { $_ -in $Versions }
        }
    }
    $TargetVersion = $CommonVersions | Sort-Object -Descending | Select-Object -First 1
}

if ($TargetVersion) {
    Write-Host -ForegroundColor $processmessagecolor "Targeting Microsoft Graph version: $TargetVersion"
    foreach ($Name in $RequiredModules) {
        try {
            # Find the specific module file for this version
            $ModuleInfo = Get-Module -ListAvailable $Name | Where-Object Version -eq $TargetVersion | Select-Object -First 1
            if ($ModuleInfo) {
                Import-Module $ModuleInfo.Path -Force -ErrorAction Stop
            }
            else {
                throw "Module path not found for $Name version $TargetVersion"
            }
        }
        catch {
            Write-Host -ForegroundColor $errormessagecolor "[002] - Failed to load $Name version $TargetVersion`n"
            Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
            Write-Warning "A restart of the PowerShell session is likely required to clear loaded assemblies."
            if ($debug) {
                Stop-Transcript | Out-Null
            }
            exit 2
        }
    }
}
else {
    Write-Warning "Could not find a common version for all modules. Attempting to load latest available..."
    foreach ($Module in $RequiredModules) {
        try {
            Import-Module $Module -Force -ErrorAction Stop
        }
        catch {
            Write-Host -ForegroundColor $errormessagecolor "[002] - Unable to load Microsoft Graph PowerShell module`n"
            Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
            if ($debug) {
                Stop-Transcript | Out-Null
            }
            exit 2
        }
    }
}

write-host -foregroundcolor $processmessagecolor "Microsoft Graph PowerShell module loaded"

## Connect to Office 365 admin service
write-host -foregroundcolor $processmessagecolor "Connecting to Microsoft 365 Admin service"
$context = Get-MgContext

# Disconnect any existing stale connections
if ($null -ne $context) {
    Write-Host -ForegroundColor $warningmessagecolor "Existing connection detected. Disconnecting to ensure fresh authentication..."
    Disconnect-MgGraph | Out-Null
    $context = $null
}

if ($null -eq $context) {
    try {
        if ($UseDeviceCode) {
            Write-Host -ForegroundColor $warningmessagecolor "`nUsing device code authentication..."
            Write-Host -ForegroundColor $warningmessagecolor "Opening browser to https://microsoft.com/devicelogin..."
            try { Start-Process "https://microsoft.com/devicelogin" } catch { Write-Warning "Could not open browser automatically." }

            Connect-MgGraph -UseDeviceCode -Scopes "Domain.Read.All","Organization.Read.All" -NoWelcome -ErrorAction Stop 2>&1 | ForEach-Object {
                if ($_ -match "enter the code ([A-Z0-9]+) to authenticate") {
                    Write-Host -ForegroundColor Cyan "`nDevice Code: $($matches[1])`n"
                }
                elseif ($_ -notmatch "To sign in, use a web browser") {
                    Write-Host $_
                }
            }
        }
        else {
            try {
                Connect-MgGraph -Scopes "Domain.Read.All","Organization.Read.All" -NoWelcome -ErrorAction Stop | Out-Null
            }
            catch {
                Write-Host -ForegroundColor $warningmessagecolor "Interactive login failed. Falling back to Device Code authentication..."
                Write-Host -ForegroundColor $warningmessagecolor "Opening browser to https://microsoft.com/devicelogin..."
                try { Start-Process "https://microsoft.com/devicelogin" } catch { Write-Warning "Could not open browser automatically." }

                Connect-MgGraph -UseDeviceCode -Scopes "Domain.Read.All","Organization.Read.All" -NoWelcome -ErrorAction Stop 2>&1 | ForEach-Object {
                    if ($_ -match "enter the code ([A-Z0-9]+) to authenticate") {
                        Write-Host -ForegroundColor Cyan "`nDevice Code: $($matches[1])`n"
                    }
                    elseif ($_ -notmatch "To sign in, use a web browser") {
                        Write-Host $_
                    }
                }
            }
        }
        $context = Get-MgContext
        if (-not $context) {
            throw "Failed to establish Microsoft Graph connection."
        }
        Write-Host -ForegroundColor $processmessagecolor "Successfully connected to Microsoft Graph"
        Write-Host -ForegroundColor $processmessagecolor "Account = $($context.Account)"
        Write-Host -ForegroundColor $processmessagecolor "TenantId = $($context.TenantId)"
    }
    catch {
        Write-Host -ForegroundColor $errormessagecolor "`n[003] - Unable to connect to Microsoft Graph - $($_.Exception.Message)`n"
        Write-Host -ForegroundColor $warningmessagecolor "If you're experiencing authentication issues, try:"
        Write-Host -ForegroundColor $warningmessagecolor "  1. Run: Disconnect-MgGraph"
        Write-Host -ForegroundColor $warningmessagecolor "  2. Clear browser cache/cookies"
        Write-Host -ForegroundColor $warningmessagecolor "  3. Try device code flow: Use the -UseDeviceCode parameter`n"
        if ($debug) {
            Stop-Transcript | Out-Null
        }
        exit 3
    }
}
else {
    Write-Host -ForegroundColor $processmessagecolor "Existing Microsoft Graph connection detected"
    Write-Host -ForegroundColor $processmessagecolor "Account = $($context.Account)"
}
write-host -foregroundcolor $processmessagecolor "Connected to Microsoft 365 Admin service"

## Auto detect SharePoint Online admin domain
write-host -foregroundcolor $processmessagecolor "Determining SharePoint Administration URL"
try {
    # Try to get tenant name from the current context or organization
    $tenantname = $null
    
    # Method 1: Try getting from context account
    if ($context.Account) {
        $accountParts = $context.Account -split '@'
        if ($accountParts.Count -eq 2) {
            $domain = $accountParts[1]
            if ($domain -match '^(.+)\.onmicrosoft\.com$') {
                $tenantname = $matches[1]
                Write-Host -ForegroundColor $processmessagecolor "Tenant name from context account: $tenantname"
            }
            elseif ($domain -notlike "*.onmicrosoft.com") {
                # If not an onmicrosoft.com domain, we need to query organization
                Write-Host -ForegroundColor $processmessagecolor "Custom domain detected: $domain"
            }
        }
    }
    
    # Method 2: Use Graph cmdlet to get organization details
    if (-not $tenantname) {
        try {
            Write-Host -ForegroundColor $processmessagecolor "Querying organization using Get-MgOrganization..."
            
            # Need to import the Organizations module if not already loaded
            if (-not (Get-Module -Name Microsoft.Graph.Identity.DirectoryManagement)) {
                $ModuleInfo = Get-Module -ListAvailable Microsoft.Graph.Identity.DirectoryManagement | Where-Object Version -eq $TargetVersion | Select-Object -First 1
                if ($ModuleInfo) {
                    Import-Module $ModuleInfo.Path -Force -ErrorAction Stop
                }
            }
            
            $org = Get-MgOrganization -ErrorAction Stop
            if ($org.VerifiedDomains) {
                $onmicrosoftDomain = $org.VerifiedDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" -and $_.IsInitial -eq $true } | Select-Object -First 1
                if ($onmicrosoftDomain) {
                    $onname = $onmicrosoftDomain.Name.split(".")
                    $tenantname = $onname[0]
                    Write-Host -ForegroundColor $processmessagecolor "Tenant name from Get-MgOrganization: $tenantname"
                }
            }
        }
        catch {
            Write-Host -ForegroundColor $warningmessagecolor "Unable to get organization using cmdlet: $_"
        }
    }
    
    # Method 3: Use REST API to get organization details
    if (-not $tenantname) {
        try {
            Write-Host -ForegroundColor $processmessagecolor "Querying organization via REST API..."
            $orgResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction Stop
            
            if ($orgResponse.value -and $orgResponse.value.Count -gt 0) {
                $org = $orgResponse.value[0]
                if ($org.verifiedDomains) {
                    $onmicrosoftDomain = $org.verifiedDomains | Where-Object { $_.name -like "*.onmicrosoft.com" -and $_.isInitial -eq $true } | Select-Object -First 1
                    if ($onmicrosoftDomain) {
                        $onname = $onmicrosoftDomain.name.split(".")
                        $tenantname = $onname[0]
                        Write-Host -ForegroundColor $processmessagecolor "Tenant name from REST API: $tenantname"
                    }
                }
            }
        }
        catch {
            Write-Host -ForegroundColor $warningmessagecolor "Unable to get organization details via REST API: $_"
        }
    }
    
    # Method 4: Manual input as last resort
    if (-not $tenantname) {
        if (-not $noprompt) {
            Write-Host -ForegroundColor $warningmessagecolor "`nUnable to auto-detect tenant name."
            Write-Host -ForegroundColor $processmessagecolor "Please enter your tenant name (the part before .onmicrosoft.com)"
            Write-Host -ForegroundColor $processmessagecolor "Example: If your domain is contoso.onmicrosoft.com, enter 'contoso'"
            do {
                $tenantname = Read-Host -Prompt "Tenant name"
            } until (-not [string]::IsNullOrWhiteSpace($tenantname))
            Write-Host -ForegroundColor $processmessagecolor "Using manually entered tenant name: $tenantname"
        }
        else {
            throw "Unable to determine tenant name automatically and -noprompt is enabled."
        }
    }
    
    if (-not $tenantname) {
        throw "Unable to determine tenant name. Please ensure you have the necessary permissions."
    }
    
    $tenanturl = "https://" + $tenantname + "-admin.sharepoint.com"
    Write-host -ForegroundColor $processmessagecolor "SharePoint admin URL = $tenanturl"
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[004] - Unable to determine SharePoint admin URL"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null
    }
    exit 4
}

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