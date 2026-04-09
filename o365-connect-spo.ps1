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

function Invoke-GraphRequestWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("GET", "POST", "PATCH", "PUT", "DELETE")]
        [string]$Method,
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [object]$Body,
        [string]$ContentType = "application/json",
        [int]$MaxRetries = 3
    )

    for ($i = 0; $i -lt $MaxRetries; $i++) {
        try {
            if ($PSBoundParameters.ContainsKey('Body')) {
                return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $Body -ContentType $ContentType -OutputType PSObject -ErrorAction Stop
            }

            return Invoke-MgGraphRequest -Method $Method -Uri $Uri -OutputType PSObject -ErrorAction Stop
        }
        catch {
            $isRateLimited = $_.Exception.Message -match "429|TooManyRequests|throttle"
            $isRetryable = $_.Exception.Message -match "500|502|503|504|ServiceUnavailable|GatewayTimeout"

            if (($isRateLimited -or $isRetryable) -and $i -lt ($MaxRetries - 1)) {
                $waitTime = [math]::Pow(2, $i) * 5
                Write-Host -ForegroundColor $warningmessagecolor "Graph API transient error detected. Waiting $waitTime seconds before retry"
                Start-Sleep -Seconds $waitTime
            }
            else {
                throw
            }
        }
    }
}

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

# Microsoft Graph Authentication module
if (get-module -listavailable -name Microsoft.Graph.Authentication) {
    write-host -ForegroundColor $processmessagecolor "Microsoft.Graph.Authentication module installed"
}
else {
    write-host -ForegroundColor $warningmessagecolor -backgroundcolor $errormessagecolor "[001] - Microsoft.Graph.Authentication module not installed`n"
    if (-not $noprompt) {
        do {
            $response = read-host -Prompt "`nDo you wish to install the Microsoft.Graph.Authentication module (Y/N)?"
        } until (-not [string]::isnullorempty($response))
        if ($result -eq 'Y' -or $result -eq 'y') {
            write-host -foregroundcolor $processmessagecolor "Installing Microsoft.Graph.Authentication module - Administration escalation required"
            Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name Microsoft.Graph.Authentication -Force -confirm:$false" -wait -WindowStyle Hidden
            write-host -foregroundcolor $processmessagecolor "Microsoft.Graph.Authentication module installed"
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
        write-host -foregroundcolor $processmessagecolor "Installing Microsoft.Graph.Authentication module - Administration escalation required"
        Start-Process powershell -Verb runAs -ArgumentList "install-Module -Name Microsoft.Graph.Authentication -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "Microsoft.Graph.Authentication module installed"
    }
}
if (-not $noupdate) {
    write-host -foregroundcolor $processmessagecolor "Check whether newer version of Microsoft.Graph.Authentication module is available"
    #get version of the module (selects the first if there are more versions installed)
    $version = (Get-InstalledModule -name Microsoft.Graph.Authentication) | Sort-Object Version -Descending  | Select-Object Version -First 1
    #get version of the module in psgallery
    $psgalleryversion = Find-Module -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object Version -First 1
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
                $response = read-host -Prompt "`nDo you wish to update the Microsoft.Graph.Authentication module (Y/N)?"
            } until (-not [string]::isnullorempty($response))
            if ($response -eq 'Y' -or $response -eq 'y') {
                write-host -foregroundcolor $processmessagecolor "Updating Microsoft.Graph.Authentication module - Administration escalation required"
                Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name Microsoft.Graph.Authentication -Force -confirm:$false" -wait -WindowStyle Hidden
                write-host -foregroundcolor $processmessagecolor "Microsoft.Graph.Authentication module updated"
            }
            else {
                write-host -foregroundcolor $processmessagecolor "Microsoft.Graph.Authentication module not updated"
            }
        }
        else {
        write-host -foregroundcolor $processmessagecolor "Updating Microsoft.Graph.Authentication module - Administration escalation required"
        Start-Process powershell -Verb runAs -ArgumentList "update-Module -Name Microsoft.Graph.Authentication -Force -confirm:$false" -wait -WindowStyle Hidden
        write-host -foregroundcolor $processmessagecolor "Microsoft.Graph.Authentication module updated"
        }
    }
}

write-host -foregroundcolor $processmessagecolor "Microsoft.Graph.Authentication module loading"
try {
    Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop
}
catch {
    Write-Host -ForegroundColor $errormessagecolor "[002] - Unable to load Microsoft.Graph.Authentication module`n"
    Write-Host -ForegroundColor $errormessagecolor $_.Exception.Message
    if ($debug) {
        Stop-Transcript | Out-Null
    }
    exit 2
}
write-host -foregroundcolor $processmessagecolor "Microsoft.Graph.Authentication module loaded"

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
    
    # Method 2: Use REST API to get organization details
    if (-not $tenantname) {
        try {
            Write-Host -ForegroundColor $processmessagecolor "Querying organization via REST API..."
            $orgResponse = Invoke-GraphRequestWithRetry -Method GET -Uri "https://graph.microsoft.com/v1.0/organization"
            
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
    
    # Method 3: Manual input as last resort
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