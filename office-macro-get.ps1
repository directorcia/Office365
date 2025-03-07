<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/office-macro-get.ps1

Description - Report Office Macros Settings

Prerequisites - PowerShell 5.1 or later
                Office 365 ProPlus or Office 2016 or later
                Microsoft Graph PowerShell module (optional)

References:

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

function Test-RegistryKey {
    param (
        [string]$Path,
        [string]$Name
    )
    
    if (Test-Path $Path) {
        $value = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
        if ($null -ne $value) {
            return $value.$Name
        }
    }
    return $null
}

function amiadmin {
    # Check for elevated permissions
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    If (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $result = $false        # Not admin
    }
    else {
        $result = $true         # admin
    }
    return($result)
}

# Main script
Clear-Host
write-host -foregroundcolor $systemmessagecolor "Script started - Report Office Macros Settings`n"
write-host -foregroundcolor $processmessagecolor "[INFO] Test for elevated priviledges`n"
if (-not(amiadmin)) {
    write-host -foregroundcolor $errormessagecolor "*** ERROR *** - Please re-run PowerShell environment as Administrator`n"
    exit 1
} else {
    write-host -foregroundcolor $systemmessagecolor "Elevated priviledges confirmed`n"
}
write-host -foregroundcolor $processmessagecolor "[Info] = Checking PowerShell version"
$ps = $PSVersionTable.PSVersion
Write-host -foregroundcolor $processmessagecolor "- Detected supported PowerShell version: $($ps.Major).$($ps.Minor)`n"
# Initialize results array
$results = @()

# Define Office applications to check
$officeApps = @("Access", "Excel", "PowerPoint", "Word", "Outlook", "Publisher", "Visio")

# Check VBA macro settings
Write-Host "Checking Office VBA Macro Security Settings..." -ForegroundColor $systemmessagecolor

foreach ($app in $officeApps) {
    # Path for Office 365 / Office 2016+
    $vbaPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\$app\Security"
    $vbaPathUser = "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\$app\Security"
    $vbaPathDefault = "HKCU:\SOFTWARE\Microsoft\Office\16.0\$app\Security"
    
    # BlockContentExecutionFromInternet (1 = blocked)
    $blockInternet = Test-RegistryKey -Path $vbaPath -Name "BlockContentExecutionFromInternet"
    if ($null -eq $blockInternet) {
        $blockInternet = Test-RegistryKey -Path $vbaPathUser -Name "BlockContentExecutionFromInternet"
    }
    
    # VBAWarnings (1 = enable all, 2 = disable with notification, 3 = disable except digitally signed, 4 = disable all)
    $vbaWarnings = Test-RegistryKey -Path $vbaPath -Name "VBAWarnings"
    if ($null -eq $vbaWarnings) {
        $vbaWarnings = Test-RegistryKey -Path $vbaPathUser -Name "VBAWarnings"
    }
    if ($null -eq $vbaWarnings) {
        $vbaWarnings = Test-RegistryKey -Path $vbaPathDefault -Name "VBAWarnings"
    }
    
    # AccessVBOM (0 = no access to VBA object model, 1 = access allowed)
    $accessVBOM = Test-RegistryKey -Path $vbaPath -Name "AccessVBOM"
    if ($null -eq $accessVBOM) {
        $accessVBOM = Test-RegistryKey -Path $vbaPathUser -Name "AccessVBOM"
    }
    if ($null -eq $accessVBOM) {
        $accessVBOM = Test-RegistryKey -Path $vbaPathDefault -Name "AccessVBOM"
    }
    
    # Determine macro status
    $macroStatus = "Unknown"
    $blockStatus = "Unknown"
    $vbaOMStatus = "Unknown"
    
    if ($null -ne $vbaWarnings) {
        switch ($vbaWarnings) {
            1 { $macroStatus = "Enabled (All macros enabled - NOT SECURE)" }
            2 { $macroStatus = "Disabled with notification" }
            3 { $macroStatus = "Disabled except digitally signed macros" }
            4 { $macroStatus = "Disabled (All macros disabled without notification)" }
            default { $macroStatus = "Default settings" }
        }
    } else {
        $macroStatus = "Default settings (Disabled with notification)"
    }
    
    if ($null -ne $blockInternet) {
        $blockStatus = if ($blockInternet -eq 1) { "Blocked" } else { "Allowed" }
    } else {
        $blockStatus = "Default (Allowed)"
    }
    
    if ($null -ne $accessVBOM) {
        $vbaOMStatus = if ($accessVBOM -eq 1) { "Allowed" } else { "Restricted" }
    } else {
        $vbaOMStatus = "Default (Restricted)"
    }
    
    $results += [PSCustomObject]@{
        Application = $app
        MacroStatus = $macroStatus
        InternetMacros = $blockStatus
        VBAObjectModel = $vbaOMStatus
    }
}

# Check for system-wide trusted locations
$trustedLocationPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security\Trusted Locations"
$trustedLocationPathUser = "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security\Trusted Locations"

$allowUserLocations = Test-RegistryKey -Path $trustedLocationPath -Name "AllowUserLocations"
if ($null -eq $allowUserLocations) {
    $allowUserLocations = Test-RegistryKey -Path $trustedLocationPathUser -Name "AllowUserLocations"
}

# Check for trust access to VBA project
$trustAccessPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security"
$trustAccessPathUser = "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security"

$trustAccess = Test-RegistryKey -Path $trustAccessPath -Name "TrustAccessVBOM"
if ($null -eq $trustAccess) {
    $trustAccess = Test-RegistryKey -Path $trustAccessPathUser -Name "TrustAccessVBOM"
}

# Check for macro blocking from Office 2016+ policy
$macroBlockPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security"
$macroBlockPathUser = "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security"

$macroBlock = Test-RegistryKey -Path $macroBlockPath -Name "MacroRuntimeScanScope"
if ($null -eq $macroBlock) {
    $macroBlock = Test-RegistryKey -Path $macroBlockPathUser -Name "MacroRuntimeScanScope"
}

# Check for Block Macro Group Policy setting (newer M365 Apps feature)
$blockMacroPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security"
$blockMacroPathUser = "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security"

$blockMacrosWithNotification = Test-RegistryKey -Path $blockMacroPath -Name "BlockMacrosWithNotification"
if ($null -eq $blockMacrosWithNotification) {
    $blockMacrosWithNotification = Test-RegistryKey -Path $blockMacroPathUser -Name "BlockMacrosWithNotification"
}

# Check for AMSI scanning
$amsiPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security"
$amsiPathUser = "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security"

$amsiScan = Test-RegistryKey -Path $amsiPath -Name "MacroRuntimeScanScope"
if ($null -eq $amsiScan) {
    $amsiScan = Test-RegistryKey -Path $amsiPathUser -Name "MacroRuntimeScanScope"
}

# Display results
Write-Host "`nMacro Security Settings by Application:" -ForegroundColor $processmessagecolor
$results | Format-Table -AutoSize

Write-Host "`nGlobal Office 365 Macro Settings:" -ForegroundColor $processmessagecolor
Write-Host "User Trusted Locations: $(if ($null -eq $allowUserLocations) {"Default (Allowed)"} elseif ($allowUserLocations -eq 1) {"Allowed"} else {"Disallowed"})"
Write-Host "Trust Access to VBA Project: $(if ($null -eq $trustAccess) {"Default (Disabled)"} elseif ($trustAccess -eq 1) {"Enabled"} else {"Disabled"})"
Write-Host "AMSI Macro Scanning: $(if ($null -eq $amsiScan) {"Default"} elseif ($amsiScan -eq 2) {"Enabled for All Macros"} elseif ($amsiScan -eq 1) {"Enabled for Downloaded Files Only"} else {"Disabled"})"

Write-Host "`nSummary:" -ForegroundColor $warningmessagecolor
$totalDisabled = ($results | Where-Object { $_.MacroStatus -like "*Disabled*" }).Count
$totalEnabled = ($results | Where-Object { $_.MacroStatus -like "*Enabled*" }).Count
$totalDefault = ($results | Where-Object { $_.MacroStatus -like "*Default*" }).Count
$totalBlocked = ($results | Where-Object { $_.InternetMacros -eq "Blocked" }).Count

Write-Host "$totalDisabled out of $($results.Count) Office applications have macros explicitly disabled"
Write-Host "$totalEnabled out of $($results.Count) Office applications have macros explicitly enabled"
Write-Host "$totalDefault out of $($results.Count) Office applications use default macro settings"
Write-Host "$totalBlocked out of $($results.Count) Office applications block macros from the internet"

if ($totalEnabled -gt 0 -or ($totalDefault -gt 0 -and $totalBlocked -eq 0)) {
    Write-Host "`nWARNING: Macros appear to be enabled in some applications. This may pose a security risk." -ForegroundColor $errormessagecolor
} else {
    Write-Host "`nMacros appear to be properly restricted across Office applications." -ForegroundColor $processmessagecolor
}

# Check for Office 365 tenant-wide settings via Microsoft Graph (requires Graph modules)
if (Get-Module -ListAvailable -Name "Microsoft.Graph.Identity.SignIns") {
    try {
        Write-Host "`nChecking Microsoft 365 tenant-wide settings ...`n" -ForegroundColor $systemmessagecolor
        Write-Host "[INFO] Attempt to connect to Microsoft Graph with Policy.Read.All" -ForegroundColor $processmessagecolor
        try {
            Disconnect-Graph -ErrorAction SilentlyContinue | Out-Null
        } catch {
            # Ignore errors if not connected
        }
        try {
            Connect-MgGraph -Scopes "Policy.Read.All" -NoWelcome -ErrorAction Stop | Out-Null
        } catch {
            Write-Host "Could not connect to Microsoft Graph: $($_.Exception.Message)`n" -ForegroundColor $errormessagecolor
            return
        }
        $context=get-mgcontext
        write-host -foregroundcolor $processmessagecolor "[INFO] Connected to Microsoft Graph as:",$context.account
        
        $macroPolicy = Get-MgPolicyAuthorizationPolicy | Select-Object -ExpandProperty DefaultUserRolePermissions
        if ($macroPolicy -match "AllowedToCreateApps") {
            Write-Host "`nUsers can register apps that use Office add-ins and macros: Enabled" -ForegroundColor $warningmessagecolor
        } else {
            Write-Host "`nUsers can register apps that use Office add-ins and macros: Disabled" -ForegroundColor $processmessagecolor
        }
    } catch {
        Write-Host "Could not check tenant-wide settings: $($_.Exception.Message)" -ForegroundColor $errormessagecolor
        Write-Host "Note: To check tenant-wide settings, install the Microsoft.Graph modules and have sufficient permissions." -ForegroundColor $warningmessagecolor
    }
}
Write-Host "`nScript completed`n" -ForegroundColor $systemmessagecolor