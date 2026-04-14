<#
================================================================================
  IMPORTANT: PowerShell 7 Required
================================================================================

If you receive an error about "#requires" or PowerShell version when running
this script, you are using Windows PowerShell 5.1.

This script REQUIRES PowerShell 7 or higher.

QUICK INSTALL (choose one):
    • WinGet:          winget install Microsoft.PowerShell
    • Direct download: https://aka.ms/powershell-release?tag=stable
    • Microsoft Store: Search "PowerShell"

After installing, launch "PowerShell 7" (not "Windows PowerShell") and run
this script again.

Why PowerShell 7?
    • Required for modern Microsoft Graph and Azure modules
    • Better performance, security, and cross-platform support
    • Actively maintained by Microsoft

================================================================================
#>
#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Prompt before installing missing modules")]
    [switch]$Prompt,

    [Parameter(HelpMessage = "Create a transcript log of all operations")]
    [switch]$CreateLog,

    [Parameter(HelpMessage = "Path for log file (default: current directory)")]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Container)) {
            throw "LogPath must be a valid directory path"
        }
        return $true
    })]
    [string]$LogPath = (Get-Location).Path,

    [Parameter(HelpMessage = "Skip deprecated module cleanup")]
    [switch]$SkipDeprecatedCleanup,

    [Parameter(HelpMessage = "Specify which modules to process (default: all)")]
    [ValidateSet('All', 'Graph', 'Azure', 'Exchange', 'Teams', 'SharePoint', 'PowerApps')]
    [string[]]$ModuleScope = @('All'),

    [Parameter(HelpMessage = "Use alternative PowerShell Gallery repository")]
    [string]$Repository = 'PSGallery'
)

$ErrorActionPreference = 'Stop'

$Script:Colors = @{
    System  = 'Cyan'
    Process = 'Green'
    Warning = 'Yellow'
    Error   = 'Red'
    Info    = 'White'
}

$Script:ModuleList = @(
    @{ Name = 'Microsoft.Graph'; Description = 'Microsoft Graph PowerShell SDK'; Category = 'Graph'; RequiresSpecialHandling = $false },
    @{ Name = 'Microsoft.Graph.Authentication'; Description = 'Microsoft Graph Authentication'; Category = 'Graph'; RequiresSpecialHandling = $false },
    @{ Name = 'MicrosoftTeams'; Description = 'Microsoft Teams PowerShell Module'; Category = 'Teams'; RequiresSpecialHandling = $false },
    @{ Name = 'ExchangeOnlineManagement'; Description = 'Exchange Online PowerShell V3 Module'; Category = 'Exchange'; RequiresSpecialHandling = $false },
    @{ Name = 'Az'; Description = 'Azure PowerShell Module'; Category = 'Azure'; RequiresSpecialHandling = $false },
    @{ Name = 'PnP.PowerShell'; Description = 'SharePoint PnP PowerShell Module'; Category = 'SharePoint'; RequiresSpecialHandling = $false },
    @{ Name = 'Microsoft.PowerApps.PowerShell'; Description = 'PowerApps PowerShell Module'; Category = 'PowerApps'; RequiresSpecialHandling = $false },
    @{ Name = 'Microsoft.PowerApps.Administration.PowerShell'; Description = 'PowerApps Administration PowerShell Module'; Category = 'PowerApps'; RequiresSpecialHandling = $false },
    @{ Name = 'PowershellGet'; Description = 'PowerShellGet Module'; Category = 'Core'; RequiresSpecialHandling = $true },
    @{ Name = 'PackageManagement'; Description = 'Package Management Module'; Category = 'Core'; RequiresSpecialHandling = $true },
    @{ Name = 'Microsoft.Online.SharePoint.PowerShell'; Description = 'SharePoint Online Management Shell'; Category = 'SharePoint'; RequiresSpecialHandling = $false },
    @{ Name = 'Microsoft.WinGet.Client'; Description = 'Windows Package Manager Client'; Category = 'Other'; RequiresSpecialHandling = $false }
)

$Script:DeprecatedModules = @(
    @{ Name = 'AzureAD'; Replacement = 'Microsoft.Graph'; Reason = 'AzureAD module is deprecated.' },
    @{ Name = 'AzureADPreview'; Replacement = 'Microsoft.Graph'; Reason = 'AzureAD Preview module is deprecated.' },
    @{ Name = 'MSOnline'; Replacement = 'Microsoft.Graph'; Reason = 'MSOnline module is deprecated.' },
    @{ Name = 'AIPService'; Replacement = 'Microsoft.Graph'; Reason = 'AIPService is replaced by Graph APIs.' },
    @{ Name = 'aadrm'; Replacement = 'Microsoft.Graph'; Reason = 'AADRM support ended July 15, 2020.' },
    @{ Name = 'SharePointPnPPowerShellOnline'; Replacement = 'PnP.PowerShell'; Reason = 'Legacy PnP module replaced by PnP.PowerShell.' },
    @{ Name = 'WindowsAutoPilotIntune'; Replacement = 'Microsoft.Graph'; Reason = 'Functionality moved to Graph SDK.' },
    @{ Name = 'O365CentralizedAddInDeployment'; Replacement = 'ExchangeOnlineManagement'; Reason = 'Functionality integrated into Exchange Online Management.' },
    @{ Name = 'MSCommerce'; Replacement = 'Microsoft.Graph'; Reason = 'Commerce functionality is available through Graph.' }
)

function Write-ColorOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('System', 'Process', 'Warning', 'Error', 'Info')]
        [string]$Type = 'Info'
    )

    Write-Host $Message -ForegroundColor $Script:Colors[$Type]
}

function Write-ProgressHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter()]
        [string]$Subtitle = ''
    )

    $separator = '=' * 80
    Write-ColorOutput $separator -Type System
    Write-ColorOutput $Title -Type System
    if ($Subtitle) {
        Write-ColorOutput $Subtitle -Type Info
    }
    Write-ColorOutput $separator -Type System
}

function Test-Administrator {
    [CmdletBinding()]
    param()

    try {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Ensure-PackageProvider {
    [CmdletBinding()]
    param()

    $provider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if (-not $provider) {
        Write-ColorOutput 'Installing NuGet package provider...' -Type Process
        Install-PackageProvider -Name NuGet -Force -Confirm:$false | Out-Null
    }
    else {
        Write-ColorOutput "NuGet package provider already installed ($($provider.Version))." -Type Info
    }
}

function Ensure-Repository {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $repo = Get-PSRepository -Name $Name -ErrorAction SilentlyContinue
    if (-not $repo) {
        throw "Repository '$Name' was not found."
    }

    if ($repo.InstallationPolicy -ne 'Trusted') {
        if ($Prompt) {
            $response = Read-Host "Repository '$Name' is Untrusted. Set to Trusted (Y/N)?"
            if ($response -notmatch '^[Yy]') {
                Write-ColorOutput "Continuing with Untrusted repository '$Name'." -Type Warning
                return
            }
        }

        Set-PSRepository -Name $Name -InstallationPolicy Trusted
        Write-ColorOutput "Repository '$Name' set to Trusted." -Type Process
    }
}

function Get-FilteredModuleList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Scope
    )

    if ($Scope -contains 'All') {
        return $Script:ModuleList
    }

    $selected = foreach ($scopeItem in $Scope) {
        $Script:ModuleList | Where-Object { $_.Category -eq $scopeItem }
    }

    return $selected | Sort-Object Name -Unique
}

function Remove-DeprecatedModules {
    [CmdletBinding()]
    param()

    foreach ($deprecatedModule in $Script:DeprecatedModules) {
        $installed = Get-Module -ListAvailable -Name $deprecatedModule.Name
        if (-not $installed) {
            continue
        }

        Write-ColorOutput "Removing deprecated module: $($deprecatedModule.Name)" -Type Warning
        Write-ColorOutput "  Reason: $($deprecatedModule.Reason)" -Type Warning
        Write-ColorOutput "  Replacement: $($deprecatedModule.Replacement)" -Type Info

        try {
            Uninstall-Module -Name $deprecatedModule.Name -AllVersions -Force -Confirm:$false -ErrorAction Stop
            Write-ColorOutput "  Removed $($deprecatedModule.Name)." -Type Process
        }
        catch {
            Write-ColorOutput "  Could not fully remove $($deprecatedModule.Name): $($_.Exception.Message)" -Type Warning
        }
    }
}

function Install-OrUpdateModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Module,

        [Parameter(Mandatory)]
        [string]$RepositoryName
    )

    $name = $Module.Name
    $installed = Get-InstalledModule -Name $name -ErrorAction SilentlyContinue

    if ($Prompt) {
        $action = if ($installed) { 'Update' } else { 'Install' }
        $response = Read-Host "$action module '$name' (Y/N)?"
        if ($response -notmatch '^[Yy]') {
            Write-ColorOutput "Skipped module: $name" -Type Warning
            return $true
        }
    }

    $installParams = @{
        Name        = $name
        Repository  = $RepositoryName
        Scope       = 'AllUsers'
        Force       = $true
        Confirm     = $false
        ErrorAction = 'Stop'
    }

    $updateParams = @{
        Name        = $name
        Force       = $true
        Confirm     = $false
        ErrorAction = 'Stop'
    }

    if ($Module.RequiresSpecialHandling) {
        $installParams.AllowClobber = $true
    }

    if ($installed) {
        Write-ColorOutput "Updating module: $name" -Type Process
        Update-Module @updateParams
    }
    else {
        Write-ColorOutput "Installing module: $name" -Type Process
        Install-Module @installParams
    }

    return $true
}

try {
    # Early elevation check (friendlier than the generic #Requires message)
    try {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        $isAdmin = $false
    }

    if ($CreateLog) {
        $logPath = Join-Path $LogPath "o365-setup-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
        Start-Transcript -Path $logPath -Append
        Write-ColorOutput "📋 Logging to: $logPath" -Type System
    }

    Write-ProgressHeader "Microsoft Cloud PowerShell Module Updater v2.9" "Enhanced version with improved module management and version cleanup"

    if (-not $isAdmin) {
        $scriptPath = $PSCommandPath
        $pathForDisplay = ($scriptPath -replace "'", "''")
        $elevateCmd = "Start-Process -Verb RunAs pwsh -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File', '$pathForDisplay')"

        Write-ColorOutput "❌ Administrator privileges are required." -Type Error
        Write-ColorOutput "This updater makes system-wide module changes (AllUsers scope)." -Type Info
        Write-ColorOutput "How to continue:" -Type Info
        Write-ColorOutput "  1) Close this window and launch PowerShell 7 as Administrator" -Type Info
        Write-ColorOutput "     - Start menu: search 'PowerShell 7', right‑click > Run as administrator" -Type Info
        Write-ColorOutput "     - Windows Terminal: open a new Administrator window" -Type Info
        Write-ColorOutput "  2) Or copy/paste to relaunch elevated:" -Type Info
        Write-ColorOutput "     $elevateCmd" -Type Process
        Write-ColorOutput "" -Type Info
        Write-ColorOutput "Tip: Append your parameters after the -File path if needed." -Type Info
        exit 1
    }

    $psVersion = $PSVersionTable.PSVersion
    Write-ColorOutput "PowerShell Version: $($psVersion.Major).$($psVersion.Minor).$($psVersion.Build)" -Type Info
    Write-ColorOutput 'Configuration:' -Type Info
    Write-ColorOutput "  Scope: $($ModuleScope -join ', ')" -Type Info
    Write-ColorOutput "  Repository: $Repository" -Type Info
    Write-ColorOutput "  Prompt: $Prompt" -Type Info
    Write-ColorOutput "  Skip Deprecated Cleanup: $SkipDeprecatedCleanup" -Type Info
    Write-ColorOutput "  Create Log: $CreateLog" -Type Info

    Ensure-PackageProvider
    Ensure-Repository -Name $Repository

    if (-not $SkipDeprecatedCleanup) {
        Write-ColorOutput 'Checking and removing deprecated modules...' -Type System
        Remove-DeprecatedModules
    }
    else {
        Write-ColorOutput 'Skipping deprecated module cleanup.' -Type Warning
    }

    $modulesToProcess = Get-FilteredModuleList -Scope $ModuleScope
    if (-not $modulesToProcess -or $modulesToProcess.Count -eq 0) {
        throw 'No modules selected for processing. Check ModuleScope.'
    }

    Write-ColorOutput "Processing $($modulesToProcess.Count) module(s)..." -Type System

    $successCount = 0
    $failureCount = 0

    foreach ($module in $modulesToProcess) {
        try {
            $completed = Install-OrUpdateModule -Module $module -RepositoryName $Repository
            if ($completed) {
                $successCount++
            }
        }
        catch {
            $failureCount++
            Write-ColorOutput "Failed module '$($module.Name)': $($_.Exception.Message)" -Type Error
        }
    }

    Write-ProgressHeader "Script Completed Successfully"
    Write-ColorOutput "✅ Microsoft Cloud PowerShell Module Updater completed successfully!" -Type Process
    Write-ColorOutput "End Time: $(Get-Date -Format (Get-Culture).DateTimeFormat.FullDateTimePattern)" -Type System
    Write-ColorOutput "" -Type Info
    Write-ColorOutput "💡 Recommendations:" -Type Info
    Write-ColorOutput "  • Restart PowerShell to use updated modules" -Type Info
    Write-ColorOutput "  • Run 'Get-Module -ListAvailable' to verify installations" -Type Info
    Write-ColorOutput "  • Check module documentation for any breaking changes" -Type Info

    if ($failureCount -gt 0) {
        exit 2
    }
}
catch {
    Write-ColorOutput "💥 Fatal Error: $($_.Exception.Message)" -Type Error
    Write-ColorOutput "Stack Trace:" -Type Error
    Write-ColorOutput $_.ScriptStackTrace -Type Error

    exit 1
}
finally {
    # Stop transcript if it was started
    if ($CreateLog) {
        try {
            Stop-Transcript
            Write-ColorOutput "📋 Log file saved successfully" -Type System
        }
        catch {
            # Transcript might not be running
        }
    }
}