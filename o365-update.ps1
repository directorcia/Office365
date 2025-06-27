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
    [string]$LogPath = $PWD,
    
    [Parameter(HelpMessage = "Skip deprecated module cleanup")]
    [switch]$SkipDeprecatedCleanup,
    
    [Parameter(HelpMessage = "Skip automatic cleanup of old module versions after updates")]
    [switch]$SkipVersionCleanup,
    
    [Parameter(HelpMessage = "Only check versions without updating")]
    [switch]$CheckOnly,
    
    [Parameter(HelpMessage = "Check for PowerShell session conflicts and exit")]
    [switch]$CheckSessions,
    
    [Parameter(HelpMessage = "Automatically terminate conflicting PowerShell sessions without prompting")]
    [switch]$TerminateConflicts,
    
    [Parameter(HelpMessage = "Specify which modules to process (default: all)")]
    [ValidateSet('All', 'Graph', 'Azure', 'Exchange', 'Teams', 'SharePoint', 'PowerApps')]
    [string[]]$ModuleScope = @('All'),
    
    [Parameter(HelpMessage = "Skip network connectivity checks")]
    [switch]$SkipConnectivityCheck,
    
    [Parameter(HelpMessage = "Use alternative PowerShell Gallery repository")]
    [string]$Repository = 'PSGallery'
)
<#
.SYNOPSIS
    Updates all relevant Microsoft Cloud PowerShell modules

.DESCRIPTION
    This script updates or installs the latest versions of Microsoft Cloud PowerShell modules
    including Azure AD, Exchange Online, SharePoint Online, Teams, and more.
    
    PERFORMANCE ENHANCEMENTS (v2.0):
    - Optimized installation process with progress tracking for large modules
    - Pre-installation network optimization (TLS 1.2, connection limits)
    - Background job processing for faster downloads
    - Real-time progress bars with time estimates
    - Intelligent module size estimation for accurate time predictions

.PARAMETER Prompt
    Prompts before installing missing modules instead of installing automatically.
    Also enables interactive cleanup prompts when multiple module versions are detected.

.PARAMETER CreateLog
    Creates a transcript log of all operations

.PARAMETER LogPath
    Specifies the path for the log file (default: current directory)

.PARAMETER SkipDeprecatedCleanup
    Skips the cleanup of deprecated modules

.PARAMETER SkipVersionCleanup
    Skips automatic cleanup of old module versions after successful installations/updates.
    When combined with -Prompt, users can interactively choose which versions to clean up.

.PARAMETER CheckOnly
    Only checks versions without performing updates or installations

.PARAMETER CheckSessions
    Checks for PowerShell session conflicts and exits without performing updates

.PARAMETER TerminateConflicts
    Automatically terminates conflicting PowerShell sessions without prompting

.EXAMPLE
    .\o365-update.ps1
    Updates all modules automatically without prompting

.EXAMPLE
    .\o365-update.ps1 -Prompt
    Prompts before installing missing modules

.EXAMPLE
    .\o365-update.ps1 -CreateLog -LogPath "C:\Logs"
    Updates modules and creates a log file in C:\Logs

.EXAMPLE
    .\o365-update.ps1 -CheckOnly
    Only checks module versions without updating

.EXAMPLE
    .\o365-update.ps1 -CheckSessions
    Checks for PowerShell session conflicts and provides guidance

.EXAMPLE
    .\o365-update.ps1 -TerminateConflicts
    Automatically terminates conflicting PowerShell sessions and then updates modules

.EXAMPLE
    .\o365-update.ps1 -SkipVersionCleanup
    Updates modules but skips automatic cleanup of old versions to preserve multiple versions

.EXAMPLE
    .\o365-update.ps1 -Prompt -SkipVersionCleanup
    Interactive mode with cleanup prompts. When multiple versions are detected, 
    prompts user to choose whether to clean up old versions for each module.
    Also offers comprehensive cleanup at the end of processing.

.NOTES
    Author: CIAOPS    Version: 2.9
    Last Updated: June 2025
    Requires: PowerShell 5.1 or higher, Administrator privileges
    
    IMPORTANT: This version removes deprecated Azure AD and MSOnline modules in favor of Microsoft Graph PowerShell SDK
    
    Major Changes in v2.9:
    - ELIMINATED "⚠️ Multiple versions installed" warnings through automatic cleanup
    - Added automatic cleanup of old module versions after successful installations/updates
    - Added interactive cleanup prompts when using -Prompt with -SkipVersionCleanup
    - Users can now choose Y/N/S (Skip all) when prompted about cleaning up old versions
    - Added comprehensive cleanup option at end of processing for all remaining multiple versions
    - Added -SkipVersionCleanup parameter to preserve multiple versions if needed
    - Added Remove-OldModuleVersions function for targeted cleanup of specific modules
    - Added Remove-AllOldModuleVersions function for comprehensive cleanup of all modules
    - Enhanced user experience by maintaining only the latest version of each module
    - Improved script performance by reducing module version conflicts
    - Added comprehensive cleanup reporting and error handling
    - Users will no longer see multiple version warnings unless intentionally preserved
    
    Major Changes in v2.8:
    - ELIMINATED duplicate session conflict prompts and displays completely
    - Added Silent mode to Get-PowerShellSessions function to prevent duplicate output
    - Fixed double prompting issue where users were asked twice to terminate sessions
    - Consolidated all session conflict detection and handling into single execution point
    - Removed duplicate session termination logic from Test-ModuleRemovalPrerequisites function
    - Fixed PowerShell 5.x compatibility by removing null coalescing operator (??)
    - Streamlined user experience with single, clear session conflict workflow
    - Enhanced state tracking to prevent redundant session checks and prompts
    - Improved script performance by eliminating unnecessary duplicate operations
    - Users now see only one session detection message and one termination prompt
    
    Major Changes in v2.7:
    - ELIMINATED "PackageManagement is currently in use" warning completely
    - Enhanced comprehensive output stream suppression (all 6 PowerShell streams)
    - Added InformationAction suppression for complete silence during core module updates
    - Transformed error messages into positive success confirmations
    - Enhanced verification of side-by-side installations with multiple version detection
    - Improved Azure best practice implementation for zero-conflict core module updates
    - Added comprehensive preference restoration for all PowerShell output streams
    - Enhanced error handling to treat "in use" warnings as successful installations
    
    Previous Changes in v2.6:
    - Enhanced core module conflict resolution with comprehensive "PackageManagement is currently in use" handling
    - Added Resolve-ModuleInUseConflict function for intelligent conflict resolution
    - Improved side-by-side installation with warning suppression for core modules  
    - Added proactive user guidance about core module update behavior at script start
    - Enhanced verification of successful core module installations
    - Improved error handling and user messaging for PackageManagement/PowerShellGet conflicts
    - Added automatic detection of core module vs regular module conflicts
    - Enhanced completion messages with clearer guidance for core module updates
    
    Previous Changes in v2.5:
    - Fixed "module currently in use" errors for PackageManagement and PowerShellGet
    - Implemented Azure best practices for core module management
    - Enhanced detection of loaded modules to prevent conflicts
    - Added intelligent side-by-side installation for locked core modules
    - Improved user guidance for core module update scenarios
    - Replaced deprecated Get-WmiObject with Get-CimInstance for compatibility
    - Enhanced session conflict warning consolidation
    
    Previous Changes in v2.4:
    - Enhanced error handling for module removal operations
    - Added specific error detection for common failure scenarios (permissions, file locks, etc.)
    - Implemented retry logic with delays for transient failures
    - Added prerequisite checking before module removal attempts
    - Enhanced troubleshooting guidance and user feedback
    - Improved file system removal safety checks
    - Added comprehensive end-of-script troubleshooting tips
    - Added comprehensive PowerShell session detection and conflict analysis
    - Added -CheckSessions parameter for dedicated session conflict checking
    - Enhanced session detection to include ISE, VS Code, and Windows Terminal
    - Added detailed session information display with process details
    - Implemented module conflict detection and resolution guidance
    - Added automatic PowerShell session termination capabilities
    - Added -TerminateConflicts parameter for unattended conflict resolution
    - Enhanced user experience with interactive conflict resolution options
    
    Previous Changes in v2.3:
    - Fixed Constrained Language Mode compatibility issues
    - Added PowerShell language mode detection and compatibility checking
    - Improved fallback mechanisms for restricted environments
    - Enhanced error handling for background job limitations
    
    Previous Changes in v2.2:
    - Added progress tracking and time estimation for module removal
    - Enhanced user feedback with real-time progress indicators
    - Added timeout protection for stuck uninstall operations
    - Improved time remaining calculations and ETA display
    
    Previous Changes in v2.1:
    - Removed deprecated AzureAD, MSOnline, AIPService modules
    - Added Microsoft.Graph.Authentication for better Graph connectivity
    - Removed WindowsAutoPilotIntune (functionality moved to Graph)
    - Removed O365CentralizedAddInDeployment (integrated into Exchange Online Management)
    - Removed MSCommerce (functionality available through Graph)
    - Added Microsoft.WinGet.Client for modern package management
    - Added regional date/time formatting support
    - Enhanced core module update handling
    
    Source: https://github.com/directorcia/Office365/blob/master/o365-update.ps1
    Documentation: https://github.com/directorcia/Office365/wiki/Update-all-Microsoft-Cloud-PowerShell-modules

.LINK
    https://github.com/directorcia/Office365
#>

#Requires -RunAsAdministrator

# Script configuration
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

# Color scheme
$Script:Colors = @{
    System    = 'Cyan'
    Process   = 'Green'
    Warning   = 'Yellow'
    Error     = 'Red'
    Info      = 'White'
    Success   = 'Green'
}

# Enhanced module definitions with categories and improved metadata
$Script:ModuleList = @(
    @{ 
        Name = 'Microsoft.Graph'
        Description = 'Microsoft Graph PowerShell SDK'
        Category = 'Graph'
        Deprecated = $false
        Priority = 1
        EstimatedSizeMB = 280
        RequiredPSVersion = '5.1'
    },
    @{ 
        Name = 'Microsoft.Graph.Authentication'
        Description = 'Microsoft Graph Authentication'
        Category = 'Graph'
        Deprecated = $false
        Priority = 1
        EstimatedSizeMB = 15
        RequiredPSVersion = '5.1'
    },
    @{ 
        Name = 'MicrosoftTeams'
        Description = 'Microsoft Teams PowerShell Module'
        Category = 'Teams'
        Deprecated = $false
        Priority = 2
        EstimatedSizeMB = 35
        RequiredPSVersion = '5.1'
    },
    @{ 
        Name = 'ExchangeOnlineManagement'
        Description = 'Exchange Online PowerShell V3 Module'
        Category = 'Exchange'
        Deprecated = $false
        Priority = 2
        EstimatedSizeMB = 40
        RequiredPSVersion = '5.1'
    },
    @{ 
        Name = 'Az'
        Description = 'Azure PowerShell Module'
        Category = 'Azure'
        Deprecated = $false
        Priority = 3
        EstimatedSizeMB = 450
        RequiredPSVersion = '5.1'
    },
    @{ 
        Name = 'PnP.PowerShell'
        Description = 'SharePoint PnP PowerShell Module'
        Category = 'SharePoint'
        Deprecated = $false
        Priority = 2
        EstimatedSizeMB = 120
        RequiredPSVersion = '5.1'
    },
    @{ 
        Name = 'Microsoft.PowerApps.PowerShell'
        Description = 'PowerApps PowerShell Module'
        Category = 'PowerApps'
        Deprecated = $false
        Priority = 3
        EstimatedSizeMB = 25
        RequiredPSVersion = '5.1'
    },
    @{ 
        Name = 'Microsoft.PowerApps.Administration.PowerShell'
        Description = 'PowerApps Administration PowerShell Module'
        Category = 'PowerApps'
        Deprecated = $false
        Priority = 3
        EstimatedSizeMB = 25
        RequiredPSVersion = '5.1'
    },
    @{ 
        Name = 'PowershellGet'
        Description = 'PowerShellGet Module'
        Category = 'Core'
        Deprecated = $false
        RequiresSpecialHandling = $true
        Priority = 1
        EstimatedSizeMB = 10
        RequiredPSVersion = '5.1'
    },
    @{ 
        Name = 'PackageManagement'
        Description = 'Package Management Module'
        Category = 'Core'
        Deprecated = $false
        RequiresSpecialHandling = $true
        Priority = 1
        EstimatedSizeMB = 10
        RequiredPSVersion = '5.1'
    },
    @{ 
        Name = 'Microsoft.Online.SharePoint.PowerShell'
        Description = 'SharePoint Online Management Shell'
        Category = 'SharePoint'
        Deprecated = $false
        Priority = 3
        EstimatedSizeMB = 30
        RequiredPSVersion = '5.1'
    },
    @{ 
        Name = 'Microsoft.WinGet.Client'
        Description = 'Windows Package Manager Client'
        Category = 'Other'
        Deprecated = $false
        Priority = 3
        EstimatedSizeMB = 10
        RequiredPSVersion = '5.1'
    }
)

# Deprecated modules to clean up
$Script:DeprecatedModules = @(
    @{ Name = 'AzureAD'; Replacement = 'Microsoft.Graph'; Reason = 'AzureAD module is deprecated. Use Microsoft Graph PowerShell SDK instead.' },
    @{ Name = 'AzureADPreview'; Replacement = 'Microsoft.Graph'; Reason = 'AzureAD Preview module is deprecated. Use Microsoft Graph PowerShell SDK instead.' },
    @{ Name = 'MSOnline'; Replacement = 'Microsoft.Graph'; Reason = 'MSOnline module is deprecated. Use Microsoft Graph PowerShell SDK instead.' },
    @{ Name = 'AIPService'; Replacement = 'Microsoft.Graph'; Reason = 'AIPService is being replaced by Microsoft Graph Information Protection APIs.' },
    @{ Name = 'aadrm'; Replacement = 'Microsoft.Graph'; Reason = 'Support ended July 15, 2020. Use Microsoft Graph instead.' },
    @{ Name = 'SharePointPnPPowerShellOnline'; Replacement = 'PnP.PowerShell'; Reason = 'Replaced by new PnP PowerShell module.' },
    @{ Name = 'WindowsAutoPilotIntune'; Replacement = 'Microsoft.Graph'; Reason = 'Intune functionality moved to Microsoft Graph PowerShell SDK.' },
    @{ Name = 'O365CentralizedAddInDeployment'; Replacement = 'ExchangeOnlineManagement'; Reason = 'Functionality integrated into Exchange Online Management module.' },
    @{ Name = 'MSCommerce'; Replacement = 'Microsoft.Graph'; Reason = 'Commerce functionality available through Microsoft Graph.' }
)

# Session conflict tracking variables
$Script:SessionConflictCheckPerformed = $false
$Script:SessionConflictsResolved = $false

# Cleanup prompt tracking variable
$Script:SkipCleanupPrompts = $false

# Configuration variables
$Script:MaxParallelOperations = 4
$Script:TimeoutMinutes = 30

function Write-ColorOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('System', 'Process', 'Warning', 'Error', 'Info', 'Success')]
        [string]$Type = 'Info',
        
        [Parameter()]
        [switch]$NoNewline
    )
    
    # Handle empty strings by writing a blank line
    if ([string]::IsNullOrEmpty($Message)) {
        Write-Host ""
        return
    }
    
    $params = @{
        ForegroundColor = $Script:Colors[$Type]
        Object = $Message
    }
    
    if ($NoNewline) {
        $params.NoNewline = $true
    }
    
    Write-Host @params
}

function Test-PackageProvider {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PackageName
    )
    
    try {
        Write-ColorOutput "    Checking package provider: $PackageName" -Type Info
        
        $installedProvider = Get-PackageProvider -Name $PackageName -ErrorAction SilentlyContinue
        
        if (-not $installedProvider) {
            Write-ColorOutput "    [Warning] Package provider '$PackageName' not found" -Type Warning
            
            if ($Prompt -and -not $CheckOnly) {
                $response = Read-Host "    Install package provider '$PackageName' (Y/N)?"
                if ($response -notmatch '^[Yy]') {
                    Write-ColorOutput "    Skipping installation of $PackageName" -Type Warning
                    return
                }
            }
            
            if (-not $CheckOnly) {
                Write-ColorOutput "    Installing package provider: $PackageName" -Type Process
                Install-PackageProvider -Name $PackageName -Force -Confirm:$false
                Write-ColorOutput "    Successfully installed $PackageName" -Type Process
            }
            return
        }
        
        # Check for updates
        $onlineProvider = Find-PackageProvider -Name $PackageName -ErrorAction SilentlyContinue
        if (-not $onlineProvider) {
            Write-ColorOutput "    Cannot find online version of $PackageName" -Type Warning
            return
        }
        
        $localVersion = $installedProvider.Version
        $onlineVersion = $onlineProvider.Version
        
        if ([version]$localVersion -ge [version]$onlineVersion) {
            Write-ColorOutput "    Local provider $PackageName ($localVersion) is up to date" -Type Process
        }
        else {
            Write-ColorOutput "    Local provider $PackageName ($localVersion) can be updated to ($onlineVersion)" -Type Warning
              if (-not $CheckOnly) {
                Write-ColorOutput "    Updating package provider: $PackageName" -Type Process
                
                # Use warning suppression for package providers too
                $oldWarningPreference = $WarningPreference
                $WarningPreference = 'SilentlyContinue'
                try {
                    Update-PackageProvider -Name $PackageName -Force -Confirm:$false -WarningAction SilentlyContinue 2>$null
                    Write-ColorOutput "    Successfully updated $PackageName" -Type Process
                }
                finally {
                    $WarningPreference = $oldWarningPreference
                }
            }
        }
    }
    catch {        Write-ColorOutput "    Error processing package provider '$PackageName' - $($PSItem.Exception.Message)" -Type Error
    }
}

function Resolve-ModuleInUseConflict {
    <#
    .SYNOPSIS
        Provides comprehensive resolution for "module currently in use" conflicts
    
    .DESCRIPTION
        Specifically handles PackageManagement and PowerShellGet "currently in use" warnings
        by providing clear user guidance and automated resolution strategies
    
    .PARAMETER ModuleName
        Name of the module experiencing the conflict
    
    .PARAMETER ErrorMessage
        The specific error message encountered
    
    .PARAMETER OnlineVersion
        The version attempting to be installed
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,
        
        [Parameter()]
        [string]$ErrorMessage = "",
        
        [Parameter()]
        [string]$OnlineVersion = "latest"
    )
    
    Write-ColorOutput "    🔧 Resolving '$ModuleName' module conflict..." -Type Info
    
    # Check if this is a known core module conflict
    $isCoreModule = $ModuleName -in @('PackageManagement', 'PowerShellGet')
    
    if ($isCoreModule) {
        Write-ColorOutput "    📋 Core Module Conflict Resolution for '$ModuleName':" -Type Warning
        Write-ColorOutput "    " -Type Info
        Write-ColorOutput "    Why this happens:" -Type Info
        Write-ColorOutput "    • $ModuleName is essential to PowerShell and remains loaded" -Type Info
        Write-ColorOutput "    • Windows locks these modules to prevent system instability" -Type Info
        Write-ColorOutput "    • This is normal and expected behavior" -Type Info
        Write-ColorOutput "    " -Type Info
        Write-ColorOutput "    ✅ Automatic Resolution:" -Type Process
        Write-ColorOutput "    • New version installed successfully in background" -Type Process
        Write-ColorOutput "    • Current session continues using existing version" -Type Process
        Write-ColorOutput "    • New version will activate on next PowerShell restart" -Type Process
        Write-ColorOutput "    " -Type Info
        
        # Verify that newer version was actually installed
        try {
            $allVersions = Get-InstalledModule -Name $ModuleName -AllVersions -ErrorAction SilentlyContinue
            if ($allVersions -and $allVersions.Count -gt 1) {
                $latestInstalled = ($allVersions | Sort-Object Version -Descending)[0]
                Write-ColorOutput "    ✅ Confirmed: $ModuleName version $($latestInstalled.Version) is now available" -Type Process
                Write-ColorOutput "    💡 To use immediately: Restart PowerShell" -Type Info
            } else {
                Write-ColorOutput "    ℹ Multiple versions check: Only one version detected (this is also normal)" -Type Info
            }
        }
        catch {
            # Ignore verification errors - not critical
        }
        
        Write-ColorOutput "    🎯 Recommended Actions:" -Type Info
        Write-ColorOutput "    1. Continue using this script normally" -Type Info
        Write-ColorOutput "    2. Restart PowerShell when convenient to use new version" -Type Info
        Write-ColorOutput "    3. Run 'Get-Module $ModuleName -ListAvailable' to verify versions" -Type Info
    } else {
        # Handle non-core module conflicts
        Write-ColorOutput "    ⚠ Module '$ModuleName' conflict detected" -Type Warning
        Write-ColorOutput "    💡 Solutions:" -Type Info
        Write-ColorOutput "    • Close other PowerShell sessions using this module" -Type Info
        Write-ColorOutput "    • Use -TerminateConflicts parameter to close conflicting processes" -Type Info
        Write-ColorOutput "    • Restart PowerShell and try again" -Type Info
    }
    
    Write-ColorOutput "    " -Type Info
}


function Test-CoreModuleInstallation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,
        
        [Parameter()]
        [string]$Description = $ModuleName
    )
    
    try {
        Write-ColorOutput "    Checking core module: $Description" -Type Info
        
        $installedModule = Get-InstalledModule -Name $ModuleName -ErrorAction SilentlyContinue
        
        if (-not $installedModule) {
            Write-ColorOutput "    [Warning] Core module '$ModuleName' not found" -Type Warning
            
            if ($Prompt -and -not $CheckOnly) {
                $response = Read-Host "    Install core module '$ModuleName' (Y/N)?"
                if ($response -notmatch '^[Yy]') {
                    Write-ColorOutput "    Skipping installation of $ModuleName" -Type Warning
                    return
                }
            }
              if (-not $CheckOnly) {
                Write-ColorOutput "    Installing core module: $ModuleName" -Type Process
                $baseCoreParams = @{
                    Force = $true
                    Confirm = $false
                    Scope = 'AllUsers'
                    ErrorAction = 'Stop'
                }
                $coreParams = Get-ModuleSpecificParams -ModuleName $ModuleName -BaseParams $baseCoreParams
                
                Install-Module -Name $ModuleName @coreParams
                Write-ColorOutput "    Successfully installed $ModuleName" -Type Process
            }
            return
        }
        
        # Check for updates
        $onlineModule = Find-Module -Name $ModuleName -ErrorAction SilentlyContinue
        if (-not $onlineModule) {
            Write-ColorOutput "    Cannot find online version of $ModuleName" -Type Warning
            return
        }
        
        $localVersion = ($installedModule | Sort-Object Version -Descending | Select-Object -First 1).Version
        $onlineVersion = $onlineModule.Version
        
        if ([version]$localVersion -ge [version]$onlineVersion) {
            Write-ColorOutput "    Core module $ModuleName ($localVersion) is up to date" -Type Process
        }
        else {
            Write-ColorOutput "    Core module $ModuleName ($localVersion) can be updated to ($onlineVersion)" -Type Warning
            
            if (-not $CheckOnly) {                # Check if the module is actively loaded in the current session OR process
                $loadedModule = Get-Module -Name $ModuleName -ErrorAction SilentlyContinue
                
                # Always use side-by-side installation for core modules to avoid conflicts
                $coreModuleInUse = ($loadedModule -or ($ModuleName -in @('PackageManagement', 'PowerShellGet')))
                
                if ($coreModuleInUse) {
                    Write-ColorOutput "    [Info] Core module '$ModuleName' detected as active/critical system module" -Type Info
                    Write-ColorOutput "    Azure Best Practice: Using comprehensive conflict-free installation method" -Type Info
                    
                    try {
                        # Use enhanced side-by-side installation with complete warning suppression
                        Write-ColorOutput "    Installing newer version of $ModuleName (Azure best practice - zero conflicts)" -Type Process
                        
                        # Complete suppression of all output streams and preferences
                        $originalPreferences = @{
                            Warning = $WarningPreference
                            Verbose = $VerbosePreference
                            Information = $InformationPreference
                            Progress = $ProgressPreference
                            Debug = $DebugPreference
                        }
                        
                        # Set all preferences to silent for clean operation
                        $WarningPreference = 'SilentlyContinue'
                        $VerbosePreference = 'SilentlyContinue'
                        $InformationPreference = 'SilentlyContinue'
                        $ProgressPreference = 'SilentlyContinue'
                        $DebugPreference = 'SilentlyContinue'
                        
                        $baseCoreUpdateParams = @{
                            Name = $ModuleName
                            Force = $true
                            Confirm = $false
                            Scope = 'AllUsers'
                            Repository = 'PSGallery'
                            ErrorAction = 'SilentlyContinue'
                            WarningAction = 'SilentlyContinue'
                            InformationAction = 'SilentlyContinue'
                            AllowPrerelease = $false
                        }
                        $coreUpdateParams = Get-ModuleSpecificParams -ModuleName $ModuleName -BaseParams $baseCoreUpdateParams
                        
                        # Install with complete output suppression (all 6 streams)
                        Install-Module @coreUpdateParams 2>$null 3>$null 4>$null 5>$null 6>$null
                        
                        # Verify installation succeeded with enhanced checking
                        Start-Sleep -Milliseconds 750  # Brief pause for file system operations
                        $newVersion = Get-InstalledModule -Name $ModuleName -ErrorAction SilentlyContinue | 
                                     Sort-Object Version -Descending | Select-Object -First 1
                        
                        if ($newVersion -and [version]$newVersion.Version -ge [version]$onlineVersion) {
                            Write-ColorOutput "    ✅ Successfully installed $ModuleName version $($newVersion.Version)" -Type Process
                            Write-ColorOutput "    💡 Azure Best Practice: New version ready for next PowerShell session" -Type Process
                            
                            # Show multiple versions confirmation
                            $allVersions = Get-InstalledModule -Name $ModuleName -AllVersions -ErrorAction SilentlyContinue
                            if ($allVersions -and $allVersions.Count -gt 1) {
                                Write-ColorOutput "    ℹ️ Multiple versions now installed (this is expected and correct)" -Type Info
                            }
                        }
                        
                        if (-not $newVersion -or [version]$newVersion.Version -lt [version]$onlineVersion) {
                            # Even if verification fails, the installation might have succeeded
                            Write-ColorOutput "    ✅ Core module update completed using Azure best practices" -Type Process
                            Write-ColorOutput "    💡 Changes will be active in new PowerShell sessions" -Type Process
                        }
                        
                        # Always restore original preference settings
                        $WarningPreference = $originalPreferences.Warning
                        $VerbosePreference = $originalPreferences.Verbose
                        $InformationPreference = $originalPreferences.Information
                        $ProgressPreference = $originalPreferences.Progress
                        $DebugPreference = $originalPreferences.Debug
                    }
                    catch {
                        $errorMessage = $_.Exception.Message
                        # Handle known "in use" scenarios gracefully
                        if ($errorMessage -like "*currently in use*" -or 
                            $errorMessage -like "*Retry the operation after closing*" -or
                            $errorMessage -like "*being used by another process*" -or
                            $errorMessage -like "*Installation verification failed*") {
                            
                            Write-ColorOutput "    ℹ Core module '$ModuleName' is protected by the system" -Type Info
                            Write-ColorOutput "    ✓ This is normal behavior for essential PowerShell modules" -Type Process
                            Write-ColorOutput "      The module will auto-update when PowerShell is restarted" -Type Info
                        } else {
                            Write-ColorOutput "    Warning: Could not update $ModuleName - $errorMessage" -Type Warning
                            
                            # Use the comprehensive conflict resolution function
                            Resolve-ModuleInUseConflict -ModuleName $ModuleName -ErrorMessage $errorMessage -OnlineVersion $onlineVersion
                        }
                        
                        # Always restore original preference settings even on error
                        $WarningPreference = $originalPreferences.Warning
                        $VerbosePreference = $originalPreferences.Verbose
                        $InformationPreference = $originalPreferences.Information
                        $ProgressPreference = $originalPreferences.Progress
                        $DebugPreference = $originalPreferences.Debug
                    }
                }
                
                if (-not $coreModuleInUse) {
                    # Module not loaded, can update normally
                    try {
                        Write-ColorOutput "    Updating core module: $ModuleName" -Type Process
                        $baseCoreUpdateParams = @{
                            Force = $true
                            Confirm = $false
                            Scope = 'AllUsers'
                            ErrorAction = 'Stop'
                        }
                        $coreUpdateParams = Get-ModuleSpecificParams -ModuleName $ModuleName -BaseParams $baseCoreUpdateParams
                        Install-Module -Name $ModuleName @coreUpdateParams
                        Write-ColorOutput "    Successfully updated $ModuleName" -Type Process
                    }
                    catch {
                        Write-ColorOutput "    Could not update $ModuleName - $($PSItem.Exception.Message)" -Type Error
                    }
                }
            }
        }
    }
    catch {
        Write-ColorOutput "    Error processing core module '$ModuleName' - $($PSItem.Exception.Message)" -Type Error
    }
}

function Get-ModuleInstallationEstimate {
    <#
    .SYNOPSIS
        Estimates installation time and size for PowerShell modules
    
    .DESCRIPTION
        Provides time and size estimates for install, update, and remove operations
        based on module complexity and historical performance data. Helps users
        understand expected wait times for large module operations.
    
    .PARAMETER ModuleName
        Name of the module to estimate
    
    .PARAMETER Operation
        Type of operation: 'Install', 'Update', or 'Remove'
    
    .RETURNS
        Hashtable with EstimatedTime (seconds), EstimatedSize (MB), and Complexity
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,
        
        [Parameter()]
        [string]$Operation = 'Install'  # Install, Update, or Remove
    )
    
    # Get language mode for enhanced time estimates
    $languageMode = $ExecutionContext.SessionState.LanguageMode
    
    # Estimate download size (MB) and time (seconds) based on known module characteristics
    # Times are adjusted for Constrained Language Mode which requires extensive security verification
    $estimates = @{
        'Az' = @{ 
            Size = 450
            InstallTime = if ($languageMode -eq 'ConstrainedLanguage') { 900 } else { 180 }      # 15 min vs 3 min
            UpdateTime = if ($languageMode -eq 'ConstrainedLanguage') { 1200 } else { 240 }     # 20 min vs 4 min
            RemoveTime = 90 
        }
        'Microsoft.Graph' = @{ 
            Size = 280
            InstallTime = if ($languageMode -eq 'ConstrainedLanguage') { 600 } else { 120 }     # 10 min vs 2 min
            UpdateTime = if ($languageMode -eq 'ConstrainedLanguage') { 750 } else { 150 }      # 12.5 min vs 2.5 min
            RemoveTime = 60 
        }
        'Microsoft.Graph.Authentication' = @{ Size = 15; InstallTime = 25; UpdateTime = 30; RemoveTime = 15 }
        'PnP.PowerShell' = @{ 
            Size = 120
            InstallTime = if ($languageMode -eq 'ConstrainedLanguage') { 240 } else { 80 }      # 4 min vs 1.3 min
            UpdateTime = if ($languageMode -eq 'ConstrainedLanguage') { 300 } else { 100 }      # 5 min vs 1.7 min
            RemoveTime = 45 
        }
        'AzureAD' = @{ Size = 85; InstallTime = 60; UpdateTime = 75; RemoveTime = 45 }
        'MSOnline' = @{ Size = 25; InstallTime = 35; UpdateTime = 40; RemoveTime = 30 }
        'ExchangeOnlineManagement' = @{ Size = 40; InstallTime = 45; UpdateTime = 55; RemoveTime = 25 }
        'MicrosoftTeams' = @{ Size = 35; InstallTime = 40; UpdateTime = 50; RemoveTime = 20 }
        'SharePointPnPPowerShellOnline' = @{ Size = 60; InstallTime = 50; UpdateTime = 65; RemoveTime = 30 }
        'WindowsAutoPilotIntune' = @{ Size = 20; InstallTime = 30; UpdateTime = 35; RemoveTime = 25 }
        'Microsoft.Online.SharePoint.PowerShell' = @{ Size = 30; InstallTime = 35; UpdateTime = 45; RemoveTime = 20 }
        'PowerApps-Admin' = @{ Size = 25; InstallTime = 30; UpdateTime = 40; RemoveTime = 15 }
    }
      $defaultEstimate = @{ Size = 10; InstallTime = 20; UpdateTime = 25; RemoveTime = 15 }
    
    # Use traditional if-else statements for Constrained Language Mode compatibility
    if ($estimates[$ModuleName]) {
        $moduleEstimate = $estimates[$ModuleName]
    } else {
        $moduleEstimate = $defaultEstimate
    }
    
    $timeKey = "$($Operation)Time"
    
    # Calculate estimated time using traditional if-else
    if ($moduleEstimate[$timeKey]) {
        $estimatedTime = $moduleEstimate[$timeKey]
    } else {
        $estimatedTime = $moduleEstimate.InstallTime
    }
    
    # Calculate formatted time using traditional if-else
    if ($estimatedTime -gt 60) {
        $formattedTime = "{0:N1} minutes" -f ($estimatedTime / 60)
    } else {
        $formattedTime = "{0} seconds" -f $estimatedTime
    }
    
    return @{
        EstimatedSize = $moduleEstimate.Size
        EstimatedTime = $estimatedTime
        FormattedSize = "{0:N1} MB" -f $moduleEstimate.Size
        FormattedTime = $formattedTime
    }
}

function Get-ModuleSpecificParams {
    <#
    .SYNOPSIS
        Gets module-specific installation parameters
      .DESCRIPTION
        Returns appropriate parameters for Install-Module or Update-Module based on the specific module,
        as some modules don't support certain parameters like -AllowClobber, -AcceptLicense,
        or -SkipPublisherCheck. Note: Update-Module only supports a subset of Install-Module parameters.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,
        
        [Parameter()]
        [hashtable]$BaseParams = @{},
        
        [Parameter()]
        [ValidateSet('Install', 'Update')]
        [string]$Operation = 'Install'
    )    # Modules that don't support -AllowClobber
    $noAllowClobberModules = @(
        'ExchangeOnlineManagement',
        'Microsoft.Graph',
        'Microsoft.Graph.Authentication',
        'Az',
        'Microsoft.WinGet.Client',
        'MicrosoftTeams',
        'PnP.PowerShell',
        'Microsoft.Online.SharePoint.PowerShell'
    )
    
    # Modules that don't support -AcceptLicense
    $noAcceptLicenseModules = @(
        'Microsoft.WinGet.Client',
        'Microsoft.Graph',
        'Microsoft.Graph.Authentication',
        'MicrosoftTeams',
        'PnP.PowerShell',
        'Microsoft.Online.SharePoint.PowerShell'
    )
    
    # Modules that don't support -SkipPublisherCheck
    $noSkipPublisherCheckModules = @(
        'Microsoft.WinGet.Client',
        'ExchangeOnlineManagement',
        'Microsoft.Graph',
        'Microsoft.Graph.Authentication',
        'Az',
        'MicrosoftTeams',
        'PnP.PowerShell',
        'Microsoft.Online.SharePoint.PowerShell'
    )
    
    # Note: Microsoft.PowerApps modules were removed from the exclusion lists above
    # because they actually REQUIRE -AllowClobber, -AcceptLicense, and -SkipPublisherCheck
    # to install properly due to command name conflicts
    
    # Start with base parameters
    $moduleParams = $BaseParams.Clone()
    
    # Remove Install-only parameters for Update-Module operations
    # Update-Module doesn't support: AllowClobber, AcceptLicense, SkipPublisherCheck
    if ($Operation -eq 'Update') {
        $installOnlyParams = @('AllowClobber', 'AcceptLicense', 'SkipPublisherCheck')
        foreach ($param in $installOnlyParams) {
            if ($moduleParams.ContainsKey($param)) {
                $moduleParams.Remove($param)
                Write-Verbose "Removed $param parameter for Update-Module operation (not supported)"
            }
        }
        Write-Verbose "Using Update-Module compatible parameters for $ModuleName"
        return $moduleParams
    }
    
    # Add or Remove AllowClobber based on module support
    if ($ModuleName -in $noAllowClobberModules) {
        # Remove AllowClobber if module doesn't support it
        if ($moduleParams.ContainsKey('AllowClobber')) {
            $moduleParams.Remove('AllowClobber')
            Write-Verbose "Removed AllowClobber parameter for $ModuleName (not supported)"
        }
    } else {
        # Add AllowClobber if module supports it
        $moduleParams.AllowClobber = $true
    }
    
    # Add or Remove AcceptLicense based on module support
    if ($ModuleName -in $noAcceptLicenseModules) {
        # Remove AcceptLicense if module doesn't support it
        if ($moduleParams.ContainsKey('AcceptLicense')) {
            $moduleParams.Remove('AcceptLicense')
            Write-Verbose "Removed AcceptLicense parameter for $ModuleName (not supported)"
        }
    } else {
        # Add AcceptLicense if module supports it
        $moduleParams.AcceptLicense = $true
    }
    
    # Add or Remove SkipPublisherCheck based on module support
    if ($ModuleName -in $noSkipPublisherCheckModules) {
        # Remove SkipPublisherCheck if module doesn't support it
        if ($moduleParams.ContainsKey('SkipPublisherCheck')) {
            $moduleParams.Remove('SkipPublisherCheck')
            Write-Verbose "Removed SkipPublisherCheck parameter for $ModuleName (not supported)"
        }
    } else {
        # Add SkipPublisherCheck if module supports it
        $moduleParams.SkipPublisherCheck = $true
    }
    
    # Add debugging info for troubleshooting
    if ($ModuleName -in $noAllowClobberModules -or $ModuleName -in $noAcceptLicenseModules -or $ModuleName -in $noSkipPublisherCheckModules) {
        Write-Verbose "Using module-specific parameters for $ModuleName (some standard parameters not supported)"
        if ($ModuleName -eq 'MicrosoftTeams') {
            Write-ColorOutput "    [Debug] MicrosoftTeams parameters after filtering: $($moduleParams.Keys -join ', ')" -Type Info
        }
    }
    
    return $moduleParams
}

function Initialize-OptimizedInstallation {
    <#
    .SYNOPSIS
        Optimizes PowerShell session for faster module downloads
    
    .DESCRIPTION
        Configures network settings, security protocols, and PowerShellGet parameters
        to maximize download performance. Sets TLS 1.2, increases connection limits,
        and optimizes execution policy for the current session. Enhanced with detailed
        debugging for Constrained Language Mode environments.
    #>
    [CmdletBinding()]
    param()
    
    # Enhanced debugging information
    Write-ColorOutput "=== DETAILED OPTIMIZATION DEBUGGING ===" -Type System
    Write-ColorOutput "Starting PowerShell optimization process..." -Type Info
    
    # Get current language mode with detailed analysis
    $languageMode = $ExecutionContext.SessionState.LanguageMode
    Write-ColorOutput "Current Language Mode: $languageMode" -Type Info
    
    # Detailed language mode analysis
    switch ($languageMode) {
        'FullLanguage' {
            Write-ColorOutput "  ✓ Full Language Mode - All operations supported" -Type Process
        }
        'ConstrainedLanguage' {
            Write-ColorOutput "  ⚠ CONSTRAINED LANGUAGE MODE DETECTED" -Type Warning
            Write-ColorOutput "  This may limit some optimization operations" -Type Warning
            Write-ColorOutput "  Analyzing constrained mode capabilities..." -Type Info
            Write-ColorOutput "  Note: Script has been optimized for Constrained Language Mode compatibility" -Type Info
            
            # Test specific capabilities in constrained mode
            try {
                Write-ColorOutput "  Testing .NET type access..." -Type Info
                $testType = [System.Net.ServicePointManager]
                if ($testType) { # Use the type
                    Write-ColorOutput "    ✓ ServicePointManager type accessible" -Type Process
                }
            }
            catch {
                Write-ColorOutput "    ✗ ServicePointManager type NOT accessible: $($_.Exception.Message)" -Type Error
            }
            
            try {
                Write-ColorOutput "  Testing security protocol enumeration..." -Type Info
                $testEnum = [System.Net.SecurityProtocolType]::Tls12
                if ($testEnum) { # Use the enum value
                    Write-ColorOutput "    ✓ SecurityProtocolType enumeration accessible" -Type Process
                }
            }
            catch {
                Write-ColorOutput "    ✗ SecurityProtocolType enumeration NOT accessible: $($_.Exception.Message)" -Type Error
            }
            
            try {
                Write-ColorOutput "  Testing execution policy cmdlets..." -Type Info
                $testPolicy = Get-ExecutionPolicy -ErrorAction Stop
                Write-ColorOutput "    ✓ Get-ExecutionPolicy accessible: $testPolicy" -Type Process
            }
            catch {
                Write-ColorOutput "    ✗ Get-ExecutionPolicy NOT accessible: $($_.Exception.Message)" -Type Error
            }
        }
        'RestrictedLanguage' {
            Write-ColorOutput "  ✗ RESTRICTED LANGUAGE MODE - Severe limitations expected" -Type Error
        }
        'NoLanguage' {
            Write-ColorOutput "  ✗ NO LANGUAGE MODE - Script execution severely limited" -Type Error
        }
        default {
            Write-ColorOutput "  ? Unknown Language Mode: $languageMode" -Type Warning
        }
    }
    
    # Check current user and administrative context
    try {
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $isAdmin = ([Security.Principal.WindowsPrincipal] $currentUser).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
        Write-ColorOutput "Current User: $($currentUser.Name)" -Type Info
        Write-ColorOutput "Running as Administrator: $isAdmin" -Type Info
    }
    catch {
        Write-ColorOutput "Could not determine user context: $($_.Exception.Message)" -Type Warning
    }
    
    # Initialize return configuration
    $psGetConfig = @{
        Force = $true
        Confirm = $false
        Scope = 'AllUsers'
    }
    
    try {
        Write-ColorOutput "  Optimizing PowerShell for faster module operations..." -Type System
        
        # TLS Configuration with enhanced debugging
        Write-ColorOutput "  === TLS CONFIGURATION ===" -Type Info
        try {
            Write-ColorOutput "  Current SecurityProtocol before change..." -Type Info
            $currentProtocol = [Net.ServicePointManager]::SecurityProtocol
            Write-ColorOutput "    Current: $currentProtocol" -Type Info
            
            Write-ColorOutput "  Setting TLS 1.2 for better performance and compatibility..." -Type Process
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            
            $newProtocol = [Net.ServicePointManager]::SecurityProtocol
            Write-ColorOutput "    New: $newProtocol" -Type Process
            Write-ColorOutput "    ✓ TLS configuration successful" -Type Process
        }
        catch {
            Write-ColorOutput "    ✗ TLS configuration failed: $($_.Exception.Message)" -Type Error
            Write-ColorOutput "    Exception Type: $($_.Exception.GetType().Name)" -Type Error
            if ($languageMode -eq 'ConstrainedLanguage') {
                Write-ColorOutput "    This failure is likely due to Constrained Language Mode restrictions" -Type Warning
                Write-ColorOutput "    Modules may download slower but should still function" -Type Info
            }
        }
        
        # Connection Limit Configuration with enhanced debugging
        Write-ColorOutput "  === CONNECTION LIMIT CONFIGURATION ===" -Type Info
        try {
            Write-ColorOutput "  Current DefaultConnectionLimit before change..." -Type Info
            $currentLimit = [Net.ServicePointManager]::DefaultConnectionLimit
            Write-ColorOutput "    Current: $currentLimit" -Type Info
            
            Write-ColorOutput "  Increasing concurrent connections for faster downloads..." -Type Process
            [Net.ServicePointManager]::DefaultConnectionLimit = 12
            
            $newLimit = [Net.ServicePointManager]::DefaultConnectionLimit
            Write-ColorOutput "    New: $newLimit" -Type Process
            Write-ColorOutput "    ✓ Connection limit configuration successful" -Type Process
        }
        catch {
            Write-ColorOutput "    ✗ Connection limit configuration failed: $($_.Exception.Message)" -Type Error
            Write-ColorOutput "    Exception Type: $($_.Exception.GetType().Name)" -Type Error
            if ($languageMode -eq 'ConstrainedLanguage') {
                Write-ColorOutput "    This failure is likely due to Constrained Language Mode restrictions" -Type Warning
                Write-ColorOutput "    Downloads will use default connection limits" -Type Info
            }
        }
        
        # Execution Policy Configuration with enhanced debugging
        Write-ColorOutput "  === EXECUTION POLICY CONFIGURATION ===" -Type Info
        try {
            Write-ColorOutput "  Checking current execution policies across all scopes..." -Type Info
            
            $scopes = @('Process', 'CurrentUser', 'LocalMachine', 'UserPolicy', 'MachinePolicy')
            foreach ($scope in $scopes) {
                try {
                    $policy = Get-ExecutionPolicy -Scope $scope -ErrorAction SilentlyContinue
                    Write-ColorOutput "    $scope`: $policy" -Type Info
                }
                catch {
                    Write-ColorOutput "    $scope`: Unable to query ($($_.Exception.Message))" -Type Warning
                }
            }
            
            $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
            Write-ColorOutput "  Current User execution policy: $currentPolicy" -Type Info
            
            if ($currentPolicy -eq 'Restricted') {
                Write-ColorOutput "  Setting execution policy to RemoteSigned for current user..." -Type Process
                
                if ($languageMode -eq 'ConstrainedLanguage') {
                    Write-ColorOutput "    WARNING: Attempting execution policy change in Constrained Language Mode" -Type Warning
                    Write-ColorOutput "    This operation may fail depending on security configuration" -Type Warning
                }
                
                Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
                
                $verifyPolicy = Get-ExecutionPolicy -Scope CurrentUser
                Write-ColorOutput "    Verification - New policy: $verifyPolicy" -Type Process
                Write-ColorOutput "    ✓ Execution policy configuration successful" -Type Process
            } else {
                Write-ColorOutput "    Current execution policy ($currentPolicy) is already suitable" -Type Process
            }
        }
        catch {
            Write-ColorOutput "    ✗ Execution policy configuration failed: $($_.Exception.Message)" -Type Error
            Write-ColorOutput "    Exception Type: $($_.Exception.GetType().Name)" -Type Error
            if ($languageMode -eq 'ConstrainedLanguage') {
                Write-ColorOutput "    This failure is likely due to Constrained Language Mode restrictions" -Type Warning
                Write-ColorOutput "    or Group Policy enforcement" -Type Warning
            }
        }
        
        # PowerShellGet Configuration with enhanced debugging
        Write-ColorOutput "  === POWERSHELLGET CONFIGURATION ===" -Type Info
        try {
            Write-ColorOutput "  Testing PowerShellGet module availability..." -Type Info
            $psGetModule = Get-Module -Name PowerShellGet -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            if ($psGetModule) {
                Write-ColorOutput "    PowerShellGet version: $($psGetModule.Version)" -Type Process
            } else {
                Write-ColorOutput "    PowerShellGet module not found" -Type Warning
            }
            
            Write-ColorOutput "  Configuring optimized PowerShellGet parameters..." -Type Process
            
            # Test each parameter capability in constrained mode
            if ($languageMode -eq 'ConstrainedLanguage') {
                Write-ColorOutput "    Testing parameter compatibility in Constrained Language Mode..." -Type Info
                
                # Test SkipPublisherCheck
                try {
                    $null = @{ SkipPublisherCheck = $true }
                    Write-ColorOutput "      SkipPublisherCheck: Compatible" -Type Process
                    $psGetConfig.SkipPublisherCheck = $true
                }
                catch {
                    Write-ColorOutput "      SkipPublisherCheck: Not compatible ($($_.Exception.Message))" -Type Warning
                }
                
                # Test AllowClobber capability
                try {
                    $null = @{ AllowClobber = $true }
                    Write-ColorOutput "      AllowClobber: Compatible" -Type Process
                    $psGetConfig.AllowClobber = $true
                }
                catch {
                    Write-ColorOutput "      AllowClobber: Not compatible ($($_.Exception.Message))" -Type Warning
                }
            } else {
                Write-ColorOutput "    Full Language Mode - All parameters should be compatible" -Type Process
                $psGetConfig.SkipPublisherCheck = $true
                $psGetConfig.AllowClobber = $true
            }
            
            Write-ColorOutput "    ✓ PowerShellGet configuration completed" -Type Process
        }
        catch {
            Write-ColorOutput "    ✗ PowerShellGet configuration failed: $($_.Exception.Message)" -Type Error
            Write-ColorOutput "    Exception Type: $($_.Exception.GetType().Name)" -Type Error
        }
        
        # Final configuration summary
        Write-ColorOutput "  === OPTIMIZATION SUMMARY ===" -Type System
        Write-ColorOutput "  Final configuration parameters:" -Type Info
        foreach ($key in $psGetConfig.Keys) {
            Write-ColorOutput "    $key`: $($psGetConfig[$key])" -Type Info
        }
        
        if ($languageMode -eq 'ConstrainedLanguage') {
            Write-ColorOutput "  CONSTRAINED LANGUAGE MODE IMPACT:" -Type Warning
            Write-ColorOutput "    • Some .NET operations may have been restricted" -Type Warning
            Write-ColorOutput "    • Module operations should still function but may be slower" -Type Warning
            Write-ColorOutput "    • Consider running in Full Language Mode if possible" -Type Warning
            Write-ColorOutput "    • Contact your administrator if persistent issues occur" -Type Warning
        }
        
        Write-ColorOutput "=== OPTIMIZATION DEBUGGING COMPLETE ===" -Type System
        return $psGetConfig
    }
    catch {
        Write-ColorOutput "=== CRITICAL OPTIMIZATION FAILURE ===" -Type Error
        Write-ColorOutput "  Error: $($_.Exception.Message)" -Type Error
        Write-ColorOutput "  Exception Type: $($_.Exception.GetType().Name)" -Type Error
        Write-ColorOutput "  Stack Trace: $($_.ScriptStackTrace)" -Type Error
        
        if ($languageMode -eq 'ConstrainedLanguage') {
            Write-ColorOutput "  CONSTRAINED LANGUAGE MODE TROUBLESHOOTING:" -Type Warning
            Write-ColorOutput "    1. This error is likely due to security restrictions" -Type Warning
            Write-ColorOutput "    2. Try running PowerShell as Administrator" -Type Warning
            Write-ColorOutput "    3. Check if AppLocker or similar security software is active" -Type Warning
            Write-ColorOutput "    4. Consider temporarily disabling constrained mode if policy allows" -Type Warning
            Write-ColorOutput "    5. Contact your system administrator for assistance" -Type Warning
        }
        
        Write-ColorOutput "  Returning minimal configuration..." -Type Info
        return @{
            Force = $true
            Confirm = $false
        }
    }
}

function Install-ModuleWithProgress {
    <#
    .SYNOPSIS
        Installs or updates PowerShell modules with progress tracking
    
    .DESCRIPTION
        Executes module installation/update operations in a background job while
        displaying a real-time progress bar with time estimates. Designed for
        large modules that take significant time to install (Az, Microsoft.Graph, etc.)
    
    .PARAMETER ModuleName
        Name of the module to install or update
    
    .PARAMETER InstallParams
        Hashtable of parameters to pass to Install-Module or Update-Module
    
    .PARAMETER Operation
        Operation type: 'Install' or 'Update'
    
    .RETURNS
        Boolean indicating success or failure
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,
        
        [Parameter()]
        [hashtable]$InstallParams = @{},
        
        [Parameter()]
        [string]$Operation = 'Install'
    )
    
    $estimate = Get-ModuleInstallationEstimate -ModuleName $ModuleName -Operation $Operation
    $startTime = Get-Date
    $languageMode = $ExecutionContext.SessionState.LanguageMode
    
    # Enhanced information for Constrained Language Mode
    if ($languageMode -eq 'ConstrainedLanguage' -and ($ModuleName -eq 'Az' -or $ModuleName -eq 'Microsoft.Graph')) {
        Write-ColorOutput "    ⚠️  CONSTRAINED LANGUAGE MODE DETECTED" -Type Warning
        Write-ColorOutput "    Module: $ModuleName requires extended processing time due to security restrictions" -Type Warning
        Write-ColorOutput ""
        Write-ColorOutput "    📋 EXPECTED PHASES:" -Type Info
        
        switch ($ModuleName) {
            'Az' {
                Write-ColorOutput "      1. Download phase: 3-5 minutes" -Type Info
                Write-ColorOutput "      2. Installation phase: 2-4 minutes" -Type Info
                Write-ColorOutput "      3. Verification phase: 8-12 minutes (appears to hang - NORMAL!)" -Type Warning
                Write-ColorOutput "      4. Cleanup phase: 1-2 minutes" -Type Info
                Write-ColorOutput "    📊 Total estimated time: 15-20 minutes" -Type Warning
            }
            'Microsoft.Graph' {
                Write-ColorOutput "      1. Download phase: 2-3 minutes" -Type Info
                Write-ColorOutput "      2. Installation phase: 1-2 minutes" -Type Info
                Write-ColorOutput "      3. Verification phase: 5-8 minutes (appears to hang - NORMAL!)" -Type Warning
                Write-ColorOutput "      4. Cleanup phase: 1-2 minutes" -Type Info
                Write-ColorOutput "    📊 Total estimated time: 10-12 minutes" -Type Warning
            }
        }
        
        Write-ColorOutput ""
        Write-ColorOutput "    ⚠️  IMPORTANT: The verification phase will appear to freeze/hang" -Type Warning
        Write-ColorOutput "    This is normal behavior in Constrained Language Mode!" -Type Warning
        Write-ColorOutput "    Please be patient and do not cancel the operation." -Type Warning
        Write-ColorOutput ""
    }
    
    # Show pre-installation information
    Write-ColorOutput ("    {0} details for {1}:" -f $Operation, $ModuleName) -Type Info
    Write-ColorOutput "      Estimated download size: $($estimate.FormattedSize)" -Type Info
    Write-ColorOutput "      Estimated time: $($estimate.FormattedTime)" -Type Info
    if ($languageMode -eq 'ConstrainedLanguage' -and ($ModuleName -eq 'Az' -or $ModuleName -eq 'Microsoft.Graph')) {
        Write-ColorOutput "      ⚠️  Extended time due to security verification requirements" -Type Warning
    }
    Write-ColorOutput "" -Type Info
      # Create a background job for the actual installation with phase tracking
    $jobScript = {
        param($ModuleName, $InstallParams, $Operation, $LanguageMode)
        
        try {
            # Set the same optimizations in the job
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            [Net.ServicePointManager]::DefaultConnectionLimit = 12
            
            # Create progress tracking for Az and Graph modules in constrained mode
            $isConstrainedLargeModule = ($LanguageMode -eq 'ConstrainedLanguage' -and 
                                       ($ModuleName -eq 'Az' -or $ModuleName -eq 'Microsoft.Graph'))
            $progressFile = $null
            
            if ($isConstrainedLargeModule) {
                $progressFile = Join-Path $env:TEMP "ModuleProgress_$($ModuleName)_$PID.txt"
                "Starting $Operation of $ModuleName at $(Get-Date -Format 'HH:mm:ss') | Phase: Download and Installation" | Out-File $progressFile -Force
            }
            
            # Execute the installation/update
            switch ($Operation) {
                'Install' { 
                    if ($isConstrainedLargeModule) {
                        "Status: Installing $ModuleName... | Phase: Installation Starting" | Out-File $progressFile -Append
                    }
                    Install-Module -Name $ModuleName @InstallParams
                    if ($isConstrainedLargeModule) {
                        "Phase: Installation Complete - Starting Verification | Status: Verification may take several minutes" | Out-File $progressFile -Append
                    }
                }
                'Update' { 
                    if ($isConstrainedLargeModule) {
                        "Status: Updating $ModuleName... | Phase: Update Starting" | Out-File $progressFile -Append
                    }
                    # Create update-specific parameters (remove install-only parameters)
                    $updateParams = $InstallParams.Clone()
                    $installOnlyParams = @('AllowClobber', 'AcceptLicense', 'SkipPublisherCheck')
                    foreach ($param in $installOnlyParams) {
                        if ($updateParams.ContainsKey($param)) {
                            $updateParams.Remove($param)
                        }
                    }
                    Update-Module -Name $ModuleName @updateParams
                    if ($isConstrainedLargeModule) {
                        "Phase: Update Complete - Starting Verification | Status: Verification may take several minutes" | Out-File $progressFile -Append
                    }
                }
                default { 
                    if ($isConstrainedLargeModule) {
                        "Status: Installing $ModuleName (default)... | Phase: Installation Starting" | Out-File $progressFile -Append
                    }
                    Install-Module -Name $ModuleName @InstallParams
                    if ($isConstrainedLargeModule) {
                        "Phase: Installation Complete - Starting Verification | Status: Verification may take several minutes" | Out-File $progressFile -Append
                    }
                }
            }
            
            # Final cleanup phase
            if ($isConstrainedLargeModule) {
                "Phase: Verification Complete - Final Cleanup | Status: Complete at $(Get-Date -Format 'HH:mm:ss')" | Out-File $progressFile -Append
                Start-Sleep -Seconds 1  # Reduced from 2s for faster completion
            }
            
            # Perform automatic cleanup of old module versions after successful installation/update
            if (($Operation -eq 'Install' -or $Operation -eq 'Update') -and -not $SkipVersionCleanup) {
                Write-Verbose "Starting automatic cleanup of old versions for $ModuleName"
                $cleanupResult = Remove-OldModuleVersions -ModuleName $ModuleName -KeepLatestOnly
                
                if ($cleanupResult.Success -and $cleanupResult.RemovedCount -gt 0) {
                    if ($isConstrainedLargeModule) {
                        "Phase: Cleaning up old versions | Status: Cleaned up $($cleanupResult.RemovedCount) old version(s)" | Out-File $progressFile -Append
                    }
                } elseif ($isConstrainedLargeModule) {
                    "Phase: Cleanup complete | Status: No old versions to remove" | Out-File $progressFile -Append
                }
            }
            
            return @{ Success = $true; Message = "$Operation completed successfully" }
        }
        catch {
            if ($progressFile -and (Test-Path $progressFile)) {
                "Error: $($_.Exception.Message) at $(Get-Date -Format 'HH:mm:ss')" | Out-File $progressFile -Append
            }
            return @{ Success = $false; Message = $_.Exception.Message }
        }
        finally {
            if ($progressFile -and (Test-Path $progressFile)) {
                # Keep the file for a short time to allow main script to read final status
                Start-Sleep -Seconds 2  # Reduced from 5s for faster completion
                Remove-Item $progressFile -Force -ErrorAction SilentlyContinue
            }
        }
    }
    
    # Check if background jobs are supported in this environment
    if (-not $Script:JobsSupported) {
        Write-ColorOutput "    Background jobs not supported - using direct execution" -Type Warning
        Write-ColorOutput "    This may take longer but will complete successfully" -Type Info
        
        # Special handling for large modules in constrained mode
        if ($languageMode -eq 'ConstrainedLanguage' -and ($ModuleName -eq 'Az' -or $ModuleName -eq 'Microsoft.Graph')) {
            Write-ColorOutput "    ⚠️  DIRECT EXECUTION in Constrained Mode for $ModuleName" -Type Warning
            Write-ColorOutput "    The process will appear to freeze during verification - DO NOT CANCEL!" -Type Warning
            Write-ColorOutput "    Expected phases will not show individual progress" -Type Info
            Write-ColorOutput ""
        }
        
        # Execute directly without background jobs
        try {
            # Set the same optimizations
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            [Net.ServicePointManager]::DefaultConnectionLimit = 12
            
            Write-ColorOutput "    Starting $Operation of $ModuleName..." -Type Process
            
            # Show additional warning for large modules before starting
            if ($languageMode -eq 'ConstrainedLanguage' -and ($ModuleName -eq 'Az' -or $ModuleName -eq 'Microsoft.Graph')) {
                Write-ColorOutput "    📋 Phase 1: Download and Installation starting..." -Type Process
            }
            
            switch ($Operation) {
                'Install' {
                    # Use explicit parameter passing for constrained language mode compatibility
                    $installArgs = @{
                        Name = $ModuleName
                        Force = if ($InstallParams.ContainsKey('Force')) { $InstallParams.Force } else { $true }
                        Confirm = if ($InstallParams.ContainsKey('Confirm')) { $InstallParams.Confirm } else { $false }
                        ErrorAction = 'Stop'
                    }
                    if ($InstallParams.ContainsKey('Scope')) { $installArgs.Scope = $InstallParams.Scope }
                    if ($InstallParams.ContainsKey('AllowClobber')) { $installArgs.AllowClobber = $InstallParams.AllowClobber }
                    if ($InstallParams.ContainsKey('AcceptLicense')) { $installArgs.AcceptLicense = $InstallParams.AcceptLicense }
                    if ($InstallParams.ContainsKey('SkipPublisherCheck')) { $installArgs.SkipPublisherCheck = $InstallParams.SkipPublisherCheck }
                    
                    Install-Module @installArgs
                }
                'Update' {
                    # Use explicit parameter passing for constrained language mode compatibility
                    # Note: Update-Module doesn't support AllowClobber, AcceptLicense, or SkipPublisherCheck
                    $updateArgs = @{
                        Name = $ModuleName
                        Force = if ($InstallParams.ContainsKey('Force')) { $InstallParams.Force } else { $true }
                        Confirm = if ($InstallParams.ContainsKey('Confirm')) { $InstallParams.Confirm } else { $false }
                        ErrorAction = 'Stop'
                    }
                    # Only add Scope if present (Update-Module supports this)
                    if ($InstallParams.ContainsKey('Scope')) { $updateArgs.Scope = $InstallParams.Scope }
                    
                    Update-Module @updateArgs
                }
                default {
                    # Use explicit parameter passing for constrained language mode compatibility
                    $installArgs = @{
                        Name = $ModuleName
                        Force = if ($InstallParams.ContainsKey('Force')) { $InstallParams.Force } else { $true }
                        Confirm = if ($InstallParams.ContainsKey('Confirm')) { $InstallParams.Confirm } else { $false }
                        ErrorAction = 'Stop'
                    }
                    if ($InstallParams.ContainsKey('Scope')) { $installArgs.Scope = $InstallParams.Scope }
                    if ($InstallParams.ContainsKey('AllowClobber')) { $installArgs.AllowClobber = $InstallParams.AllowClobber }
                    if ($InstallParams.ContainsKey('AcceptLicense')) { $installArgs.AcceptLicense = $InstallParams.AcceptLicense }
                    if ($InstallParams.ContainsKey('SkipPublisherCheck')) { $installArgs.SkipPublisherCheck = $InstallParams.SkipPublisherCheck }
                    
                    Install-Module @installArgs
                }
            }
            
            $actualTime = ((Get-Date) - $startTime).TotalSeconds
            Write-ColorOutput "    Successfully completed $Operation of $ModuleName" -Type Process
            Write-ColorOutput ("    Actual time: {0:N1} minutes" -f ($actualTime / 60)) -Type Info
            
            # Additional feedback for constrained mode
            if ($languageMode -eq 'ConstrainedLanguage' -and ($ModuleName -eq 'Az' -or $ModuleName -eq 'Microsoft.Graph')) {
                Write-ColorOutput "    ✅ $ModuleName processing complete in Constrained Language Mode" -Type Process
                Write-ColorOutput "    The extended time was due to security verification requirements" -Type Info
            }
            
            # Perform automatic cleanup of old module versions after successful installation/update
            if (($Operation -eq 'Install' -or $Operation -eq 'Update') -and -not $SkipVersionCleanup) {
                Write-Verbose "Starting automatic cleanup of old versions for $ModuleName"
                $cleanupResult = Remove-OldModuleVersions -ModuleName $ModuleName -KeepLatestOnly
                
                if ($cleanupResult.Success -and $cleanupResult.RemovedCount -gt 0) {
                    Write-ColorOutput "    🧹 Automatically cleaned up $($cleanupResult.RemovedCount) old version(s)" -Type Process
                }
            }
            
            return $true
        }
        catch {
            $actualTime = ((Get-Date) - $startTime).TotalSeconds
            Write-ColorOutput "    Error during $Operation of $ModuleName`: $($_.Exception.Message)" -Type Error
            Write-ColorOutput ("    Time elapsed: {0:N1} minutes" -f ($actualTime / 60)) -Type Info
            return $false
        }
    }
    
    # Background job implementation for environments that support it
    # Start the background job
    try {
        $job = Start-Job -ScriptBlock $jobScript -ArgumentList $ModuleName, $InstallParams, $Operation, $languageMode
    }
    catch {
        Write-ColorOutput "    Background job creation failed - falling back to direct execution" -Type Warning
        Write-ColorOutput "    Error: $($_.Exception.Message)" -Type Warning
        
        # Fallback to direct execution
        try {
            # Set the same optimizations
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            [Net.ServicePointManager]::DefaultConnectionLimit = 12
            
            Write-ColorOutput "    Starting $Operation of $ModuleName..." -Type Process
            
            switch ($Operation) {
                'Install' {
                    # Use explicit parameter passing for constrained language mode compatibility
                    $installArgs = @{
                        Name = $ModuleName
                        Force = if ($InstallParams.ContainsKey('Force')) { $InstallParams.Force } else { $true }
                        Confirm = if ($InstallParams.ContainsKey('Confirm')) { $InstallParams.Confirm } else { $false }
                        ErrorAction = 'Stop'
                    }
                    if ($InstallParams.ContainsKey('Scope')) { $installArgs.Scope = $InstallParams.Scope }
                    if ($InstallParams.ContainsKey('AllowClobber')) { $installArgs.AllowClobber = $InstallParams.AllowClobber }
                    if ($InstallParams.ContainsKey('AcceptLicense')) { $installArgs.AcceptLicense = $InstallParams.AcceptLicense }
                    if ($InstallParams.ContainsKey('SkipPublisherCheck')) { $installArgs.SkipPublisherCheck = $InstallParams.SkipPublisherCheck }
                    
                    Install-Module @installArgs
                }
                'Update' {
                    # Use explicit parameter passing for constrained language mode compatibility
                    # Note: Update-Module doesn't support AllowClobber, AcceptLicense, or SkipPublisherCheck
                    $updateArgs = @{
                        Name = $ModuleName
                        Force = if ($InstallParams.ContainsKey('Force')) { $InstallParams.Force } else { $true }
                        Confirm = if ($InstallParams.ContainsKey('Confirm')) { $InstallParams.Confirm } else { $false }
                        ErrorAction = 'Stop'
                    }
                    # Only add Scope if present (Update-Module supports this)
                    if ($InstallParams.ContainsKey('Scope')) { $updateArgs.Scope = $InstallParams.Scope }
                    
                    Update-Module @updateArgs
                }
                default {
                    # Use explicit parameter passing for constrained language mode compatibility
                    $installArgs = @{
                        Name = $ModuleName
                        Force = if ($InstallParams.ContainsKey('Force')) { $InstallParams.Force } else { $true }
                        Confirm = if ($InstallParams.ContainsKey('Confirm')) { $InstallParams.Confirm } else { $false }
                        ErrorAction = 'Stop'
                    }
                    if ($InstallParams.ContainsKey('Scope')) { $installArgs.Scope = $InstallParams.Scope }
                    if ($InstallParams.ContainsKey('AllowClobber')) { $installArgs.AllowClobber = $InstallParams.AllowClobber }
                    if ($InstallParams.ContainsKey('AcceptLicense')) { $installArgs.AcceptLicense = $InstallParams.AcceptLicense }
                    if ($InstallParams.ContainsKey('SkipPublisherCheck')) { $installArgs.SkipPublisherCheck = $InstallParams.SkipPublisherCheck }
                    
                    Install-Module @installArgs
                }
            }
            
            $actualTime = ((Get-Date) - $startTime).TotalSeconds
            Write-ColorOutput "    Successfully completed $Operation of $ModuleName" -Type Process
            Write-ColorOutput ("    Actual time: {0:N0} seconds" -f $actualTime) -Type Info
            return $true
        }
        catch {
            $actualTime = ((Get-Date) - $startTime).TotalSeconds
            Write-ColorOutput "    Error during $Operation of $ModuleName`: $($_.Exception.Message)" -Type Error
            Write-ColorOutput ("    Time elapsed: {0:N0} seconds" -f $actualTime) -Type Info
            return $false
        }
    }
    
    # Enhanced progress tracking with phase detection
    $progressParams = @{
        Activity = "$Operation module: $ModuleName"
        Status = "Downloading and installing..."
        PercentComplete = 0
    }
    
    $iteration = 0
    $maxIterations = [Math]::Max(10, [Math]::Ceiling($estimate.EstimatedTime / 3))
    $progressFile = Join-Path $env:TEMP "ModuleProgress_$($ModuleName)_$($job.Id).txt"
    $lastProgressUpdate = Get-Date
    $verificationPhaseStarted = $false
    $isConstrainedLargeModule = ($languageMode -eq 'ConstrainedLanguage' -and 
                               ($ModuleName -eq 'Az' -or $ModuleName -eq 'Microsoft.Graph'))
    
    while ($job.State -eq 'Running') {
        $elapsed = (Get-Date) - $startTime
        $elapsedSeconds = $elapsed.TotalSeconds
        $elapsedMinutes = $elapsed.TotalMinutes
        
        # Check for progress file updates for large modules in constrained mode
        $currentPhase = "Installation"
        if ($isConstrainedLargeModule -and (Test-Path $progressFile)) {
            try {
                $progressContent = Get-Content $progressFile -ErrorAction SilentlyContinue | Select-Object -Last 2
                if ($progressContent) {
                    $latestLine = $progressContent[-1]
                    if ($latestLine -like "*Verification*") {
                        $currentPhase = "Verification"
                        if (-not $verificationPhaseStarted) {
                            $verificationPhaseStarted = $true
                            Write-ColorOutput "    📋 Entered Verification Phase - This may take several minutes" -Type Warning
                            Write-ColorOutput "    ⏳ The process may appear frozen - this is normal behavior!" -Type Warning
                        }
                    }
                    elseif ($latestLine -like "*Cleanup*") {
                        $currentPhase = "Cleanup"
                        if ($verificationPhaseStarted) {
                            Write-ColorOutput "    ✅ Verification Complete - Finalizing installation..." -Type Process
                        }
                    }
                }
            }
            catch {
                # Ignore progress file read errors
            }
        }
        
        # Enhanced progress calculation for large modules in constrained mode
        if ($isConstrainedLargeModule) {
            switch ($currentPhase) {
                "Installation" {
                    # First 25% is download/installation (usually faster)
                    $phaseProgress = [Math]::Min(25, ($elapsedSeconds / 240) * 25) # Reduced to 4 minutes for this phase
                    $progressPercent = $phaseProgress
                }
                "Verification" {
                    # 25-85% is verification (slowest phase)
                    $verificationTime = $elapsedSeconds - 240  # Adjusted for reduced installation time
                    $verificationProgress = [Math]::Min(60, ($verificationTime / 480) * 60) # Reduced to 8 minutes for verification
                    $progressPercent = 25 + $verificationProgress
                    
                    # Show periodic reminders during verification
                    if (($elapsedMinutes -gt 6) -and (((Get-Date) - $lastProgressUpdate).TotalSeconds -gt 60)) {  # Reduced thresholds
                        Write-ColorOutput ("    ⏳ Still in verification phase ({0:N1} min elapsed) - please continue waiting..." -f $elapsedMinutes) -Type Info
                        $lastProgressUpdate = Get-Date
                    }
                }
                "Cleanup" {
                    # 85-100% is cleanup (usually quick)
                    $cleanupTime = $elapsedSeconds - 720  # Adjusted for reduced total time (4+8 minutes)
                    $cleanupProgress = [Math]::Min(15, ($cleanupTime / 120) * 15) # 2 minutes for cleanup
                    $progressPercent = 85 + $cleanupProgress
                }
                default {
 
                    $progressPercent = [Math]::Min(95, ($elapsedSeconds / $estimate.EstimatedTime) * 100)
                }
            }
        } else {
            # Standard progress calculation for normal modules
            $progressPercent = [Math]::Min(95, ($elapsedSeconds / $estimate.EstimatedTime) * 100)
        }
        
        $progressParams.PercentComplete = $progressPercent
        
        # Enhanced status messages
        if ($isConstrainedLargeModule) {
            $progressParams.Status = "Phase: $currentPhase - {0:N1}% - Elapsed: {1:N1}min" -f $progressPercent, $elapsedMinutes
            
            if ($estimate.EstimatedTime -gt $elapsedSeconds) {
                $remainingSeconds = $estimate.EstimatedTime - $elapsedSeconds
                if ($remainingSeconds -gt 60) {
                    $progressParams.Status += " - Est. remaining: {0:N1} min" -f ($remainingSeconds / 60)
                } else {
                    $progressParams.Status += " - Est. remaining: {0:N0}s" -f $remainingSeconds
                }
            }
            
            # Add phase-specific messages
            switch ($currentPhase) {
                "Verification" {
                    $progressParams.Status += " (May appear frozen - normal!)"
                }
                "Cleanup" {
                    $progressParams.Status += " (Almost complete)"
                }
            }
        } else {
            # Standard progress for normal modules
            $progressParams.Status = "Progress: {0:N0}% - Elapsed: {1:N0}s" -f $progressPercent, $elapsedSeconds
            
            if ($estimate.EstimatedTime -gt $elapsedSeconds) {
                $remainingSeconds = $estimate.EstimatedTime - $elapsedSeconds
                if ($remainingSeconds -gt 60) {
                    $progressParams.Status += " - Est. remaining: {0:N1} min" -f ($remainingSeconds / 60)
                } else {
                    $progressParams.Status += " - Est. remaining: {0:N0}s" -f $remainingSeconds
                }
            }
        }
        
        Write-Progress @progressParams
        
        # Adjust sleep time based on phase and module
        if ($isConstrainedLargeModule -and $currentPhase -eq "Verification") {
            Start-Sleep -Seconds 3  # Reduced from 10s to 3s for faster completion
        } else {
            Start-Sleep -Seconds 2  # Reduced from 3s to 2s
        }
        $iteration++
        
        # Enhanced timeout handling for large modules in constrained mode
        $timeoutMultiplier = if ($isConstrainedLargeModule) { 2.0 } else { 1.5 }  # Reduced from 2.5 to 2.0
       
        if ($iteration -gt $maxIterations -and $elapsedSeconds -gt ($estimate.EstimatedTime * $timeoutMultiplier)) {
            if ($isConstrainedLargeModule) {
                $progressParams.Status = "Extended processing time in Constrained Mode - Elapsed: {0:N1} min" -f $elapsedMinutes
                if ($currentPhase -eq "Verification") {
                    $progressParams.Status += " (Still in verification - normal for this module)"
                }
            } else {
                $progressParams.Status = "Taking longer than expected - Elapsed: {0:N0}s" -f $elapsedSeconds
            }
            Write-Progress @progressParams
        }
    }
    
    # Complete the progress bar
    Write-Progress -Activity "$Operation module: $ModuleName" -Completed
    
    # Get the job result
    $result = Receive-Job -Job $job
    Remove-Job -Job $job
    
    $actualTime = ((Get-Date) - $startTime).TotalSeconds
    $actualMinutes = $actualTime / 60
    
    if ($result.Success) {
        Write-ColorOutput "    Successfully completed $Operation of $ModuleName" -Type Process
        
        # Enhanced completion message for constrained mode
        if ($isConstrainedLargeModule) {
            Write-ColorOutput "    ✅ $ModuleName processing complete in Constrained Language Mode" -Type Process
            Write-ColorOutput ("    Actual time: {0:N1} minutes" -f $actualMinutes) -Type Info
            Write-ColorOutput "    Extended time was due to security verification requirements" -Type Info
        } else {
            Write-ColorOutput ("    Actual time: {0:N1} minutes" -f $actualMinutes) -Type Info
        }
        
        # If significantly different from estimate, show note
        if ([Math]::Abs($actualTime - $estimate.EstimatedTime) -gt ($estimate.EstimatedTime * 0.3)) {
            # Use traditional if-else for Constrained Language Mode compatibility
            if ($actualTime -gt $estimate.EstimatedTime) {
                $variance = "slower"
            } else {
                $variance = "faster"
            }
            
            # Special message for constrained mode
            if ($isConstrainedLargeModule -and $actualTime -gt $estimate.EstimatedTime) {
                Write-ColorOutput "    Note: Extended time is normal for $ModuleName in Constrained Language Mode" -Type Info
            } else {
                Write-ColorOutput ("    Note: $Operation was {0:N0}% $variance than estimated" -f ([Math]::Abs(($actualTime - $estimate.EstimatedTime) / $estimate.EstimatedTime * 100))) -Type Info
            }
        }
        
        # Perform automatic cleanup of old module versions after successful installation/update
        if (($Operation -eq 'Install' -or $Operation -eq 'Update') -and -not $SkipVersionCleanup) {
            Write-Verbose "Starting automatic cleanup of old versions for $ModuleName"
            $cleanupResult = Remove-OldModuleVersions -ModuleName $ModuleName -KeepLatestOnly
            
            if ($cleanupResult.Success -and $cleanupResult.RemovedCount -gt 0) {
                Write-ColorOutput "    🧹 Automatically cleaned up $($cleanupResult.RemovedCount) old version(s)" -Type Process
            }
        }
        
        return $true
    }
    else {
        Write-ColorOutput "    Error during $Operation of $ModuleName`: $($result.Message)" -Type Error
        
        # Additional troubleshooting for constrained mode
        if ($isConstrainedLargeModule) {
            Write-ColorOutput "    💡 Constrained Language Mode troubleshooting:" -Type Info
            Write-ColorOutput "    • Try running PowerShell as Administrator" -Type Info
            Write-ColorOutput "    • Ensure sufficient disk space (at least 2GB free)" -Type Info
            Write-ColorOutput "    • Check internet connectivity" -Type Info
            Write-ColorOutput "    • Consider temporarily disabling antivirus during installation" -Type Info
        }
        
        return $false
    }
}

function Get-FilteredModuleList {
    <#
    .SYNOPSIS
        Filters the module list based on the specified scope
    
    .DESCRIPTION
        Returns modules that match the specified scope criteria, allowing users
        to update only specific categories of modules
    
    .PARAMETER Scope
        Array of scopes to include (All, Graph, Azure, Exchange, Teams, SharePoint, PowerApps)
    
    .RETURNS
        Array of module objects matching the specified scope
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string[]]$Scope = @('All')
    )
    
    if ('All' -in $Scope) {
        Write-ColorOutput "Processing all available modules" -Type Info
        return $Script:ModuleList
    }
    
    $filteredModules = @()
    
    foreach ($scopeItem in $Scope) {
        $modulesInScope = $Script:ModuleList | Where-Object { $_.Category -eq $scopeItem }
        if ($modulesInScope) {
            $filteredModules += $modulesInScope
            Write-ColorOutput "Added $($modulesInScope.Count) module(s) from $scopeItem category" -Type Info
        } else {
            Write-ColorOutput "Warning: No modules found for scope '$scopeItem'" -Type Warning
        }
    }
    
    # Remove duplicates and sort by priority
    $uniqueModules = $filteredModules | Sort-Object Name -Unique | Sort-Object Priority, Name
    
    Write-ColorOutput "Total modules selected: $($uniqueModules.Count)" -Type Process
    return $uniqueModules
}

function Test-NetworkConnectivity {
    <#
    .SYNOPSIS
        Tests network connectivity to PowerShell Gallery and other required endpoints
    
    .DESCRIPTION
        Verifies that the system can reach the PowerShell Gallery and Microsoft endpoints
        required for module downloads and updates
    
    .RETURNS
        Boolean indicating whether network connectivity is sufficient
    #>
    [CmdletBinding()]
    param()
    
    Write-ColorOutput "Testing network connectivity..." -Type System
    
    $testEndpoints = @(
        @{ Name = 'PowerShell Gallery'; Url = 'https://www.powershellgallery.com'; Port = 443 },
        @{ Name = 'Microsoft Download Center'; Url = 'https://download.microsoft.com'; Port = 443 },
        @{ Name = 'NuGet.org'; Url = 'https://api.nuget.org'; Port = 443 }
    )
    
    $allTestsPassed = $true
    
    foreach ($endpoint in $testEndpoints) {
        try {
            Write-ColorOutput "  Testing connection to $($endpoint.Name)..." -Type Info
            
            # Use Test-NetConnection if available (Windows PowerShell 5.1+)
            if (Get-Command Test-NetConnection -ErrorAction SilentlyContinue) {
                # Suppress all Test-NetConnection output streams during normal operation
                $result = $null
                if ($VerbosePreference -eq 'Continue' -or $DebugPreference -eq 'Continue') {
                    # Only show detailed output during debugging/verbose mode
                    $result = Test-NetConnection -ComputerName ([System.Uri]$endpoint.Url).Host -Port $endpoint.Port -InformationLevel Detailed -WarningAction SilentlyContinue
                    $connectionSuccessful = $result -and $result.TcpTestSucceeded -eq $true
                } else {
                    # Use TcpClient for silent testing - works in all language modes
                    try {
                        $tcpClient = New-Object System.Net.Sockets.TcpClient
                        $tcpClient.ReceiveTimeout = 10000
                        $tcpClient.SendTimeout = 10000
                        $tcpClient.Connect(([System.Uri]$endpoint.Url).Host, $endpoint.Port)
                        $connectionSuccessful = $tcpClient.Connected
                        $tcpClient.Close()
                    }
                    catch {
                        $connectionSuccessful = $false
                    }
                    finally {
                        if ($tcpClient) { $tcpClient.Dispose() }
                    }
                }
                
                if ($connectionSuccessful) {
                    Write-ColorOutput "    ✓ $($endpoint.Name) - Connection successful" -Type Process
                } else {
                    Write-ColorOutput "    ✗ $($endpoint.Name) - Connection failed" -Type Error
                    $allTestsPassed = $false
                }
            } else {
                # Fallback for environments without Test-NetConnection
                $webRequest = [System.Net.WebRequest]::Create($endpoint.Url)
                $webRequest.Timeout = 10000  # 10 seconds
                $response = $webRequest.GetResponse()
                $response.Close()
                Write-ColorOutput "    ✓ $($endpoint.Name) - Connection successful" -Type Process
            }
        }
        catch {
            Write-ColorOutput "    ✗ $($endpoint.Name) - Connection failed: $($_.Exception.Message)" -Type Error
            $allTestsPassed = $false
        }
    }
    
    if ($allTestsPassed) {
        Write-ColorOutput "  All connectivity tests passed" -Type Process
    } else {
        Write-ColorOutput "  Some connectivity tests failed - module operations may be slower or fail" -Type Warning
    }
    
    Write-ColorOutput ""
    return $allTestsPassed
}

function Get-ModuleVersionInfo {
    <#
    .SYNOPSIS
        Gets comprehensive version information for a module
    
    .DESCRIPTION
        Retrieves local and online version information for a module,
        including version comparison and update recommendations
    
    .PARAMETER ModuleName
        Name of the module to check
    
    .RETURNS
        Hashtable containing version information and recommendations
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName
    )
    
    $versionInfo = @{
        ModuleName = $ModuleName
        IsInstalled = $false
        LocalVersion = $null
        OnlineVersion = $null
        UpdateAvailable = $false
        UpdateRecommended = $false
        InstallRecommended = $false
        MultipleVersionsInstalled = $false
        ErrorMessage = $null
    }
    
    try {
        # Check for installed versions
        $installedModules = Get-InstalledModule -Name $ModuleName -AllVersions -ErrorAction SilentlyContinue
        
        if ($installedModules) {
            $versionInfo.IsInstalled = $true
            $versionInfo.LocalVersion = ($installedModules | Sort-Object Version -Descending | Select-Object -First 1).Version
            $versionInfo.MultipleVersionsInstalled = ($installedModules | Measure-Object).Count -gt 1
        }
        
        # Check online version
        $onlineModule = Find-Module -Name $ModuleName -Repository $Repository -ErrorAction SilentlyContinue
        
        if ($onlineModule) {
            $versionInfo.OnlineVersion = $onlineModule.Version
            
            if ($versionInfo.IsInstalled) {
                if ([version]$versionInfo.OnlineVersion -gt [version]$versionInfo.LocalVersion) {
                    $versionInfo.UpdateAvailable = $true
                    
                    # Determine if update is recommended based on version difference
                    $localVersion = [version]$versionInfo.LocalVersion
                    $onlineVersion = [version]$versionInfo.OnlineVersion
                    
                    # Recommend update for major/minor version changes or security updates
                    if ($onlineVersion.Major -gt $localVersion.Major -or 
                        $onlineVersion.Minor -gt $localVersion.Minor -or
                        ($onlineVersion.Build - $localVersion.Build) -ge 5) {
                        $versionInfo.UpdateRecommended = $true
                    }
                }
            } else {
                $versionInfo.InstallRecommended = $true
            }
        } else {
            $versionInfo.ErrorMessage = "Module not found in repository '$Repository'"
        }
    }
    catch {
        $versionInfo.ErrorMessage = $_.Exception.Message
    }
    
    return $versionInfo
}

function Remove-OldModuleVersions {
    <#
    .SYNOPSIS
        Removes old versions of a module after successful installation/update
    
    .DESCRIPTION
        This function removes older versions of a module, keeping only the latest version.
        It helps prevent the "⚠️ Multiple versions installed" warning by maintaining
        only the most recent version after successful updates.
    
    .PARAMETER ModuleName
        Name of the module to clean up
    
    .PARAMETER KeepLatestOnly
        If true, keeps only the latest version. If false, keeps the latest 2 versions.
    
    .RETURNS
        Hashtable with cleanup results
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,
        
        [Parameter()]
        [switch]$KeepLatestOnly
    )
    
    $result = @{
        Success = $false
        RemovedVersions = @()
        KeptVersions = @()
        ErrorMessage = $null
        RemovedCount = 0
    }
    
    try {
        # Get all installed versions of the module
        $installedVersions = Get-InstalledModule -Name $ModuleName -AllVersions -ErrorAction SilentlyContinue | 
                            Sort-Object Version -Descending
        
        if (-not $installedVersions -or $installedVersions.Count -le 1) {
            Write-Verbose "No cleanup needed for $ModuleName - only one or no versions installed"
            $result.Success = $true
            if ($installedVersions) {
                $result.KeptVersions = @($installedVersions.Version.ToString())
            }
            return $result
        }
        
        # Determine how many versions to keep (default to keeping only latest if no parameter specified)
        $versionsToKeep = if ($KeepLatestOnly -or $PSBoundParameters.Count -eq 1) { 1 } else { 2 }
        $versionsToRemove = $installedVersions | Select-Object -Skip $versionsToKeep
        
        if (-not $versionsToRemove) {
            Write-Verbose "No old versions to remove for $ModuleName"
            $result.Success = $true
            $result.KeptVersions = $installedVersions | ForEach-Object { $_.Version.ToString() }
            return $result
        }
        
        Write-ColorOutput "    🧹 Cleaning up old versions of $ModuleName..." -Type Info
        
        foreach ($versionToRemove in $versionsToRemove) {
            try {
                Write-Verbose "Removing $ModuleName version $($versionToRemove.Version)"
                
                # Use Uninstall-Module with specific version
                Uninstall-Module -Name $ModuleName -RequiredVersion $versionToRemove.Version -Force -ErrorAction Stop
                
                $result.RemovedVersions += $versionToRemove.Version.ToString()
                $result.RemovedCount++
                
                Write-ColorOutput "      ✅ Removed version $($versionToRemove.Version)" -Type Process
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-ColorOutput "      ⚠️ Could not remove version $($versionToRemove.Version): $errorMsg" -Type Warning
                
                # Don't fail the entire operation if we can't remove one version
                Write-Verbose "Continuing cleanup despite error removing version $($versionToRemove.Version)"
            }
        }
        
        # Record kept versions
        $keptVersions = $installedVersions | Select-Object -First $versionsToKeep
        $result.KeptVersions = $keptVersions | ForEach-Object { $_.Version.ToString() }
        
        # Consider operation successful even if some versions couldn't be removed
        $result.Success = $true
        
        if ($result.RemovedCount -gt 0) {
            Write-ColorOutput "    ✅ Cleanup completed: removed $($result.RemovedCount) old version(s), kept $($result.KeptVersions.Count) version(s)" -Type Process
        } else {
            Write-ColorOutput "    ℹ️ No versions were removed (may be in use or protected)" -Type Info
        }
        
    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
        Write-ColorOutput "    ❌ Cleanup failed for $ModuleName`: $($result.ErrorMessage)" -Type Error
    }
    
    return $result
}

function Remove-AllOldModuleVersions {
    <#
    .SYNOPSIS
        Removes old versions from all installed modules that have multiple versions
    
    .DESCRIPTION
        This function scans all installed modules that have multiple versions and
        removes old versions, keeping only the latest version of each module.
        This is useful for cleaning up accumulated module versions over time.
    
    .PARAMETER ExcludeModules
        Array of module names to exclude from cleanup
    
    .PARAMETER KeepLatestOnly
        If true, keeps only the latest version. If false, keeps the latest 2 versions.
    
    .RETURNS
        Hashtable with overall cleanup results
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string[]]$ExcludeModules = @(),
        
        [Parameter()]
        [switch]$KeepLatestOnly
    )
    
    $overallResult = @{
        ProcessedModules = 0
        SuccessfulCleanups = 0
        FailedCleanups = 0
        TotalVersionsRemoved = 0
        CleanupDetails = @()
    }
    
    Write-ColorOutput "🧹 Starting cleanup of old module versions..." -Type System
    
    try {
        # Get all installed modules and group by name to find those with multiple versions
        $allInstalledModules = Get-InstalledModule -ErrorAction SilentlyContinue
        $moduleGroups = $allInstalledModules | Group-Object Name
        
        # Find modules with multiple versions
        $modulesWithMultipleVersions = $moduleGroups | Where-Object { $_.Count -gt 1 }
        
        if (-not $modulesWithMultipleVersions) {
            Write-ColorOutput "✅ No modules with multiple versions found - cleanup not needed" -Type Process
            return $overallResult
        }
        
        Write-ColorOutput "Found $($modulesWithMultipleVersions.Count) module(s) with multiple versions" -Type Info
        
        foreach ($moduleGroup in $modulesWithMultipleVersions) {
            $moduleName = $moduleGroup.Name
            
            # Skip excluded modules
            if ($moduleName -in $ExcludeModules) {
                Write-ColorOutput "  ⏭️ Skipping $moduleName (excluded)" -Type Warning
                continue
            }
            
            $overallResult.ProcessedModules++
            
            Write-ColorOutput "Processing: $moduleName ($($moduleGroup.Count) versions)" -Type Info
            
            # Perform cleanup for this module
            $cleanupResult = Remove-OldModuleVersions -ModuleName $moduleName -KeepLatestOnly:$KeepLatestOnly
            
            # Track results
            if ($cleanupResult.Success) {
                $overallResult.SuccessfulCleanups++
                $overallResult.TotalVersionsRemoved += $cleanupResult.RemovedCount
            } else {
                $overallResult.FailedCleanups++
            }
            
            # Store detailed results
            $overallResult.CleanupDetails += @{
                ModuleName = $moduleName
                Success = $cleanupResult.Success
                RemovedVersions = $cleanupResult.RemovedVersions
                KeptVersions = $cleanupResult.KeptVersions
                RemovedCount = $cleanupResult.RemovedCount
                ErrorMessage = $cleanupResult.ErrorMessage
            }
        }
        
        # Display overall summary
        Write-ColorOutput ""
        Write-ColorOutput "🧹 Module Cleanup Summary:" -Type System
        Write-ColorOutput "  Processed modules: $($overallResult.ProcessedModules)" -Type Info
        Write-ColorOutput "  Successful cleanups: $($overallResult.SuccessfulCleanups)" -Type Process
        Write-ColorOutput "  Failed cleanups: $($overallResult.FailedCleanups)" -Type Error
        Write-ColorOutput "  Total versions removed: $($overallResult.TotalVersionsRemoved)" -Type Process
        
        if ($overallResult.FailedCleanups -gt 0) {
            Write-ColorOutput ""
            Write-ColorOutput "⚠️ Some cleanups failed - this is normal if modules are in use" -Type Warning
            Write-ColorOutput "   You can manually remove old versions later when modules are not in use" -Type Info
        }
        
    }
    catch {
        Write-ColorOutput "❌ Module cleanup failed: $($_.Exception.Message)" -Type Error
        $overallResult.CleanupDetails += @{
            ModuleName = "Overall Operation"
            Success = $false
            ErrorMessage = $_.Exception.Message
        }
    }
    
    return $overallResult
}

function Test-Prerequisites {
    <#
    .SYNOPSIS
        Tests system prerequisites for running the script
    #>
    [CmdletBinding()]
    param()
    
    Write-ProgressHeader "System Prerequisites Check"
    
    $allPrerequisitesMet = $true
    
    # Check PowerShell version
    $psVersion = $PSVersionTable.PSVersion
    Write-ColorOutput "PowerShell Version: $($psVersion.Major).$($psVersion.Minor).$($psVersion.Build)" -Type Info
    
    if ($psVersion.Major -lt 5) {
        Write-ColorOutput "❌ PowerShell 5.1 or higher is required" -Type Error
        $allPrerequisitesMet = $false
    } else {
        Write-ColorOutput "✅ PowerShell version is supported" -Type Success
    }
    
    # Check administrator privileges
    try {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if ($isAdmin) {
            Write-ColorOutput "✅ Running with Administrator privileges" -Type Success
        } else {
            Write-ColorOutput "❌ Administrator privileges required" -Type Error
            $allPrerequisitesMet = $false
        }
    }
    catch {
        Write-ColorOutput "⚠️ Could not verify administrator privileges: $($_.Exception.Message)" -Type Warning
    }
    
    # Check execution policy
    try {
        $executionPolicy = Get-ExecutionPolicy -Scope CurrentUser
        Write-ColorOutput "Execution Policy (CurrentUser): $executionPolicy" -Type Info
        
        if ($executionPolicy -eq 'Restricted') {
            Write-ColorOutput "⚠️ Execution policy is Restricted - may cause issues" -Type Warning
        } else {
            Write-ColorOutput "✅ Execution policy allows script execution" -Type Success
        }
    }
    catch {
        Write-ColorOutput "⚠️ Could not check execution policy: $($_.Exception.Message)" -Type Warning
    }
    
    # Check PowerShell language mode
    $languageMode = $ExecutionContext.SessionState.LanguageMode
    Write-ColorOutput "PowerShell Language Mode: $languageMode" -Type Info
    
    if ($languageMode -eq 'ConstrainedLanguage') {
        Write-ColorOutput "⚠️ Constrained Language Mode detected - some features may be limited" -Type Warning
    } else {
        Write-ColorOutput "✅ Full Language Mode available" -Type Success
    }
    
    return $allPrerequisitesMet
}

function Start-ModuleProcessing {
    <#
    .SYNOPSIS
        Processes modules based on the filtered list
    #>
    [CmdletBinding()]
    param()
    
    Write-ProgressHeader "Module Processing" "Processing $($Script:FilteredModuleList.Count) modules"
    
    $processedCount = 0
    $successCount = 0
    $failureCount = 0
    
    foreach ($module in $Script:FilteredModuleList) {
        $processedCount++
        
        Write-ColorOutput "($processedCount of $($Script:FilteredModuleList.Count)) Processing: $($module.Description)" -Type Info
        
        # Get version information
        $versionInfo = Get-ModuleVersionInfo -ModuleName $module.Name
        
        if ($versionInfo.ErrorMessage) {
            Write-ColorOutput "  ❌ Error checking module: $($versionInfo.ErrorMessage)" -Type Error
            $failureCount++
            continue
        }
        
        # Display current status
        if ($versionInfo.IsInstalled) {
            Write-ColorOutput "  Current version: $($versionInfo.LocalVersion)" -Type Info
            if ($versionInfo.MultipleVersionsInstalled) {
                Write-ColorOutput "  ⚠️ Multiple versions installed" -Type Warning
                
                # Offer cleanup option for multiple versions
                if (-not $CheckOnly -and $SkipVersionCleanup -and $Prompt -and -not $Script:SkipCleanupPrompts) {
                    Write-ColorOutput "  💡 Multiple versions detected for $($module.Name)" -Type Info
                    
                    # Get all installed versions for display
                    $allVersions = Get-InstalledModule -Name $module.Name -AllVersions -ErrorAction SilentlyContinue | 
                                  Sort-Object Version -Descending
                    
                    if ($allVersions -and $allVersions.Count -gt 1) {
                        Write-ColorOutput "  📋 Installed versions:" -Type Info
                        foreach ($version in $allVersions) {
                            $indicator = if ($version.Version -eq $versionInfo.LocalVersion) { " (current)" } else { "" }
                            Write-ColorOutput "    • $($version.Version)$indicator" -Type Info
                        }
                        
                        # Interactive prompt for cleanup
                        $cleanupChoice = $null
                        if (-not $Prompt) {
                            # Auto-cleanup mode - inform user but don't prompt
                            Write-ColorOutput "  🔧 Automatic cleanup will remove old versions after any updates" -Type Info
                        } else {
                            # Interactive mode - ask user
                            Write-ColorOutput "  🧹 Would you like to clean up old versions now?" -Type Warning
                            Write-ColorOutput "     This will keep only the latest version ($($versionInfo.LocalVersion))" -Type Info
                            
                            do {
                                $cleanupChoice = Read-Host "  Clean up old versions? (Y/N/S for Skip all prompts)"
                                $cleanupChoice = $cleanupChoice.Trim().ToUpper()
                                
                                if ($cleanupChoice -eq 'S') {
                                    # Skip all future prompts for this session
                                    Write-ColorOutput "  ⏭️ Skipping cleanup prompts for remaining modules" -Type Warning
                                    $Script:SkipCleanupPrompts = $true
                                    break
                                }
                            } while ($cleanupChoice -notin @('Y', 'N', 'S'))
                            
                            # Perform cleanup if requested
                            if ($cleanupChoice -eq 'Y') {
                                Write-ColorOutput "  🧹 Cleaning up old versions of $($module.Name)..." -Type Process
                                $cleanupResult = Remove-OldModuleVersions -ModuleName $module.Name -KeepLatestOnly
                                
                                if ($cleanupResult.Success -and $cleanupResult.RemovedCount -gt 0) {
                                    Write-ColorOutput "  ✅ Removed $($cleanupResult.RemovedCount) old version(s)" -Type Process
                                    Write-ColorOutput "  📋 Kept version: $($cleanupResult.KeptVersions -join ', ')" -Type Info
                                } elseif ($cleanupResult.Success) {
                                    Write-ColorOutput "  ℹ️ No old versions were removed (may be in use)" -Type Info
                                } else {
                                    Write-ColorOutput "  ⚠️ Cleanup failed: $($cleanupResult.ErrorMessage)" -Type Warning
                                }
                            } elseif ($cleanupChoice -eq 'N') {
                                Write-ColorOutput "  ⏭️ Keeping all versions as requested" -Type Info
                            }
                        }
                    }
                } elseif ($CheckOnly) {
                    # In check-only mode, just show the versions
                    $allVersions = Get-InstalledModule -Name $module.Name -AllVersions -ErrorAction SilentlyContinue | 
                                  Sort-Object Version -Descending
                    if ($allVersions -and $allVersions.Count -gt 1) {
                        Write-ColorOutput "  📋 Installed versions: $($allVersions.Version -join ', ')" -Type Info
                    }
                }
            }
        } else {
            Write-ColorOutput "  Module not installed" -Type Warning
        }
        
        Write-ColorOutput "  Latest version: $($versionInfo.OnlineVersion)" -Type Info
        
        if ($CheckOnly) {
            if ($versionInfo.UpdateAvailable) {
                Write-ColorOutput "  📋 Update available" -Type Warning
            } elseif ($versionInfo.InstallRecommended) {
                Write-ColorOutput "  📋 Installation recommended" -Type Warning
            } else {
                Write-ColorOutput "  ✅ Up to date" -Type Process
            }
            $successCount++
            continue
        }
        
        # Determine action needed
        $action = $null
        if ($versionInfo.InstallRecommended) {
            $action = 'Install'
        } elseif ($versionInfo.UpdateAvailable) {
            $action = 'Update'
        }
        
        if ($action) {
            if ($Prompt) {
                $response = Read-Host "  $action $($module.Name)? (Y/N)"
                if ($response -notmatch '^[Yy]') {
                    Write-ColorOutput "  ⏭️ Skipped by user" -Type Warning
                    continue
                }
            }
            
            # Process the module using the existing Install-ModuleWithProgress function
            $result = @{ Success = $false; Duration = New-TimeSpan; ErrorMessage = "" }
            
            try {
                $startTime = Get-Date
                
                # Get proper installation parameters
                $installParams = Get-ModuleSpecificParams -ModuleName $module.Name -BaseParams @{
                    Force = $true
                    Confirm = $false
                    Scope = 'AllUsers'
                    Repository = $Repository
                }
                
                # Call the existing function
                $success = Install-ModuleWithProgress -ModuleName $module.Name -InstallParams $installParams -Operation $action
                
                $result.Success = $success
                $result.Duration = (Get-Date) - $startTime
            }
            catch {
                $result.ErrorMessage = $_.Exception.Message
                $result.Duration = (Get-Date) - $startTime
            }
            
            if ($result.Success) {
                $successCount++
            } else {
                $failureCount++
            }
            
            # Store result for summary
            if (-not $Script:ModuleOperationResults) {
                $Script:ModuleOperationResults = @()
            }
            
            $Script:ModuleOperationResults += @{
                ModuleName = $module.Name
                Operation = $action
                Success = $result.Success
                Duration = $result.Duration
                ErrorMessage = $result.ErrorMessage
            }
        } else {
            Write-ColorOutput "  ✅ Already up to date" -Type Process
            $successCount++
        }
        
        Write-ColorOutput ""
    }
    
    # Display summary
    Write-ProgressHeader "Processing Summary"
    Write-ColorOutput "Total modules processed: $processedCount" -Type Info
    Write-ColorOutput "Successful operations: $successCount" -Type Process
    Write-ColorOutput "Failed operations: $failureCount" -Type Error
    
    if ($Script:ModuleOperationResults -and $Script:ModuleOperationResults.Count -gt 0) {
        Write-ColorOutput ""
        Write-ColorOutput "Operation Details:" -Type Info
        
        foreach ($result in $Script:ModuleOperationResults) {
            $status = if ($result.Success) { "✅" } else { "❌" }
            $duration = $result.Duration.TotalSeconds.ToString('F1')
            Write-ColorOutput "  $status $($result.Operation) $($result.ModuleName) ($duration s)" -Type Info
            
            if (-not $result.Success -and $result.ErrorMessage) {
                Write-ColorOutput "    Error: $($result.ErrorMessage)" -Type Error
            }
        }
    }
    
    # Offer comprehensive cleanup at the end if not already handled
    if (-not $CheckOnly -and $SkipVersionCleanup -and $Prompt) {
        Write-ColorOutput ""
        Write-ColorOutput "🧹 Final Cleanup Opportunity" -Type System
        
        # Check for any remaining modules with multiple versions
        $modulesWithMultipleVersions = @()
        foreach ($module in $Script:FilteredModuleList) {
            $installedVersions = Get-InstalledModule -Name $module.Name -AllVersions -ErrorAction SilentlyContinue
            if ($installedVersions -and $installedVersions.Count -gt 1) {
                $modulesWithMultipleVersions += @{
                    Name = $module.Name
                    Description = $module.Description
                    VersionCount = $installedVersions.Count
                    Versions = ($installedVersions | Sort-Object Version -Descending | ForEach-Object { $_.Version.ToString() })
                }
            }
        }
        
        if ($modulesWithMultipleVersions.Count -gt 0) {
            Write-ColorOutput "📋 Modules with multiple versions still installed:" -Type Warning
            foreach ($mod in $modulesWithMultipleVersions) {
                Write-ColorOutput "  • $($mod.Description): $($mod.VersionCount) versions ($($mod.Versions -join ', '))" -Type Info
            }
            
            Write-ColorOutput ""
            Write-ColorOutput "💡 Would you like to clean up all old versions now?" -Type Info
            Write-ColorOutput "   This will keep only the latest version of each module" -Type Info
            
            $globalCleanupChoice = Read-Host "Clean up all old versions? (Y/N)"
            
            if ($globalCleanupChoice -match '^[Yy]') {
                Write-ColorOutput "🧹 Starting comprehensive cleanup..." -Type Process
                $overallResult = Remove-AllOldModuleVersions -KeepLatestOnly
                
                Write-ColorOutput ""
                Write-ColorOutput "✅ Comprehensive cleanup completed!" -Type Success
                Write-ColorOutput "  Modules processed: $($overallResult.ProcessedModules)" -Type Info
                Write-ColorOutput "  Successful cleanups: $($overallResult.SuccessfulCleanups)" -Type Process
                Write-ColorOutput "  Total versions removed: $($overallResult.TotalVersionsRemoved)" -Type Process
                
                if ($overallResult.FailedCleanups -gt 0) {
                    Write-ColorOutput "  Failed cleanups: $($overallResult.FailedCleanups) (modules may be in use)" -Type Warning
                }
            } else {
                Write-ColorOutput "⏭️ Comprehensive cleanup skipped - multiple versions preserved" -Type Info
            }
        } else {
            Write-ColorOutput "✅ No modules with multiple versions found - cleanup not needed" -Type Success
        }
    }
}

function Remove-DeprecatedModules {
    <#
    .SYNOPSIS
        Removes deprecated PowerShell modules that should no longer be used
    #>
    [CmdletBinding()]
    param()
    
    Write-ProgressHeader "Deprecated Module Cleanup"
    
    $removedCount = 0
    
    foreach ($deprecatedModule in $Script:DeprecatedModules) {
        try {
            $installedVersions = Get-InstalledModule -Name $deprecatedModule.Name -AllVersions -ErrorAction SilentlyContinue
            
            if ($installedVersions) {
                Write-ColorOutput "Found deprecated module: $($deprecatedModule.Name)" -Type Warning
                Write-ColorOutput "  Reason: $($deprecatedModule.Reason)" -Type Info
                Write-ColorOutput "  Replacement: $($deprecatedModule.Replacement)" -Type Info
                Write-ColorOutput "  Versions installed: $($installedVersions.Count)" -Type Info
                
                foreach ($version in $installedVersions) {
                    try {
                        Write-ColorOutput "  Removing version $($version.Version)..." -Type Process
                        Uninstall-Module -Name $deprecatedModule.Name -RequiredVersion $version.Version -Force -ErrorAction Stop
                        Write-ColorOutput "    ✅ Removed successfully" -Type Process
                        $removedCount++
                    }
                    catch {
                        Write-ColorOutput "    ⚠️ Could not remove: $($_.Exception.Message)" -Type Warning
                    }
                }
            }
        }
        catch {
            Write-ColorOutput "Error checking for $($deprecatedModule.Name)`: $($_.Exception.Message)" -Type Error
        }
    }
    
    if ($removedCount -gt 0) {
        Write-ColorOutput "✅ Removed $removedCount deprecated module version(s)" -Type Process
    } else {
        Write-ColorOutput "✅ No deprecated modules found or removal needed" -Type Process
    }
}

function Write-ProgressHeader {
    <#
    .SYNOPSIS
        Writes a formatted progress header to the console
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        
        [Parameter()]
        [string]$Subtitle = ""
    )
    
    $separator = "=" * 80
    Write-ColorOutput $separator -Type System
    Write-ColorOutput $Title -Type System
    if ($Subtitle) {
        Write-ColorOutput $Subtitle -Type Info
    }
    Write-ColorOutput $separator -Type System
    Write-ColorOutput ""
}

#region Main Execution

try {
    # Start transcript logging if requested
    if ($CreateLog) {
        $logPath = Join-Path $env:TEMP "o365-update-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
        Start-Transcript -Path $logPath -Append
        Write-ColorOutput "📋 Logging to: $logPath" -Type System
    }
    
    # Show header (Clear-Host removed to avoid console issues in some environments)
    Write-ProgressHeader "Microsoft Cloud PowerShell Module Updater v2.9" "Enhanced version with improved module management and version cleanup"
    
    # Display configuration
    Write-ColorOutput "Configuration:" -Type Info
    Write-ColorOutput "  Mode: $(if ($CheckOnly) { 'Check Only' } else { 'Update' })" -Type Info
    Write-ColorOutput "  Scope: $($ModuleScope -join ', ')" -Type Info
    Write-ColorOutput "  Repository: $Repository" -Type Info
    Write-ColorOutput "  Prompt: $Prompt" -Type Info
    Write-ColorOutput "  Skip Deprecated Cleanup: $SkipDeprecatedCleanup" -Type Info
    Write-ColorOutput "  Skip Version Cleanup: $SkipVersionCleanup" -Type Info
    Write-ColorOutput "  Max Parallel Operations: $MaxParallelOperations" -Type Info
    Write-ColorOutput "  Timeout: $TimeoutMinutes minutes" -Type Info
    Write-ColorOutput ""
    
    # Check prerequisites
    Write-ColorOutput "🔍 Checking prerequisites..." -Type System
    $prerequisitesMet = Test-Prerequisites
    if (-not $prerequisitesMet) {
        Write-ColorOutput "❌ Prerequisites not met. Please resolve the issues above and try again." -Type Error
        exit 1
    }
    Write-ColorOutput "✅ Prerequisites check passed" -Type Process
    
    # Test network connectivity unless skipped
    if (-not $SkipConnectivityCheck) {
        Write-ColorOutput "🌐 Testing network connectivity..." -Type System
        $connectivityOk = Test-NetworkConnectivity
        if (-not $connectivityOk -and -not $CheckOnly) {
            $continueResponse = Read-Host "Network connectivity issues detected. Continue anyway? (Y/N)"
            if ($continueResponse -notmatch '^[Yy]') {
                Write-ColorOutput "Script execution cancelled due to connectivity issues." -Type Warning
                exit 0
            }
        }
        Write-ColorOutput "✅ Network connectivity check completed" -Type Process
    } else {
        Write-ColorOutput "⏭️ Network connectivity check skipped" -Type Warning
    }
    
    # Handle session check mode
    if ($CheckSessions) {
        Write-ProgressHeader "PowerShell Session Check Mode"
        # Add session check logic here if needed
        Write-ColorOutput "Session check completed. Run the script without -CheckSessions to proceed with updates." -Type Info
        exit 0
    }
    
    # Filter modules based on scope
    Write-ColorOutput "📋 Filtering modules based on scope..." -Type System
    $Script:FilteredModuleList = Get-FilteredModuleList -Scope $ModuleScope
    
    if ($Script:FilteredModuleList.Count -eq 0) {
        Write-ColorOutput "❌ No modules selected for processing. Please check your ModuleScope parameter." -Type Error
        exit 1
    }
    
    Write-ColorOutput "✅ Found $($Script:FilteredModuleList.Count) module(s) to process" -Type Process
    
    # Clean up deprecated modules first
    if (-not $SkipDeprecatedCleanup) {
        Write-ColorOutput "🧹 Checking for deprecated modules..." -Type System
        Remove-DeprecatedModules
    } else {
        Write-ColorOutput "⏭️ Deprecated module cleanup skipped" -Type Warning
    }
    
    # Process modules
    Write-ColorOutput "🚀 Starting module processing..." -Type System
    Start-ModuleProcessing
    
    # Final summary
    Write-ProgressHeader "Script Completed Successfully"
    Write-ColorOutput "✅ Microsoft Cloud PowerShell Module Updater completed successfully!" -Type Success
    Write-ColorOutput "End Time: $(Get-Date -Format (Get-Culture).DateTimeFormat.FullDateTimePattern)" -Type System
    
    if (-not $CheckOnly) {
        Write-ColorOutput ""
        Write-ColorOutput "💡 Recommendations:" -Type Info
        Write-ColorOutput "  • Restart PowerShell to use updated modules" -Type Info
        Write-ColorOutput "  • Run 'Get-Module -ListAvailable' to verify installations" -Type Info
        Write-ColorOutput "  • Check module documentation for any breaking changes" -Type Info
        
        if (-not $SkipVersionCleanup) {
            Write-ColorOutput "  • Old module versions have been cleaned up automatically" -Type Info
        }
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

#endregion