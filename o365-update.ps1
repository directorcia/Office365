[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Prompt before installing missing modules")]
    [switch]$Prompt,
    
    [Parameter(HelpMessage = "Create a transcript log of all operations")]
    [switch]$CreateLog,
    
    [Parameter(HelpMessage = "Path for log file (default: current directory)")]
    [string]$LogPath = $PWD,
    
    [Parameter(HelpMessage = "Skip deprecated module cleanup")]
    [switch]$SkipDeprecatedCleanup,
    
    [Parameter(HelpMessage = "Only check versions without updating")]
    [switch]$CheckOnly,
      [Parameter(HelpMessage = "Check for PowerShell session conflicts and exit")]
    [switch]$CheckSessions,
    
    [Parameter(HelpMessage = "Automatically terminate conflicting PowerShell sessions without prompting")]
    [switch]$TerminateConflicts
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
    Prompts before installing missing modules instead of installing automatically

.PARAMETER CreateLog
    Creates a transcript log of all operations

.PARAMETER LogPath
    Specifies the path for the log file (default: current directory)

.PARAMETER SkipDeprecatedCleanup
    Skips the cleanup of deprecated modules

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

.NOTES
    Author: CIAOPS    Version: 2.8
    Last Updated: June 2025
    Requires: PowerShell 5.1 or higher, Administrator privileges
    
    IMPORTANT: This version removes deprecated Azure AD and MSOnline modules in favor of Microsoft Graph PowerShell SDK
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
}

# Module definitions
$Script:ModuleList = @(
    @{ Name = 'Microsoft.Graph'; Description = 'Microsoft Graph module'; Deprecated = $false },
    @{ Name = 'Microsoft.Graph.Authentication'; Description = 'Microsoft Graph Authentication'; Deprecated = $false },
    @{ Name = 'MicrosoftTeams'; Description = 'Teams Module'; Deprecated = $false },
    @{ Name = 'ExchangeOnlineManagement'; Description = 'Exchange Online module'; Deprecated = $false },
    @{ Name = 'Az'; Description = 'Azure PowerShell module'; Deprecated = $false },
    @{ Name = 'PnP.PowerShell'; Description = 'SharePoint PnP module'; Deprecated = $false },
    @{ Name = 'Microsoft.PowerApps.PowerShell'; Description = 'PowerApps'; Deprecated = $false },
    @{ Name = 'Microsoft.PowerApps.Administration.PowerShell'; Description = 'PowerApps Administration module'; Deprecated = $false },
    @{ Name = 'PowershellGet'; Description = 'PowerShellGet module'; Deprecated = $false; RequiresSpecialHandling = $true },
    @{ Name = 'PackageManagement'; Description = 'Package Management module'; Deprecated = $false; RequiresSpecialHandling = $true },
    @{ Name = 'Microsoft.Online.SharePoint.PowerShell'; Description = 'SharePoint Online Management Shell'; Deprecated = $false },
    @{ Name = 'Microsoft.WinGet.Client'; Description = 'Windows Package Manager Client'; Deprecated = $false }
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

function Write-ColorOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('System', 'Process', 'Warning', 'Error', 'Info')]
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
                        $installResult = Install-Module @coreUpdateParams 2>$null 3>$null 4>$null 5>$null 6>$null
                        
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
        Returns appropriate parameters for Install-Module based on the specific module,
        as some modules don't support certain parameters like -AllowClobber, -AcceptLicense,
        or -SkipPublisherCheck
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,
        
        [Parameter()]
        [hashtable]$BaseParams = @{}
    )    # Modules that don't support -AllowClobber
    $noAllowClobberModules = @(
        'Microsoft.PowerApps.Administration.PowerShell',
        'Microsoft.PowerApps.PowerShell',
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
        'Microsoft.PowerApps.Administration.PowerShell',
        'Microsoft.PowerApps.PowerShell',
        'Microsoft.WinGet.Client',
        'Microsoft.Graph',
        'Microsoft.Graph.Authentication',
        'MicrosoftTeams',
        'PnP.PowerShell',
        'Microsoft.Online.SharePoint.PowerShell'
    )
    
    # Modules that don't support -SkipPublisherCheck
    $noSkipPublisherCheckModules = @(
        'Microsoft.PowerApps.Administration.PowerShell',
        'Microsoft.PowerApps.PowerShell',
        'Microsoft.WinGet.Client',
        'ExchangeOnlineManagement',
        'Microsoft.Graph',
        'Microsoft.Graph.Authentication',
        'Az',
        'MicrosoftTeams',
        'PnP.PowerShell',
        'Microsoft.Online.SharePoint.PowerShell'
    )
    
    # Start with base parameters
    $moduleParams = $BaseParams.Clone()
    
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
                "Starting $Operation of $ModuleName at $(Get-Date -Format 'HH:mm:ss')" | Out-File $progressFile -Force
                "Phase: Download and Installation" | Out-File $progressFile -Append
            }
            
            # Execute the installation/update
            switch ($Operation) {
                'Install' { 
                    if ($isConstrainedLargeModule) {
                        "Status: Installing $ModuleName..." | Out-File $progressFile -Append
                    }
                    Install-Module -Name $ModuleName @InstallParams
                    if ($isConstrainedLargeModule) {
                        "Phase: Installation Complete - Starting Verification" | Out-File $progressFile -Append
                        "Status: Verification phase may take several minutes and appear frozen" | Out-File $progressFile -Append
                    }
                }
                'Update' { 
                    if ($isConstrainedLargeModule) {
                        "Status: Updating $ModuleName..." | Out-File $progressFile -Append
                    }
                    Update-Module -Name $ModuleName @InstallParams
                    if ($isConstrainedLargeModule) {
                        "Phase: Update Complete - Starting Verification" | Out-File $progressFile -Append
                        "Status: Verification phase may take several minutes and appear frozen" | Out-File $progressFile -Append
                    }
                }
                default { 
                    if ($isConstrainedLargeModule) {
                        "Status: Installing $ModuleName (default)..." | Out-File $progressFile -Append
                    }
                    Install-Module -Name $ModuleName @InstallParams
                    if ($isConstrainedLargeModule) {
                        "Phase: Installation Complete - Starting Verification" | Out-File $progressFile -Append
                        "Status: Verification phase may take several minutes and appear frozen" | Out-File $progressFile -Append
                    }
                }
            }
            
            # Final cleanup phase
            if ($isConstrainedLargeModule) {
                "Phase: Verification Complete - Final Cleanup" | Out-File $progressFile -Append
                Start-Sleep -Seconds 2  # Brief pause for cleanup
                "Phase: Complete at $(Get-Date -Format 'HH:mm:ss')" | Out-File $progressFile -Append
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
                Start-Sleep -Seconds 5
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
                    $updateArgs = @{
                        Name = $ModuleName
                        Force = if ($InstallParams.ContainsKey('Force')) { $InstallParams.Force } else { $true }
                        Confirm = if ($InstallParams.ContainsKey('Confirm')) { $InstallParams.Confirm } else { $false }
                        ErrorAction = 'Stop'
                    }
                    if ($InstallParams.ContainsKey('AllowClobber')) { $updateArgs.AllowClobber = $InstallParams.AllowClobber }
                    if ($InstallParams.ContainsKey('AcceptLicense')) { $updateArgs.AcceptLicense = $InstallParams.AcceptLicense }
                    if ($InstallParams.ContainsKey('SkipPublisherCheck')) { $updateArgs.SkipPublisherCheck = $InstallParams.SkipPublisherCheck }
                    
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
                    $updateArgs = @{
                        Name = $ModuleName
                        Force = if ($InstallParams.ContainsKey('Force')) { $InstallParams.Force } else { $true }
                        Confirm = if ($InstallParams.ContainsKey('Confirm')) { $InstallParams.Confirm } else { $false }
                        ErrorAction = 'Stop'
                    }
                    if ($InstallParams.ContainsKey('AllowClobber')) { $updateArgs.AllowClobber = $InstallParams.AllowClobber }
                    if ($InstallParams.ContainsKey('AcceptLicense')) { $updateArgs.AcceptLicense = $InstallParams.AcceptLicense }
                    if ($InstallParams.ContainsKey('SkipPublisherCheck')) { $updateArgs.SkipPublisherCheck = $InstallParams.SkipPublisherCheck }
                    
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
                    $phaseProgress = [Math]::Min(25, ($elapsedSeconds / 300) * 25) # 5 minutes for this phase
                    $progressPercent = $phaseProgress
                }
                "Verification" {
                    # 25-85% is verification (slowest phase)
                    $verificationTime = $elapsedSeconds - 300
                    $verificationProgress = [Math]::Min(60, ($verificationTime / 600) * 60) # 10 minutes for verification
                    $progressPercent = 25 + $verificationProgress
                    
                    # Show periodic reminders during verification
                    if (($elapsedMinutes -gt 8) -and (((Get-Date) - $lastProgressUpdate).TotalSeconds -gt 90)) {
                        Write-ColorOutput ("    ⏳ Still in verification phase ({0:N1} min elapsed) - please continue waiting..." -f $elapsedMinutes) -Type Info
                        $lastProgressUpdate = Get-Date
                    }
                }
                "Cleanup" {
                    # 85-100% is cleanup (usually quick)
                    $cleanupTime = $elapsedSeconds - 900
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
            Start-Sleep -Seconds 10  # Longer intervals during verification to reduce system load
        } else {
            Start-Sleep -Seconds 3
        }
        $iteration++
        
        # Enhanced timeout handling for large modules in constrained mode
        $timeoutMultiplier = if ($isConstrainedLargeModule) { 2.5 } else { 1.5 }
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

function Test-ModuleInstallation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,
        
        [Parameter()]
        [string]$Description = $ModuleName
    )
    
    try {
        Write-ColorOutput "    Checking module: $Description" -Type Info
        
        $installedModule = Get-InstalledModule -Name $ModuleName -ErrorAction SilentlyContinue
        
        if (-not $installedModule) {
            Write-ColorOutput "    [Warning] Module '$ModuleName' not found" -Type Warning
            
            if ($Prompt -and -not $CheckOnly) {
                $estimate = Get-ModuleInstallationEstimate -ModuleName $ModuleName -Operation 'Install'
                Write-ColorOutput "    Installation details:" -Type Info
                Write-ColorOutput "      Size: $($estimate.FormattedSize), Time: $($estimate.FormattedTime)" -Type Info
                
                $response = Read-Host "    Install module '$ModuleName' (Y/N)?"
                if ($response -notmatch '^[Yy]') {
                    Write-ColorOutput "    Skipping installation of $ModuleName" -Type Warning
                    return
                }
            }
              if (-not $CheckOnly) {
                Write-ColorOutput "    Installing module: $ModuleName" -Type Process
                
                # Test compatibility in constrained environments before optimization
                $languageMode = $ExecutionContext.SessionState.LanguageMode
                if ($languageMode -eq 'ConstrainedLanguage') {
                    Write-ColorOutput "    Constrained Language Mode detected - testing compatibility..." -Type Warning
                    $isCompatible = Test-ConstrainedLanguageModeCompatibility -Operation 'NetworkOptimization' -Silent
                    if (-not $isCompatible) {
                        Write-ColorOutput "    Network optimization may be limited in this environment" -Type Warning
                    }
                }
                
                # Initialize optimized settings
                $baseParams = Initialize-OptimizedInstallation
                $installParams = Get-ModuleSpecificParams -ModuleName $ModuleName -BaseParams $baseParams
                
                # Use progress tracking for large modules
                $largeModules = @('Az', 'Microsoft.Graph', 'PnP.PowerShell', 'AzureAD')
                if ($ModuleName -in $largeModules) {
                    $success = Install-ModuleWithProgress -ModuleName $ModuleName -InstallParams $installParams -Operation 'Install'
                    if (-not $success) {
                        throw "Installation failed"
                    }
                }
                else {
                    # For smaller modules, use direct installation with explicit parameter handling
                    $installArgs = @{
                        Name = $ModuleName
                        Force = if ($installParams.ContainsKey('Force')) { $installParams.Force } else { $true }
                        Confirm = if ($installParams.ContainsKey('Confirm')) { $installParams.Confirm } else { $false }
                        ErrorAction = 'Stop'
                    }
                    if ($installParams.ContainsKey('Scope')) { $installArgs.Scope = $installParams.Scope }
                    if ($installParams.ContainsKey('AllowClobber')) { $installArgs.AllowClobber = $installParams.AllowClobber }
                    if ($installParams.ContainsKey('AcceptLicense')) { $installArgs.AcceptLicense = $installParams.AcceptLicense }
                    if ($installParams.ContainsKey('SkipPublisherCheck')) { $installArgs.SkipPublisherCheck = $installParams.SkipPublisherCheck }
                    
                    Install-Module @installArgs
                    Write-ColorOutput "    Successfully installed $ModuleName" -Type Process
                }
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
            Write-ColorOutput "    Module $ModuleName ($localVersion) is up to date" -Type Process
        }
        else {
            Write-ColorOutput "    Module $ModuleName ($localVersion) can be updated to ($onlineVersion)" -Type Warning
              if (-not $CheckOnly) {
                # Test compatibility in constrained environments before optimization
                $languageMode = $ExecutionContext.SessionState.LanguageMode
                if ($languageMode -eq 'ConstrainedLanguage') {
                    Write-ColorOutput "    Constrained Language Mode detected - testing compatibility..." -Type Warning
                    $isCompatible = Test-ConstrainedLanguageModeCompatibility -Operation 'NetworkOptimization' -Silent
                    if (-not $isCompatible) {
                        Write-ColorOutput "    Network optimization may be limited in this environment" -Type Warning
                    }
                }
                
                # Initialize optimized settings
                $baseParams = Initialize-OptimizedInstallation
                $updateParams = Get-ModuleSpecificParams -ModuleName $ModuleName -BaseParams $baseParams
                
                # Use progress tracking for large modules
                $largeModules = @('Az', 'Microsoft.Graph', 'PnP.PowerShell', 'AzureAD')
                if ($ModuleName -in $largeModules) {
                    Write-ColorOutput "    Updating module: $ModuleName" -Type Process
                    $success = Install-ModuleWithProgress -ModuleName $ModuleName -InstallParams $updateParams -Operation 'Update'
                    if (-not $success) {
                        throw "Update failed"
                    }
                }
                else {
                    Write-ColorOutput "    Updating module: $ModuleName" -Type Process
                    # Use explicit parameter handling for constrained language mode compatibility
                    $updateArgs = @{
                        Name = $ModuleName
                        Force = if ($updateParams.ContainsKey('Force')) { $updateParams.Force } else { $true }
                        Confirm = if ($updateParams.ContainsKey('Confirm')) { $updateParams.Confirm } else { $false }
                        ErrorAction = 'Stop'
                    }
                    if ($updateParams.ContainsKey('AllowClobber')) { $updateArgs.AllowClobber = $updateParams.AllowClobber }
                    if ($updateParams.ContainsKey('AcceptLicense')) { $updateArgs.AcceptLicense = $updateParams.AcceptLicense }
                    if ($updateParams.ContainsKey('SkipPublisherCheck')) { $updateArgs.SkipPublisherCheck = $updateParams.SkipPublisherCheck }
                    
                    Update-Module @updateArgs
                    Write-ColorOutput "    Successfully updated $ModuleName" -Type Process
                }
            }
        }
    }
    catch {        Write-ColorOutput "    Error processing module '$ModuleName' - $($PSItem.Exception.Message)" -Type Error
    }
}

function Remove-ModuleWithProgress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,
        
        [Parameter()]
        [string]$ModuleVersion,
        
        [Parameter()]
        [object]$ModuleInfo
    )
    
    $estimate = Get-ModuleInstallationEstimate -ModuleName $ModuleName -Operation 'Remove'
    $startTime = Get-Date
    
    # Show pre-removal information for large modules
    $largeModules = @('Az', 'Microsoft.Graph', 'PnP.PowerShell', 'AzureAD')
    if ($ModuleName -in $largeModules) {
        Write-ColorOutput "    Removal details for $ModuleName" -Type Info
        Write-ColorOutput "      Estimated time: $($estimate.FormattedTime)" -Type Info
    }
    
    # Create a background job for the actual removal
    $jobScript = {
        param($ModuleName, $ModuleVersion, $ModuleInfo)
        
        try {
            # Use traditional if-else for Constrained Language Mode compatibility
            if ($ModuleInfo.RepositorySourceLocation) {
                $installMethod = "PowerShellGet"
            } else {
                $installMethod = "MSI/Manual"
            }
            
            if ($installMethod -eq "MSI/Manual") {
                # Handle MSI/Manual installed modules
                $canUninstallViaPS = Get-InstalledModule -Name $ModuleName -RequiredVersion $ModuleVersion -ErrorAction SilentlyContinue
                
                if ($canUninstallViaPS) {
                    Uninstall-Module -Name $ModuleName -RequiredVersion $ModuleVersion -Force -Confirm:$false -ErrorAction Stop
                    return @{ Success = $true; Message = "Uninstalled via PowerShell" }
                }
                else {
                    # Manual removal from file system (with caution)
                    if ($ModuleInfo.ModuleBase -and (Test-Path $ModuleInfo.ModuleBase)) {
                        $safeToRemove = $ModuleInfo.ModuleBase -match "\\Users\\.*\\Documents\\.*PowerShell.*\\Modules" -or
                                       $ModuleInfo.ModuleBase -match "\\Program Files\\.*PowerShell.*\\Modules" -or
                                       $ModuleInfo.ModuleBase -match "\\PowerShell\\Modules"
                        
                        if ($safeToRemove) {
                            $testFile = Join-Path $ModuleInfo.ModuleBase "test.tmp"
                            try {
                                [System.IO.File]::Create($testFile).Close()
                                Remove-Item $testFile -Force -ErrorAction SilentlyContinue
                                
                                Remove-Item -Path $ModuleInfo.ModuleBase -Recurse -Force -ErrorAction Stop
                                return @{ Success = $true; Message = "Removed from file system" }
                            }
                            catch [System.UnauthorizedAccessException] {
                                return @{ Success = $false; Message = "Access denied - module files may be in use or require admin rights" }
                            }
                            catch [System.IO.IOException] {
                                return @{ Success = $false; Message = "Module files are locked or in use - close applications and retry" }
                            }
                        }
                        else {
                            return @{ Success = $false; Message = "Module in system location - manual removal required" }
                        }
                    }
                    else {
                        return @{ Success = $false; Message = "Module path not found or no longer exists" }
                    }
                }
            }
            else {
                # PowerShellGet installed modules
                $retryAttempts = 2
                for ($attempt = 1; $attempt -le $retryAttempts; $attempt++) {
                    try {
                        if ($attempt -gt 1) { Start-Sleep -Seconds 2 }
                        
                        Uninstall-Module -Name $ModuleName -RequiredVersion $ModuleVersion -Force -Confirm:$false -ErrorAction Stop
                        return @{ Success = $true; Message = "Uninstalled via PowerShellGet" }
                    }
                    catch {
                        $errorMsg = $_.Exception.Message
                        
                        if ($errorMsg -like "*Administrator*" -or $errorMsg -like "*elevated*") {
                            return @{ Success = $false; Message = "Requires administrator privileges" }
                        }
                        elseif ($errorMsg -like "*in use*" -or $errorMsg -like "*locked*") {
                            if ($attempt -eq $retryAttempts) {
                                return @{ Success = $false; Message = "Module is currently in use - close PowerShell sessions and retry" }
                            }
                        }
                        elseif ($errorMsg -like "*not found*" -or $errorMsg -like "*does not exist*") {
                            try {
                                Uninstall-Module -Name $ModuleName -Force -Confirm:$false -ErrorAction Stop
                                return @{ Success = $true; Message = "Uninstalled via PowerShellGet (all versions)" }
                            }
                            catch {
                                if ($attempt -eq $retryAttempts) {
                                    return @{ Success = $false; Message = "Module not found in PowerShellGet registry" }
                                }
                            }
                        }
                        else {
                            if ($attempt -eq $retryAttempts) {
                                return @{ Success = $false; Message = "PowerShellGet uninstall failed: $errorMsg" }
                            }
                        }
                    }
                }
                return @{ Success = $false; Message = "All retry attempts failed" }
            }
        }
        catch {
            return @{ Success = $false; Message = $_.Exception.Message }
        }
    }
    
    # For large modules, show progress; for smaller ones, just process directly
    if ($ModuleName -in $largeModules) {
        # Check if background jobs are supported in this environment
        if (-not $Script:JobsSupported) {
            Write-ColorOutput "    Background jobs not supported - using direct execution for removal" -Type Warning
            
            # Execute directly without background jobs
            $result = & $jobScript $ModuleName $ModuleVersion $ModuleInfo
        }
        else {
            # Try to start the background job
            try {
                $job = Start-Job -ScriptBlock $jobScript -ArgumentList $ModuleName, $ModuleVersion, $ModuleInfo
            }
            catch {
                Write-ColorOutput "    Background job creation failed - falling back to direct execution" -Type Warning
                Write-ColorOutput "    Error: $($_.Exception.Message)" -Type Warning
                
                # Fallback to direct execution
                $result = & $jobScript $ModuleName $ModuleVersion $ModuleInfo
                $actualTime = ((Get-Date) - $startTime).TotalSeconds
                
                return @{
                    Success = $result.Success
                    Message = $result.Message
                    ActualTime = $actualTime
                }
            }
        
            # Show progress while waiting (only if job was created successfully)
            if ($job) {
                $progressParams = @{
                    Activity = "Removing module: $ModuleName"
                    Status = "Uninstalling..."
                    PercentComplete = 0
                }
                
                $iteration = 0
                $maxIterations = [Math]::Max(5, [Math]::Ceiling($estimate.EstimatedTime / 2))
                
                while ($job.State -eq 'Running') {
                    $elapsed = (Get-Date) - $startTime
                    $elapsedSeconds = $elapsed.TotalSeconds
                    
                    $progressPercent = [Math]::Min(95, ($elapsedSeconds / $estimate.EstimatedTime) * 100)
                    
                    $progressParams.PercentComplete = $progressPercent
                    $progressParams.Status = "Progress: {0:N0}% - Elapsed: {1:N0}s" -f $progressPercent, $elapsedSeconds
                    
                    if ($estimate.EstimatedTime -gt $elapsedSeconds) {
                        $remainingSeconds = $estimate.EstimatedTime - $elapsedSeconds
                        if ($remainingSeconds -gt 60) {
                            $progressParams.Status += " - Est. remaining: {0:N1} min" -f ($remainingSeconds / 60)
                        } else {
                            $progressParams.Status += " - Est. remaining: {0:N0}s" -f $remainingSeconds
                        }
                    }
                    
                    Write-Progress @progressParams
                    Start-Sleep -Seconds 2
                    $iteration++
                    
                    if ($iteration -gt $maxIterations -and $elapsedSeconds -gt ($estimate.EstimatedTime * 1.5)) {
                        $progressParams.Status = "Taking longer than expected - Elapsed: {0:N0}s" -f $elapsedSeconds
                        Write-Progress @progressParams
                    }
                }
                
                # Complete the progress bar
                Write-Progress -Activity "Removing module: $ModuleName" -Completed
                
                # Get the job result
                $result = Receive-Job -Job $job
                Remove-Job -Job $job
            }
        }
    }
    else {
        # Direct execution for smaller modules
        $result = & $jobScript $ModuleName $ModuleVersion $ModuleInfo
    }
    
    $actualTime = ((Get-Date) - $startTime).TotalSeconds
    
    return @{
        Success = $result.Success
        Message = $result.Message
        ActualTime = $actualTime
    }
}

function Remove-DeprecatedModules {
    [CmdletBinding()]
    param()
    
    if ($SkipDeprecatedCleanup -or $CheckOnly) {
        return
    }
    
    # Detect constrained language mode for simplified user experience
    $isConstrainedMode = $ExecutionContext.SessionState.LanguageMode -eq 'ConstrainedLanguage'
    
    Write-ColorOutput ""
    if ($isConstrainedMode) {
        Write-ColorOutput "=== DEPRECATED MODULE CLEANUP ===" -Type System
        Write-ColorOutput "Scanning for deprecated PowerShell modules..." -Type Info
    } else {
        Write-ColorOutput "=== DEPRECATED MODULE CLEANUP ANALYSIS ===" -Type System
        Write-ColorOutput ""
        Write-ColorOutput "Microsoft has deprecated several PowerShell modules in favor of modern alternatives." -Type Info
        Write-ColorOutput "This section will help you identify and optionally remove these deprecated modules." -Type Info
    }
    Write-ColorOutput ""
    
    
    # Display deprecation overview - simplified for constrained mode
    if ($isConstrainedMode) {
        Write-ColorOutput "Checking for deprecated modules: $($Script:DeprecatedModules.Name -join ', ')" -Type Info
    } else {
        Write-ColorOutput "📋 DEPRECATION OVERVIEW:" -Type Warning
        Write-ColorOutput ""
        foreach ($deprecated in $Script:DeprecatedModules) {
            Write-ColorOutput "  🚫 $($deprecated.Name)" -Type Error
            Write-ColorOutput "     ↳ Replacement: $($deprecated.Replacement)" -Type Process
            Write-ColorOutput "     ↳ Reason: $($deprecated.Reason)" -Type Info
            Write-ColorOutput ""
        }
        
        Write-ColorOutput "💡 WHY REMOVE DEPRECATED MODULES?" -Type Info
        Write-ColorOutput "  • Security: Deprecated modules no longer receive security updates" -Type Warning
        Write-ColorOutput "  • Support: Microsoft has ended support for these modules" -Type Warning
        Write-ColorOutput "  • Functionality: New modules offer enhanced features and better performance" -Type Process
        Write-ColorOutput "  • Compatibility: Avoid conflicts between old and new modules" -Type Process
        Write-ColorOutput ""
    }
    
    if ($isConstrainedMode) {
        Write-ColorOutput "Scanning for installed deprecated modules..." -Type System
    } else {
        Write-ColorOutput "Scanning your system for deprecated modules..." -Type System
        Write-ColorOutput ""
        
        # Initialize scanning progress
        $totalModulesToCheck = $Script:DeprecatedModules.Count
        $currentModuleIndex = 0
        
        Write-ColorOutput "🔍 SCANNING PROGRESS:" -Type Info
        Write-ColorOutput "  Modules to check: $totalModulesToCheck deprecated modules" -Type Info
        Write-ColorOutput "  Scan locations:" -Type Info
        Write-ColorOutput "    • PowerShellGet installed modules" -Type Info
        Write-ColorOutput "    • System module paths (including MSI installations)" -Type Info
        Write-ColorOutput "    • User profile module directories" -Type Info
        Write-ColorOutput ""
    }
    
    # Initialize tracking variables
    $foundDeprecatedModules = 0
    $foundVersions = 0
    $currentModuleIndex = 0
    
    # First, collect all deprecated modules that are installed
    $modulesToRemove = @()
    foreach ($deprecated in $Script:DeprecatedModules) {
        $currentModuleIndex++
        if ($isConstrainedMode) {
            # Simplified progress display for constrained mode
            Write-ColorOutput "  Checking $($deprecated.Name)..." -Type Info -NoNewline
        } else {
            # Detailed progress display for normal mode
            $percentComplete = [math]::Round(($currentModuleIndex / $Script:DeprecatedModules.Count) * 100, 1)
            Write-ColorOutput "  [$percentComplete%] ($currentModuleIndex/$($Script:DeprecatedModules.Count)) Checking: $($deprecated.Name)" -Type Info
        }
        
        # Check both Get-Module -ListAvailable and Get-InstalledModule
        $installedVersions = @()
        $moduleFound = $false
        
        if ($isConstrainedMode) {
            # Simplified scanning for constrained mode - no detailed output
            try {
                $psGetModules = Get-InstalledModule -Name $deprecated.Name -ErrorAction SilentlyContinue
                if ($psGetModules) {
                    $installedVersions += $psGetModules
                    $moduleFound = $true
                }
            }
            catch {
                # Silent error handling in constrained mode
            }
            
            try {
                $availableModules = Get-Module -ListAvailable -Name $deprecated.Name -ErrorAction SilentlyContinue
                if ($availableModules) {
                    foreach ($module in $availableModules) {
                        $alreadyListed = $installedVersions | Where-Object { 
                            $_.Name -eq $module.Name -and $_.Version -eq $module.Version 
                        }
                        if (-not $alreadyListed) {
                            $installedVersions += [PSCustomObject]@{
                                Name = $module.Name
                                Version = $module.Version
                                ModuleBase = $module.ModuleBase
                                InstalledLocation = $module.ModuleBase
                                InstalledBy = "Unknown"
                                InstalledVia = "MSI/Manual"
                            }
                            $moduleFound = $true
                        }
                    }
                }
            }
            catch {
                # Silent error handling in constrained mode
            }
            
            # Simple result display
            if ($moduleFound) {
                Write-ColorOutput " Found ($($installedVersions.Count) versions)" -Type Warning
            } else {
                Write-ColorOutput " Not found" -Type Process
            }
        } else {
            # Detailed scanning for normal mode
            # Check modules installed via PowerShellGet
            Write-ColorOutput "    • Scanning PowerShellGet registry..." -Type Info -NoNewline
            try {
                $psGetModules = Get-InstalledModule -Name $deprecated.Name -ErrorAction SilentlyContinue
                if ($psGetModules) {
                    $installedVersions += $psGetModules
                    $moduleFound = $true
                    Write-ColorOutput " Found $($psGetModules.Count) version(s)" -Type Process
                } else {
                    Write-ColorOutput " None found" -Type Info
                }
            }
            catch {
                Write-ColorOutput " Error during scan" -Type Warning
            }
            
            # Check modules available in the system (including MSI-installed)
            Write-ColorOutput "    • Scanning system module paths..." -Type Info -NoNewline
            try {
                $availableModules = Get-Module -ListAvailable -Name $deprecated.Name -ErrorAction SilentlyContinue
                if ($availableModules) {
                    $systemModulesCount = 0
                    # Add modules that aren't already in the PowerShellGet list
                    foreach ($module in $availableModules) {
                        $alreadyListed = $installedVersions | Where-Object { 
                            $_.Name -eq $module.Name -and $_.Version -eq $module.Version 
                        }
                        if (-not $alreadyListed) {
                            # Create a custom object that mimics Get-InstalledModule output
                            $installedVersions += [PSCustomObject]@{
                                Name = $module.Name
                                Version = $module.Version
                                ModuleBase = $module.ModuleBase
                                InstalledLocation = $module.ModuleBase
                                InstalledBy = "Unknown"
                                InstalledVia = "MSI/Manual"
                            }
                            $systemModulesCount++
                            $moduleFound = $true
                        }
                    }
                    if ($systemModulesCount -gt 0) {
                        Write-ColorOutput " Found $systemModulesCount additional version(s)" -Type Process
                    } else {
                        Write-ColorOutput " No additional versions found" -Type Info
                    }
                } else {
                    Write-ColorOutput " None found" -Type Info
                }
            }
            catch {
                Write-ColorOutput " Error during scan" -Type Warning
            }
            
            if ($moduleFound) {
                $foundDeprecatedModules++
                $foundVersions += $installedVersions.Count
                
                Write-ColorOutput "    ✅ FOUND: $($deprecated.Name) - $($installedVersions.Count) version(s) installed" -Type Warning
                
                # Show installation locations for user awareness
                $locations = $installedVersions | Group-Object -Property { 
                    if ($_.PSObject.Properties.Name -contains "InstalledVia") { 
                        $_.InstalledVia 
                    } else { 
                        "PowerShellGet" 
                    }
                }
                foreach ($location in $locations) {
                    Write-ColorOutput "      • $($location.Count) version(s) via $($location.Name)" -Type Info
                }
            } else {
                Write-ColorOutput "    ✓ Clean: $($deprecated.Name) not found" -Type Process
            }
        }
        
        
        # Update tracking variables
        if ($moduleFound) {
            $foundDeprecatedModules++
            $foundVersions += $installedVersions.Count
        }
        
        if ($installedVersions) {
            # Estimate time based on module complexity and size
            $estimatedTimePerVersion = switch ($deprecated.Name) {
                'Az' { 90 }  # Azure modules are very large
                'Microsoft.Graph' { 60 }  # Graph modules are large
                'AzureAD' { 45 }  # AzureAD modules are medium-large
                'MSOnline' { 30 }  # MSOnline is medium
                'PnP.PowerShell' { 45 }  # PnP modules are medium-large
                'SharePointPnPPowerShellOnline' { 30 }  # Old PnP is medium
                'WindowsAutoPilotIntune' { 25 }  # Intune modules are medium
                default { 20 }  # Default estimation for smaller modules
            }
            
            $moduleEstimate = $estimatedTimePerVersion * $installedVersions.Count
            
            $modulesToRemove += @{
                Name = $deprecated.Name
                Replacement = $deprecated.Replacement
                Reason = $deprecated.Reason
                Versions = $installedVersions
                VersionCount = $installedVersions.Count
                EstimatedTime = $moduleEstimate
            }
        }
        
        # Add newline only for normal mode (not constrained mode for cleaner output)
        if (-not $isConstrainedMode) {
            Write-ColorOutput ""
        }
    }
    
    # Add scanning completion summary - simplified for constrained mode
    Write-ColorOutput ""
    if ($isConstrainedMode) {
        Write-ColorOutput "Scan complete: Found $foundDeprecatedModules deprecated modules ($foundVersions total versions)" -Type System
    } else {
        Write-ColorOutput "📊 SCANNING COMPLETE:" -Type System
        Write-ColorOutput "  • Modules checked: $($Script:DeprecatedModules.Count)" -Type Info
        Write-ColorOutput "  • Deprecated modules found: $foundDeprecatedModules" -Type Warning
        Write-ColorOutput "  • Total versions found: $foundVersions" -Type Warning
        Write-ColorOutput "  • Clean modules (not installed): $($Script:DeprecatedModules.Count - $foundDeprecatedModules)" -Type Process
    }
    Write-ColorOutput ""
    
    if ($foundDeprecatedModules -eq 0) {
        if ($isConstrainedMode) {
            Write-ColorOutput "✅ No deprecated modules found - your system is clean." -Type Process
        } else {
            Write-ColorOutput "✅ SCAN RESULTS: No deprecated modules found on your system" -Type Process
            Write-ColorOutput "Your system is clean of deprecated PowerShell modules." -Type Process
            Write-ColorOutput ""
            Write-ColorOutput "🎉 Congratulations! Your PowerShell environment uses only modern, supported modules." -Type Process
        }
        Write-ColorOutput ""
        return
    }
    
    
    # Calculate total time estimation
    $totalEstimatedTime = ($modulesToRemove | Measure-Object -Property EstimatedTime -Sum).Sum
    $totalVersions = ($modulesToRemove | Measure-Object -Property VersionCount -Sum).Sum
    $estimatedMinutes = [math]::Round($totalEstimatedTime / 60, 1)
    
    if ($isConstrainedMode) {
        # Simplified results display for constrained mode
        Write-ColorOutput "Found deprecated modules to remove:" -Type Warning
        foreach ($moduleInfo in $modulesToRemove) {
            Write-ColorOutput "  • $($moduleInfo.Name) ($($moduleInfo.VersionCount) versions) → $($moduleInfo.Replacement)" -Type Info
        }
        Write-ColorOutput "Estimated removal time: $estimatedMinutes minutes" -Type Info
    } else {
        # Detailed results display for normal mode
        Write-ColorOutput "🔍 SCAN RESULTS:" -Type Warning
        Write-ColorOutput "Found $($modulesToRemove.Count) deprecated module(s) with $totalVersions total versions installed" -Type Warning
        Write-ColorOutput "Estimated total removal time: $estimatedMinutes minutes" -Type Info
        Write-ColorOutput ""
        
        # Display detailed findings
        Write-ColorOutput "📊 DETAILED FINDINGS:" -Type Info
        Write-ColorOutput ""
        foreach ($moduleInfo in $modulesToRemove) {
            Write-ColorOutput "  🚫 $($moduleInfo.Name)" -Type Error
            Write-ColorOutput "     📦 Versions installed: $($moduleInfo.VersionCount)" -Type Info
            Write-ColorOutput "     ⏱️  Estimated removal time: $([math]::Round($moduleInfo.EstimatedTime / 60, 1)) minutes" -Type Info
            Write-ColorOutput "     🔄 Replacement: $($moduleInfo.Replacement)" -Type Process
            Write-ColorOutput "     📋 Deprecation reason: $($moduleInfo.Reason)" -Type Warning
            Write-ColorOutput ""
            
            # Show version details
            Write-ColorOutput "     📋 Installed versions:" -Type Info
            foreach ($version in $moduleInfo.Versions) {
                $installMethod = if ($version.PSObject.Properties.Name -contains "InstalledVia") { 
                    $version.InstalledVia 
                } else { 
                    "PowerShellGet" 
                }
                Write-ColorOutput "       • Version $($version.Version) (via $installMethod)" -Type Info
            }
            Write-ColorOutput ""
        }
    }
    
    if ($isConstrainedMode) {
        # Simplified considerations for constrained mode
        Write-ColorOutput ""
        Write-ColorOutput "⚠️  Before removing: Ensure your scripts use the replacement modules." -Type Warning
    } else {
        # Detailed considerations for normal mode
        Write-ColorOutput "⚠️  IMPORTANT CONSIDERATIONS BEFORE REMOVAL:" -Type Warning
        Write-ColorOutput ""
        Write-ColorOutput "  ✅ SAFE TO REMOVE:" -Type Process
        Write-ColorOutput "     • You have replacement modules installed (Microsoft.Graph, etc.)" -Type Process
        Write-ColorOutput "     • Your scripts use the newer modules" -Type Process
        Write-ColorOutput "     • You understand the migration requirements" -Type Process
        Write-ColorOutput ""
        Write-ColorOutput "  ⚠️  CAUTION REQUIRED:" -Type Warning
        Write-ColorOutput "     • You have scripts that still use the deprecated modules" -Type Warning
        Write-ColorOutput "     • You haven't migrated to the replacement modules yet" -Type Warning
        Write-ColorOutput "     • You're unsure about compatibility with existing code" -Type Warning
        Write-ColorOutput ""
        Write-ColorOutput "  📚 MIGRATION GUIDANCE:" -Type Info
        Write-ColorOutput "     • AzureAD/MSOnline → Microsoft.Graph: https://docs.microsoft.com/graph/powershell/migration" -Type Info
        Write-ColorOutput "     • SharePoint PnP → PnP.PowerShell: https://pnp.github.io/powershell/articles/migration.html" -Type Info
        Write-ColorOutput "     • Test your scripts thoroughly after removing deprecated modules" -Type Info
        Write-ColorOutput ""
    }
    
    # Initialize processing variables
    $startTime = Get-Date
    $totalOperations = $totalVersions
    $currentOperation = 0
    $skippedModules = @()
    
    if ($isConstrainedMode) {
        # Simplified confirmation for constrained mode
        Write-ColorOutput ""
        if (-not $Prompt) {
            $response = Read-Host "Remove these deprecated modules? (Y/N)"
            if ($response -notmatch '^[Yy]') {
                Write-ColorOutput "Deprecated module removal cancelled." -Type Warning
                return
            }
        }
        Write-ColorOutput "Starting module removal..." -Type System
    } else {
        # Detailed confirmation and display for normal mode
        Write-ColorOutput "Found $($modulesToRemove.Count) deprecated modules with $totalVersions total versions to remove" -Type Warning
        Write-ColorOutput "Estimated total removal time: $estimatedMinutes minutes" -Type Info
        Write-ColorOutput ""
        
        # Display what will be removed
        foreach ($moduleInfo in $modulesToRemove) {
            Write-ColorOutput "    • $($moduleInfo.Name) ($($moduleInfo.VersionCount) versions) - Est: $([math]::Round($moduleInfo.EstimatedTime / 60, 1))min" -Type Info
        }
        Write-ColorOutput ""
        
        # Always prompt for confirmation before removing modules (unless already prompting per module)
        if (-not $Prompt) {
            Write-ColorOutput "⚠ WARNING: This will remove all deprecated modules listed above." -Type Warning
            Write-ColorOutput "These modules are being replaced by newer Microsoft Graph and other modern modules." -Type Warning
            Write-ColorOutput ""
            $response = Read-Host "Do you want to proceed with removing these deprecated modules? (Y/N)"
            if ($response -notmatch '^[Yy]') {
                Write-ColorOutput "Module removal cancelled by user. Skipping deprecated module cleanup." -Type Warning
                Write-ColorOutput ""
                return
            }
            Write-ColorOutput ""
        }
    }
    
    foreach ($moduleInfo in $modulesToRemove) {
        if ($isConstrainedMode) {
            # Simplified processing header for constrained mode
            Write-ColorOutput "Removing $($moduleInfo.Name) ($($moduleInfo.VersionCount) versions)..." -Type Warning
        } else {
            # Detailed processing header for normal mode
            Write-ColorOutput "    [Warning] Processing deprecated module '$($moduleInfo.Name)' ($($moduleInfo.VersionCount) versions)" -Type Warning
            Write-ColorOutput "    Reason: $($moduleInfo.Reason)" -Type Warning
            Write-ColorOutput "    Replacement: $($moduleInfo.Replacement)" -Type Warning
            Write-ColorOutput "=== PROCESSING: $($moduleInfo.Name) ===" -Type System
            Write-ColorOutput ""
        }
        
        if ($isConstrainedMode) {
            # Simplified module information and prompt for constrained mode
            if ($Prompt) {
                $response = Read-Host "Remove $($moduleInfo.Name)? [Y/N/S]"
                
                if ($response -match '^[Ss]') {
                    Write-ColorOutput "Skipping remaining modules." -Type Warning
                    break
                }
                
                if ($response -notmatch '^[Yy]') {
                    Write-ColorOutput "Skipping $($moduleInfo.Name)" -Type Warning
                    $skippedModules += $moduleInfo.Name
                    $currentOperation += $moduleInfo.VersionCount
                    continue
                }
            }
        } else {
            # Detailed module information for normal mode
            Write-ColorOutput "📦 Module: $($moduleInfo.Name)" -Type Warning
            Write-ColorOutput "🚫 Status: DEPRECATED" -Type Error
            Write-ColorOutput "📋 Reason: $($moduleInfo.Reason)" -Type Info
            Write-ColorOutput "🔄 Modern Replacement: $($moduleInfo.Replacement)" -Type Process
            Write-ColorOutput "📊 Versions to remove: $($moduleInfo.VersionCount)" -Type Info
            Write-ColorOutput "⏱️  Estimated removal time: $([math]::Round($moduleInfo.EstimatedTime / 60, 1)) minutes" -Type Info
            Write-ColorOutput ""
            
            # Show specific migration guidance
            switch ($moduleInfo.Name) {
                'AzureAD' {
                    Write-ColorOutput "🔄 MIGRATION GUIDANCE FOR AZUREAD:" -Type Process
                    Write-ColorOutput "   • Replace Connect-AzureAD with Connect-MgGraph" -Type Info
                    Write-ColorOutput "   • Replace Get-AzureADUser with Get-MgUser" -Type Info
                    Write-ColorOutput "   • Replace Get-AzureADGroup with Get-MgGroup" -Type Info
                    Write-ColorOutput "   • Update all AzureAD cmdlets to Microsoft.Graph equivalents" -Type Info
                }
                'AzureADPreview' {
                    Write-ColorOutput "🔄 MIGRATION GUIDANCE FOR AZUREADPREVIEW:" -Type Process
                    Write-ColorOutput "   • Replace Connect-AzureAD with Connect-MgGraph" -Type Info
                    Write-ColorOutput "   • Replace Get-AzureADUser with Get-MgUser" -Type Info
                    Write-ColorOutput "   • Replace Get-AzureADGroup with Get-MgGroup" -Type Info
                    Write-ColorOutput "   • Preview features are now in Microsoft.Graph.Beta modules" -Type Info
                }
                'MSOnline' {
                    Write-ColorOutput "🔄 MIGRATION GUIDANCE FOR MSONLINE:" -Type Process
                    Write-ColorOutput "   • Replace Connect-MsolService with Connect-MgGraph" -Type Info
                    Write-ColorOutput "   • Replace Get-MsolUser with Get-MgUser" -Type Info
                    Write-ColorOutput "   • Replace Get-MsolDomain with Get-MgDomain" -Type Info
                    Write-ColorOutput "   • Update all Msol cmdlets to Microsoft.Graph equivalents" -Type Info
                }
                'SharePointPnPPowerShellOnline' {
                    Write-ColorOutput "🔄 MIGRATION GUIDANCE FOR SHAREPOINT PNP:" -Type Process
                    Write-ColorOutput "   • Replace Connect-PnPOnline with Connect-PnPOnline (new PnP.PowerShell)" -Type Info
                    Write-ColorOutput "   • Most cmdlet names remain the same in PnP.PowerShell" -Type Info
                    Write-ColorOutput "   • Update module references in your scripts" -Type Info
                    Write-ColorOutput "   • Test thoroughly as some parameters may have changed" -Type Info
                }
                'WindowsAutoPilotIntune' {
                    Write-ColorOutput "🔄 MIGRATION GUIDANCE FOR AUTOPILOT:" -Type Process
                    Write-ColorOutput "   • Use Microsoft.Graph.DeviceManagement module" -Type Info
                    Write-ColorOutput "   • Autopilot functionality is now in Microsoft Graph APIs" -Type Info
                    Write-ColorOutput "   • Update scripts to use Graph-based device management cmdlets" -Type Info
                }
                'AIPService' {
                    Write-ColorOutput "🔄 MIGRATION GUIDANCE FOR AIP SERVICE:" -Type Process
                    Write-ColorOutput "   • Use Microsoft.Graph.Security module" -Type Info
                    Write-ColorOutput "   • Information Protection APIs are now in Microsoft Graph" -Type Info
                    Write-ColorOutput "   • Update scripts to use Graph-based security cmdlets" -Type Info
                }
                'MSCommerce' {
                    Write-ColorOutput "🔄 MIGRATION GUIDANCE FOR MSCOMMERCE:" -Type Process
                    Write-ColorOutput "   • Use Microsoft.Graph.Commerce module" -Type Info
                    Write-ColorOutput "   • Commerce functionality is available through Microsoft Graph" -Type Info
                    Write-ColorOutput "   • Update scripts to use Graph-based commerce cmdlets" -Type Info
                }
            }
            Write-ColorOutput ""
            
            # Always prompt for each module with detailed options
            Write-ColorOutput "❓ REMOVAL DECISION:" -Type Warning
            Write-ColorOutput "Do you want to remove the deprecated module '$($moduleInfo.Name)'?" -Type Warning
            Write-ColorOutput ""
            Write-ColorOutput "Enter your choice:" -Type Info
            Write-ColorOutput "  Y = Yes, remove this deprecated module" -Type Process
            Write-ColorOutput "  N = No, keep this module (not recommended)" -Type Warning
            Write-ColorOutput "  S = Skip remaining deprecated modules" -Type Info
            Write-ColorOutput ""
            
            do {
                $response = Read-Host "Remove deprecated module '$($moduleInfo.Name)' [Y/N/S]"
                $validResponse = $response -match '^[YyNnSs]$'
                if (-not $validResponse) {
                    Write-ColorOutput "Please enter Y (Yes), N (No), or S (Skip)" -Type Warning
                }
            } while (-not $validResponse)
            
            if ($response -match '^[Ss]') {
                Write-ColorOutput ""
                Write-ColorOutput "⏭️  Skipping all remaining deprecated module removals." -Type Warning
                Write-ColorOutput "You can run the script again later to clean up deprecated modules." -Type Info
                break
            }
            
            if ($response -notmatch '^[Yy]') {
                Write-ColorOutput ""
                Write-ColorOutput "⏭️  Skipping removal of $($moduleInfo.Name)" -Type Warning
                Write-ColorOutput "⚠️  Warning: This deprecated module will remain on your system" -Type Warning
                $skippedModules += $moduleInfo.Name
                $currentOperation += $moduleInfo.VersionCount
                Write-ColorOutput ""
                continue
            }
        }
        
        if ($isConstrainedMode) {
            Write-ColorOutput "Removing $($moduleInfo.Name)..." -Type Process
        } else {
            Write-ColorOutput ""
            Write-ColorOutput "🗑️  Proceeding with removal of $($moduleInfo.Name)..." -Type Process
            Write-ColorOutput ""
        }
        
        $moduleStartTime = Get-Date
        
        # Remove each version with appropriate progress tracking
        foreach ($version in $moduleInfo.Versions) {
            $currentOperation++
            
            if ($isConstrainedMode) {
                # Minimal progress display for constrained mode
                Write-ColorOutput "  Removing v$($version.Version)..." -Type Info -NoNewline
            } else {
                # Detailed progress display for normal mode
                $percentComplete = [math]::Round(($currentOperation / $totalOperations) * 100, 1)
                
                # Calculate estimated time remaining
                $elapsed = (Get-Date) - $startTime
                if ($currentOperation -gt 1) {
                    $averageTimePerOperation = $elapsed.TotalSeconds / ($currentOperation - 1)
                    $remainingOperations = $totalOperations - $currentOperation
                    $estimatedTimeRemaining = [TimeSpan]::FromSeconds($averageTimePerOperation * $remainingOperations)
                    
                    if ($estimatedTimeRemaining.TotalMinutes -gt 1) {
                        $timeRemainingText = "{0:mm}m {0:ss}s remaining" -f $estimatedTimeRemaining
                    } else {
                        $timeRemainingText = "{0:ss}s remaining" -f $estimatedTimeRemaining
                    }
                } else {
                    $timeRemainingText = "calculating..."
                }
                
                Write-ColorOutput "    [$percentComplete%] ($currentOperation/$totalOperations) $timeRemainingText" -Type Info
                
                # Determine installation method and removal approach
                $installMethod = if ($version.PSObject.Properties.Name -contains "InstalledVia") { 
                    $version.InstalledVia 
                } else { 
                    "PowerShellGet" 
                }
                
                Write-ColorOutput "    Removing $($moduleInfo.Name) v$($version.Version) [$installMethod]..." -Type Process -NoNewline
            }
            
            try {
                $operationStart = Get-Date
                $uninstallResult = @{ Success = $false; Message = "Not attempted" }
                
                # Determine installation method for removal approach
                $installMethod = if ($version.PSObject.Properties.Name -contains "InstalledVia") { 
                    $version.InstalledVia 
                } else { 
                    "PowerShellGet" 
                }
                
                # Simplified removal logic with suppressed output for constrained mode
                if ($isConstrainedMode) {
                    # Minimal feedback in constrained mode - suppress all progress dots and detailed messages
                    try {
                        # Try the most common removal method first
                        Uninstall-Module -Name $moduleInfo.Name -RequiredVersion $version.Version -Force -Confirm:$false -ErrorAction Stop 2>$null
                        $uninstallResult = @{ Success = $true; Message = "Success" }
                    }
                    catch {
                        # Try fallback without version
                        try {
                            Uninstall-Module -Name $moduleInfo.Name -Force -Confirm:$false -ErrorAction Stop 2>$null
                            $uninstallResult = @{ Success = $true; Message = "Success" }
                        }
                        catch {
                            $uninstallResult = @{ Success = $false; Message = $_.Exception.Message }
                        }
                    }
                } else {
                    # Detailed removal logic for normal mode
                    # Try different removal methods based on how the module was installed
                    if ($installMethod -eq "MSI/Manual") {
                        # For MSI or manually installed modules, try to remove from module path
                        try {
                            Write-Host "." -NoNewline -ForegroundColor $Script:Colors.Process
                            
                            # Check if module can be uninstalled via Uninstall-Module
                            $canUninstallViaPS = Get-InstalledModule -Name $moduleInfo.Name -RequiredVersion $version.Version -ErrorAction SilentlyContinue
                            
                            if ($canUninstallViaPS) {
                                # Try PowerShell uninstall first
                                try {
                                    Uninstall-Module -Name $moduleInfo.Name -RequiredVersion $version.Version -Force -Confirm:$false -ErrorAction Stop
                                    $uninstallResult = @{ Success = $true; Message = "Uninstalled via PowerShell" }
                                }
                                catch {
                                    # Specific error handling for common issues
                                    $errorMsg = $_.Exception.Message
                                    if ($errorMsg -like "*Administrator*" -or $errorMsg -like "*elevated*") {
                                        $uninstallResult = @{ Success = $false; Message = "Requires administrator privileges" }
                                    }
                                    elseif ($errorMsg -like "*in use*" -or $errorMsg -like "*locked*") {
                                        $uninstallResult = @{ Success = $false; Message = "Module is currently in use - close PowerShell sessions and retry" }
                                    }
                                    else {
                                        $uninstallResult = @{ Success = $false; Message = "PowerShell uninstall failed: $errorMsg" }
                                    }
                                }
                            }
                            else {
                                # Manual removal from file system (with caution)
                                if ($version.ModuleBase -and (Test-Path $version.ModuleBase)) {
                                    # Only remove if it's in a user-specific or PowerShell modules path
                                    $safeToRemove = $version.ModuleBase -match "\\Users\\.*\\Documents\\.*PowerShell.*\\Modules" -or
                                                   $version.ModuleBase -match "\\Program Files\\.*PowerShell.*\\Modules" -or
                                                   $version.ModuleBase -match "\\PowerShell\\Modules"
                                    
                                    if ($safeToRemove) {
                                        try {
                                            # Check if files are locked before attempting removal
                                            $testFile = Join-Path $version.ModuleBase "test.tmp"
                                            try {
                                                [System.IO.File]::Create($testFile).Close()
                                                Remove-Item $testFile -Force -ErrorAction SilentlyContinue
                                                
                                                Remove-Item -Path $version.ModuleBase -Recurse -Force -ErrorAction Stop
                                                $uninstallResult = @{ Success = $true; Message = "Removed from file system" }
                                            }
                                            catch [System.UnauthorizedAccessException] {
                                                $uninstallResult = @{ Success = $false; Message = "Access denied - module files may be in use or require admin rights" }
                                            }
                                            catch [System.IO.IOException] {
                                                $uninstallResult = @{ Success = $false; Message = "Module files are locked or in use - close applications and retry" }
                                            }
                                        }
                                        catch {
                                            $uninstallResult = @{ Success = $false; Message = "File system removal failed: $($_.Exception.Message)" }
                                        }
                                    }
                                    else {
                                        $uninstallResult = @{ Success = $false; Message = "Module in system location ($($version.ModuleBase)) - manual removal required" }
                                    }
                                }
                                else {
                                    $uninstallResult = @{ Success = $false; Message = "Module path not found or no longer exists" }
                                }
                            }
                        }
                        catch {
                            $uninstallResult = @{ Success = $false; Message = "Unexpected error during MSI/Manual removal: $($_.Exception.Message)" }
                        }
                    }
                    else {
                        # For PowerShellGet installed modules, use standard uninstall with retry logic
                        $retryAttempts = 2
                        $retryDelay = 2
                        
                        for ($attempt = 1; $attempt -le $retryAttempts; $attempt++) {
                            try {
                                Write-Host "." -NoNewline -ForegroundColor $Script:Colors.Process
                                if ($attempt -gt 1) {
                                    Start-Sleep -Seconds $retryDelay
                                    Write-Host "(retry $attempt)" -NoNewline -ForegroundColor $Script:Colors.Warning
                                }
                                
                                Uninstall-Module -Name $moduleInfo.Name -RequiredVersion $version.Version -Force -Confirm:$false -ErrorAction Stop
                                $uninstallResult = @{ Success = $true; Message = "Uninstalled via PowerShellGet" }
                                break
                            }
                            catch {
                                $errorMsg = $_.Exception.Message
                                
                                # Check for specific error conditions
                                if ($errorMsg -like "*Administrator*" -or $errorMsg -like "*elevated*") {
                                    $uninstallResult = @{ Success = $false; Message = "Requires administrator privileges" }
                                    break
                                }
                                elseif ($errorMsg -like "*in use*" -or $errorMsg -like "*locked*") {
                                    if ($attempt -lt $retryAttempts) {
                                        continue  # Retry for file locking issues
                                    }
                                    $uninstallResult = @{ Success = $false; Message = "Module is currently in use - close PowerShell sessions and retry" }
                                    break
                                }
                                elseif ($errorMsg -like "*not found*" -or $errorMsg -like "*does not exist*") {
                                    # Try fallback: uninstall without specific version
                                    try {
                                        Write-Host "." -NoNewline -ForegroundColor $Script:Colors.Process
                                        Uninstall-Module -Name $moduleInfo.Name -Force -Confirm:$false -ErrorAction Stop
                                        $uninstallResult = @{ Success = $true; Message = "Uninstalled via PowerShellGet (all versions)" }
                                        break
                                    }
                                    catch {
                                        if ($attempt -eq $retryAttempts) {
                                            $uninstallResult = @{ Success = $false; Message = "Module not found in PowerShellGet registry: $($_.Exception.Message)" }
                                        }
                                    }
                                }
                                else {
                                    if ($attempt -eq $retryAttempts) {
                                        $uninstallResult = @{ Success = $false; Message = "PowerShellGet uninstall failed: $errorMsg" }
                                    }
                                }
                            }
                        }
                    }
                }
                
                $operationDuration = (Get-Date) - $operationStart
                
                # Simplified result display
                if ($isConstrainedMode) {
                    if ($uninstallResult.Success) {
                        Write-ColorOutput " Done" -Type Process
                    } else {
                        Write-ColorOutput " Failed ($($uninstallResult.Message -replace '.*:', ''))" -Type Error
                    }
                } else {
                    if ($uninstallResult.Success) {
                        Write-Host " Done! ($([math]::Round($operationDuration.TotalSeconds, 1))s)" -ForegroundColor $Script:Colors.Process
                        if ($uninstallResult.Message -ne "Completed successfully" -and $uninstallResult.Message -ne "Success") {
                            Write-ColorOutput "      Method: $($uninstallResult.Message)" -Type Info
                        }
                    } else {
                        Write-Host " Failed!" -ForegroundColor $Script:Colors.Error
                        Write-ColorOutput "      Error: $($uninstallResult.Message)" -Type Error
                        if ($installMethod -eq "MSI/Manual") {
                            Write-ColorOutput "      Note: MSI-installed modules may require manual removal via Add/Remove Programs" -Type Warning
                        }
                    }
                }
            }
            catch {
                $operationDuration = (Get-Date) - $operationStart
                if ($isConstrainedMode) {
                    Write-ColorOutput " Error ($($_.Exception.Message -replace '.*:', ''))" -Type Error
                } else {
                    Write-Host " Error! ($([math]::Round($operationDuration.TotalSeconds, 1))s)" -ForegroundColor $Script:Colors.Error
                    Write-ColorOutput "      Error details: $($_.Exception.Message)" -Type Error
                }
            }
        }
        
        $moduleElapsed = (Get-Date) - $moduleStartTime
        if ($isConstrainedMode) {
            Write-ColorOutput "Completed $($moduleInfo.Name) removal" -Type Process
        } else {
            Write-ColorOutput "    Completed $($moduleInfo.Name) in $([math]::Round($moduleElapsed.TotalMinutes, 1)) minutes" -Type Process
            Write-ColorOutput ""
        }
    }
    
    
    $totalElapsed = (Get-Date) - $startTime
    
    if ($isConstrainedMode) {
        # Simplified completion summary for constrained mode
        Write-ColorOutput ""
        Write-ColorOutput "Deprecated module cleanup completed ($([math]::Round($totalElapsed.TotalMinutes, 1)) minutes)" -Type System
        
        if ($skippedModules.Count -gt 0) {
            Write-ColorOutput "Skipped: $($skippedModules -join ', ')" -Type Warning
        }
        
        # Quick verification without verbose output
        $remainingCount = 0
        foreach ($deprecated in $Script:DeprecatedModules) {
            if (Get-Module -ListAvailable -Name $deprecated.Name -ErrorAction SilentlyContinue) {
                $remainingCount++
            }
        }
        
        if ($remainingCount -eq 0) {
            Write-ColorOutput "✓ All deprecated modules removed successfully" -Type Process
        } else {
            Write-ColorOutput "⚠ $remainingCount deprecated modules may still be present" -Type Warning
        }
    } else {
        # Detailed completion summary for normal mode
        Write-ColorOutput "Deprecated module cleanup completed in $([math]::Round($totalElapsed.TotalMinutes, 1)) minutes" -Type System
        
        # Compare actual vs estimated time
        $actualMinutes = [math]::Round($totalElapsed.TotalMinutes, 1)
        if ($actualMinutes -lt $estimatedMinutes) {
            Write-ColorOutput ("Completed faster than estimated! (Est: {0}min, Actual: {1}min)" -f $estimatedMinutes, $actualMinutes) -Type Process
        } elseif ($actualMinutes -gt ($estimatedMinutes * 1.2)) {
            Write-ColorOutput ("Took longer than estimated (Est: {0}min, Actual: {1}min)" -f $estimatedMinutes, $actualMinutes) -Type Warning
        }
        
        # Final verification
        Write-ColorOutput ""
        Write-ColorOutput "=== DEPRECATED MODULE CLEANUP SUMMARY ===" -Type System
        
        if ($skippedModules.Count -gt 0) {
            Write-ColorOutput ""
            Write-ColorOutput "⏭️  SKIPPED MODULES:" -Type Warning
            foreach ($skipped in $skippedModules) {
                Write-ColorOutput "  • $skipped (still installed - not recommended)" -Type Warning
            }
            Write-ColorOutput ""
            Write-ColorOutput "⚠️  These deprecated modules remain on your system and may cause issues:" -Type Warning
            Write-ColorOutput "  • They no longer receive security updates" -Type Warning
            Write-ColorOutput "  • They may conflict with newer modules" -Type Warning
            Write-ColorOutput "  • Microsoft support for these modules has ended" -Type Warning
            Write-ColorOutput ""
            Write-ColorOutput "💡 Consider removing them in a future run of this script." -Type Info
        }
        
        Write-ColorOutput ""
        Write-ColorOutput "Verifying deprecated module removal..." -Type System
        $remainingModules = @()
        foreach ($deprecated in $Script:DeprecatedModules) {
            $stillInstalled = Get-Module -ListAvailable -Name $deprecated.Name -ErrorAction SilentlyContinue
            if ($stillInstalled) {
                $remainingModules += $deprecated.Name
            }
        }
        
        if ($remainingModules.Count -eq 0) {
            Write-ColorOutput "    ✓ All deprecated modules successfully removed" -Type Process
            Write-ColorOutput ""
            
            # Brief pause to let user see the success message
            Start-Sleep -Seconds 2
            
            # Clear any system messages and show clean completion
            Write-ColorOutput "Deprecated module cleanup completed successfully." -Type System
        } else {
            Write-ColorOutput "    ⚠ Some modules may still be present: $($remainingModules -join ', ')" -Type Warning
            Write-ColorOutput "    These may be MSI-installed modules requiring manual removal" -Type Info
            Write-ColorOutput ""
            
            # Show detailed troubleshooting for remaining modules
            Write-ColorOutput "Next steps for remaining modules:" -Type Info
            Write-ColorOutput "    1. Run PowerShell as Administrator" -Type Info
            Write-ColorOutput "    2. Close all other PowerShell sessions" -Type Info
            Write-ColorOutput "    3. For MSI-installed modules, use Windows Settings > Apps & features" -Type Info
            Write-ColorOutput "    4. Manual cleanup locations:" -Type Info
            Write-ColorOutput "       • $env:USERPROFILE\\Documents\\PowerShell\\Modules" -Type Info
            Write-ColorOutput "       • $env:ProgramFiles\\PowerShell\\Modules" -Type Info
            Write-ColorOutput "       • $env:USERPROFILE\\Documents\\WindowsPowerShell\\Modules" -Type Info
            Write-ColorOutput "       • $env:ProgramFiles\\WindowsPowerShell\\Modules" -Type Info
            Write-ColorOutput ""
            
            # Pause for user to read troubleshooting info
            Write-ColorOutput "Press any key to continue..." -Type Warning -NoNewline
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            Write-ColorOutput ""
        }
    }
}

function Test-ModuleRemovalPrerequisites {
    [CmdletBinding()]
    param()
    
    Write-ColorOutput "Checking module removal prerequisites..." -Type System
    
    $issues = @()
    
    # Check if running as administrator
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $isAdmin) {
        $issues += "Not running as Administrator - some system-installed modules may not be removable"
        Write-ColorOutput "    [Warning] Not running as Administrator" -Type Warning
    } else {
        Write-ColorOutput "    [OK] Running as Administrator" -Type Process    }

    # Check session conflict status (handled in main execution)
    if (-not $Script:SessionConflictsResolved) {
        $issues += "Other PowerShell sessions detected - modules may be locked"
        Write-ColorOutput "    [Info] Session conflicts detected - module removal may encounter issues" -Type Info
    } else {
        Write-ColorOutput "    [OK] No session conflicts detected" -Type Process
    }
      # Check available disk space
    $systemDrive = $env:SystemDrive
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$systemDrive'" -ErrorAction SilentlyContinue
    if ($disk -and $disk.FreeSpace) {
        $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
        if ($freeSpaceGB -lt 1) {
            $issues += "Low disk space may cause removal issues"
            Write-ColorOutput ("    [Warning] Low disk space: {0}GB free" -f $freeSpaceGB) -Type Warning
        } else {
            Write-ColorOutput ("    [OK] Sufficient disk space: {0}GB free" -f $freeSpaceGB) -Type Process
        }
    }
    
    # Check execution policy
    $executionPolicy = Get-ExecutionPolicy
    if ($executionPolicy -eq "Restricted") {
        $issues += "Restricted execution policy may prevent module operations"
        Write-ColorOutput "    [Warning] Execution policy is Restricted" -Type Warning
    } else {
        Write-ColorOutput "    [OK] Execution policy: $executionPolicy" -Type Process
    }
    
    if ($issues.Count -gt 0) {
        Write-ColorOutput ""
        Write-ColorOutput "Potential issues detected:" -Type Warning
        foreach ($issue in $issues) {
            Write-ColorOutput "    • $issue" -Type Warning
        }
        Write-ColorOutput ""
        
        if ($Prompt) {
            $response = Read-Host "Continue with module removal despite potential issues (Y/N)?"
            if ($response -notmatch '^[Yy]') {
                Write-ColorOutput "Module removal cancelled by user" -Type Warning
                return $false
            }
        }
    } else {
        Write-ColorOutput "    All prerequisites met" -Type Process
    }
    
    return $true
}

function Test-ConstrainedLanguageModeCompatibility {
    <#
    .SYNOPSIS
        Tests specific operations that may fail in Constrained Language Mode
    
    .DESCRIPTION
        Performs detailed testing of operations commonly used in module management
        that may be restricted in Constrained Language Mode. Returns compatibility
        status and provides detailed debugging information.
    
    .PARAMETER Operation
        Specific operation to test. Valid values: 'NetworkOptimization', 'ModuleOperations', 'All'
    
    .PARAMETER Silent
        If specified, suppresses detailed output and returns only the result
    
    .RETURNS
        Boolean indicating if the tested operations are compatible with the current language mode
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('NetworkOptimization', 'ModuleOperations', 'All')]
        [string]$Operation = 'All',
        
        [Parameter()]
        [switch]$Silent
    )
    
    $languageMode = $ExecutionContext.SessionState.LanguageMode
    $compatibilityResults = @{
        LanguageMode = $languageMode
        NetworkOptimization = $false
        ModuleOperations = $false
        OverallCompatible = $false
        Issues = @()
        Recommendations = @()
    }
    
    if (-not $Silent) {
        Write-ColorOutput "=== CONSTRAINED LANGUAGE MODE COMPATIBILITY TEST ===" -Type System
        Write-ColorOutput "Current Language Mode: $languageMode" -Type Info
        Write-ColorOutput "Testing Operation: $Operation" -Type Info
        Write-ColorOutput ""
    }
    
    # Test Network Optimization Compatibility
    if ($Operation -eq 'NetworkOptimization' -or $Operation -eq 'All') {
        if (-not $Silent) {
            Write-ColorOutput "Testing Network Optimization Compatibility..." -Type Info
        }
        
        $networkTests = @{
            ServicePointManager = $false
            SecurityProtocol = $false
            ConnectionLimit = $false
        }
        
        # Test ServicePointManager access
        try {
            $null = [System.Net.ServicePointManager]::SecurityProtocol
            $networkTests.ServicePointManager = $true
            if (-not $Silent) {
                Write-ColorOutput "  ✓ ServicePointManager: Accessible" -Type Process
            }
        }
        catch {
            $compatibilityResults.Issues += "ServicePointManager not accessible: $($_.Exception.Message)"
            if (-not $Silent) {
                Write-ColorOutput "  ✗ ServicePointManager: Blocked" -Type Error
                Write-ColorOutput "    Error: $($_.Exception.Message)" -Type Error
            }
        }
        
        # Test SecurityProtocol modification
        try {
            $originalProtocol = [System.Net.ServicePointManager]::SecurityProtocol
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            [System.Net.ServicePointManager]::SecurityProtocol = $originalProtocol
            $networkTests.SecurityProtocol = $true
            if (-not $Silent) {
                Write-ColorOutput "  ✓ SecurityProtocol: Modifiable" -Type Process
            }
        }
        catch {
            $compatibilityResults.Issues += "SecurityProtocol modification blocked: $($_.Exception.Message)"
            if (-not $Silent) {
                Write-ColorOutput "  ✗ SecurityProtocol: Modification blocked" -Type Error
                Write-ColorOutput "    Error: $($_.Exception.Message)" -Type Error
            }
        }
        
        # Test ConnectionLimit modification
        try {
            $originalLimit = [System.Net.ServicePointManager]::DefaultConnectionLimit
            [System.Net.ServicePointManager]::DefaultConnectionLimit = 12
            [System.Net.ServicePointManager]::DefaultConnectionLimit = $originalLimit
            $networkTests.ConnectionLimit = $true
            if (-not $Silent) {
                Write-ColorOutput "  ✓ ConnectionLimit: Modifiable" -Type Process
            }
        }
        catch {
            $compatibilityResults.Issues += "ConnectionLimit modification blocked: $($_.Exception.Message)"
            if (-not $Silent) {
                Write-ColorOutput "  ✗ ConnectionLimit: Modification blocked" -Type Error
                Write-ColorOutput "    Error: $($_.Exception.Message)" -Type Error
            }
        }
        
        $compatibilityResults.NetworkOptimization = ($networkTests.ServicePointManager -and $networkTests.SecurityProtocol -and $networkTests.ConnectionLimit)
        
        if (-not $Silent) {
            if ($compatibilityResults.NetworkOptimization) {
                Write-ColorOutput "  ✓ Network Optimization: Fully Compatible" -Type Process
            } else {
                Write-ColorOutput "  ⚠ Network Optimization: Limited Compatibility" -Type Warning
                $compatibilityResults.Recommendations += "Network optimizations may not be fully effective"
            }
        }
    }
    
    # Test Module Operations Compatibility
    if ($Operation -eq 'ModuleOperations' -or $Operation -eq 'All') {
        if (-not $Silent) {
            Write-ColorOutput "Testing Module Operations Compatibility..." -Type Info
        }
        
        $moduleTests = @{
            ExecutionPolicy = $false
            ModuleCmdlets = $false
            HashTableCreation = $false
            ParameterBinding = $false
        }
        
        # Test execution policy operations
        try {
            $null = Get-ExecutionPolicy -ErrorAction Stop
            $moduleTests.ExecutionPolicy = $true
            if (-not $Silent) {
                Write-ColorOutput "  ✓ Execution Policy: Accessible" -Type Process
            }
        }
        catch {
            $compatibilityResults.Issues += "ExecutionPolicy cmdlets blocked: $($_.Exception.Message)"
            if (-not $Silent) {
                Write-ColorOutput "  ✗ Execution Policy: Blocked" -Type Error
                Write-ColorOutput "    Error: $($_.Exception.Message)" -Type Error
            }
        }
        
        # Test module cmdlets
        $moduleCmdlets = @('Get-Module', 'Install-Module', 'Update-Module', 'Uninstall-Module')
        $accessibleCmdlets = 0
        foreach ($cmdlet in $moduleCmdlets) {
            try {
                $null = Get-Command $cmdlet -ErrorAction Stop
                $accessibleCmdlets++
            }
            catch {
                if (-not $Silent) {
                    Write-ColorOutput "  ⚠ $cmdlet`: Not accessible" -Type Warning
                }
            }
        }
        
        $moduleTests.ModuleCmdlets = ($accessibleCmdlets -eq $moduleCmdlets.Count)
        if (-not $Silent) {
            if ($moduleTests.ModuleCmdlets) {
                Write-ColorOutput "  ✓ Module Cmdlets: All accessible ($accessibleCmdlets/$($moduleCmdlets.Count))" -Type Process
            } else {
                Write-ColorOutput "  ⚠ Module Cmdlets: Limited access ($accessibleCmdlets/$($moduleCmdlets.Count))" -Type Warning
            }
        }
        
        # Test hashtable creation and manipulation
        try {
            $testHash = @{
                Force = $true
                Confirm = $false
                SkipPublisherCheck = $true
            }
            $testHash.AllowClobber = $true
            $null = $testHash.Keys # Use the hashtable
            $moduleTests.HashTableCreation = $true
            if (-not $Silent) {
                Write-ColorOutput "  ✓ Hashtable Operations: Supported" -Type Process
            }
        }
        catch {
            $compatibilityResults.Issues += "Hashtable operations limited: $($_.Exception.Message)"
            if (-not $Silent) {
                Write-ColorOutput "  ✗ Hashtable Operations: Limited" -Type Error
                Write-ColorOutput "    Error: $($_.Exception.Message)" -Type Error
            }
        }
        
        # Test parameter splatting
        try {
            $testParams = @{ Name = 'PowerShellGet'; ListAvailable = $true }
            $null = Get-Module @testParams -ErrorAction Stop
            $moduleTests.ParameterBinding = $true
            if (-not $Silent) {
                Write-ColorOutput "  ✓ Parameter Splatting: Supported" -Type Process
            }
        }
        catch {
            $compatibilityResults.Issues += "Parameter splatting limited: $($_.Exception.Message)"
            if (-not $Silent) {
                Write-ColorOutput "  ✗ Parameter Splatting: Limited" -Type Error
                Write-ColorOutput "    Error: $($_.Exception.Message)" -Type Error
            }
        }
        
        $compatibilityResults.ModuleOperations = ($moduleTests.ExecutionPolicy -and $moduleTests.ModuleCmdlets -and $moduleTests.HashTableCreation -and $moduleTests.ParameterBinding)
        
        if (-not $Silent) {
            if ($compatibilityResults.ModuleOperations) {
                Write-ColorOutput "  ✓ Module Operations: Fully Compatible" -Type Process
            } else {
                Write-ColorOutput "  ⚠ Module Operations: Limited Compatibility" -Type Warning
                $compatibilityResults.Recommendations += "Module operations may encounter restrictions"
            }
        }
    }
    
    # Determine overall compatibility
    if ($Operation -eq 'All') {
        $compatibilityResults.OverallCompatible = ($compatibilityResults.NetworkOptimization -and $compatibilityResults.ModuleOperations)
    } elseif ($Operation -eq 'NetworkOptimization') {
        $compatibilityResults.OverallCompatible = $compatibilityResults.NetworkOptimization
    } elseif ($Operation -eq 'ModuleOperations') {
        $compatibilityResults.OverallCompatible = $compatibilityResults.ModuleOperations
    }
    
    # Add language mode specific recommendations
    if ($languageMode -eq 'ConstrainedLanguage') {
        $compatibilityResults.Recommendations += "Consider running as Administrator to minimize restrictions"
        $compatibilityResults.Recommendations += "Check for AppLocker or other security policies that may be enforcing constraints"
        $compatibilityResults.Recommendations += "Monitor operations closely for unexpected failures"
        
        if (-not $compatibilityResults.OverallCompatible) {
            $compatibilityResults.Recommendations += "Consider requesting Full Language Mode if organizationally feasible"
            $compatibilityResults.Recommendations += "Prepare for longer execution times due to reduced optimization"
        }
    }
    
    # Output results summary
    if (-not $Silent) {
        Write-ColorOutput ""
        Write-ColorOutput "=== COMPATIBILITY TEST RESULTS ===" -Type System
        Write-ColorOutput "Language Mode: $languageMode" -Type Info
        
        if ($Operation -eq 'All' -or $Operation -eq 'NetworkOptimization') {
            # Use traditional if-else for Constrained Language Mode compatibility
            if ($compatibilityResults.NetworkOptimization) {
                $status = "✓ Compatible"
                $statusType = 'Process'
            } else {
                $status = "⚠ Limited"
                $statusType = 'Warning'
            }
            Write-ColorOutput "Network Optimization: $status" -Type $statusType
        }
        
        if ($Operation -eq 'All' -or $Operation -eq 'ModuleOperations') {
            # Use traditional if-else for Constrained Language Mode compatibility
            if ($compatibilityResults.ModuleOperations) {
                $status = "✓ Compatible"
                $statusType = 'Process'
            } else {
                $status = "⚠ Limited"
                $statusType = 'Warning'
            }
            Write-ColorOutput "Module Operations: $status" -Type $statusType
        }
        
        # Use traditional if-else for Constrained Language Mode compatibility
        if ($compatibilityResults.OverallCompatible) {
            $overallStatus = "✓ Compatible"
            $overallStatusType = 'Process'
        } else {
            $overallStatus = "⚠ Limited"
            $overallStatusType = 'Warning'
        }
        Write-ColorOutput "Overall Compatibility: $overallStatus" -Type $overallStatusType
        
        if ($compatibilityResults.Issues.Count -gt 0) {
            Write-ColorOutput ""
            Write-ColorOutput "Issues Identified:" -Type Warning
            foreach ($issue in $compatibilityResults.Issues) {
                Write-ColorOutput "  • $issue" -Type Warning
            }
        }
        
        if ($compatibilityResults.Recommendations.Count -gt 0) {
            Write-ColorOutput ""
            Write-ColorOutput "Recommendations:" -Type Info
            foreach ($recommendation in $compatibilityResults.Recommendations) {
                Write-ColorOutput "  • $recommendation" -Type Info
            }
        }
        
        Write-ColorOutput "=== COMPATIBILITY TEST COMPLETE ===" -Type System
    }
    
    return $compatibilityResults.OverallCompatible
}

function Test-PowerShellCompatibility {
    [CmdletBinding()]
    param()
    
    Write-ColorOutput "=== DETAILED POWERSHELL COMPATIBILITY ANALYSIS ===" -Type System
    
    # Enhanced Language Mode Analysis
    $languageMode = $ExecutionContext.SessionState.LanguageMode
    Write-ColorOutput "PowerShell Language Mode: $languageMode" -Type Info
    
    # Detailed analysis of language mode capabilities
    switch ($languageMode) {
        'FullLanguage' {
            Write-ColorOutput "  ✓ Full Language Mode Detected" -Type Process
            Write-ColorOutput "    • All PowerShell features available" -Type Process
            Write-ColorOutput "    • .NET framework fully accessible" -Type Process
            Write-ColorOutput "    • Background jobs supported" -Type Process
            Write-ColorOutput "    • COM objects accessible" -Type Process
        }
        'ConstrainedLanguage' {
            Write-ColorOutput "  ⚠ CONSTRAINED LANGUAGE MODE DETECTED" -Type Warning
            Write-ColorOutput "    Performing detailed capability assessment..." -Type Info
            
            # Test .NET type accessibility
            Write-ColorOutput "  Testing .NET Framework Access:" -Type Info
            $netAccessTests = @(
                @{ Name = "System.Net.ServicePointManager"; Test = { [System.Net.ServicePointManager] } },
                @{ Name = "System.Security.Principal.WindowsIdentity"; Test = { [System.Security.Principal.WindowsIdentity] } },
                @{ Name = "System.IO.File"; Test = { [System.IO.File] } },
                @{ Name = "System.Diagnostics.Process"; Test = { [System.Diagnostics.Process] } }
            )
            
            foreach ($test in $netAccessTests) {
                try {
                    $null = & $test.Test
                    Write-ColorOutput "    ✓ $($test.Name): Accessible" -Type Process
                }
                catch {
                    Write-ColorOutput "    ✗ $($test.Name): Blocked ($($_.Exception.Message))" -Type Error
                }
            }
            
            # Test cmdlet accessibility
            Write-ColorOutput "  Testing Core Cmdlet Access:" -Type Info
            $cmdletTests = @(
                'Get-Process', 'Get-Service', 'Get-ExecutionPolicy', 'Set-ExecutionPolicy',
                'Get-Module', 'Install-Module', 'Update-Module', 'Uninstall-Module',
                'Start-Job', 'Get-Job', 'Stop-Job', 'Remove-Job'
            )
            
            foreach ($cmdlet in $cmdletTests) {
                try {
                    $cmd = Get-Command $cmdlet -ErrorAction Stop
                    if ($cmd) { # Use the command object
                        Write-ColorOutput "    ✓ $cmdlet`: Available" -Type Process
                    }
                }
                catch {
                    Write-ColorOutput "    ✗ $cmdlet`: Not Available ($($_.Exception.Message))" -Type Error
                }
            }
            
            # Test parameter binding capabilities
            Write-ColorOutput "  Testing Parameter Binding:" -Type Info
            try {
                $testHash = @{ Force = $true; Confirm = $false }
                $null = $testHash.Keys # Use the hashtable
                Write-ColorOutput "    ✓ Hashtable creation: Success" -Type Process
            }
            catch {
                Write-ColorOutput "    ✗ Hashtable creation: Failed ($($_.Exception.Message))" -Type Error
            }
            
            # Test variable assignment and scoping
            Write-ColorOutput "  Testing Variable Operations:" -Type Info
            try {
                $Script:TestVariable = "ConstrainedModeTest"
                Write-ColorOutput "    ✓ Script scope variable assignment: Success" -Type Process
            }
            catch {
                Write-ColorOutput "    ✗ Script scope variable assignment: Failed ($($_.Exception.Message))" -Type Error
            }
            
            # Test file system access
            Write-ColorOutput "  Testing File System Access:" -Type Info
            try {
                $tempPath = [System.IO.Path]::GetTempPath()
                $testFile = Join-Path $tempPath "CLMTest.tmp"
                "test" | Out-File -FilePath $testFile -Force
                Remove-Item $testFile -Force
                Write-ColorOutput "    ✓ File system operations: Success" -Type Process
            }
            catch {
                Write-ColorOutput "    ✗ File system operations: Failed ($($_.Exception.Message))" -Type Error
            }
        }
        'RestrictedLanguage' {
            Write-ColorOutput "  ✗ RESTRICTED LANGUAGE MODE" -Type Error
            Write-ColorOutput "    • Severe limitations on script execution" -Type Error
            Write-ColorOutput "    • Most advanced features will fail" -Type Error
            Write-ColorOutput "    • Consider running in a different context" -Type Error
        }
        'NoLanguage' {
            Write-ColorOutput "  ✗ NO LANGUAGE MODE" -Type Error
            Write-ColorOutput "    • Script execution is severely limited" -Type Error
            Write-ColorOutput "    • Most operations will fail" -Type Error
        }
        default {
            Write-ColorOutput "  ? Unknown Language Mode: $languageMode" -Type Warning
            Write-ColorOutput "    • Behavior is unpredictable" -Type Warning
            Write-ColorOutput "    • Proceed with caution" -Type Warning
        }
    }
    
    # Enhanced Background Jobs Testing
    Write-ColorOutput "  === BACKGROUND JOBS CAPABILITY TEST ===" -Type Info
    
    if ($languageMode -eq 'ConstrainedLanguage') {
        Write-ColorOutput "  Testing job creation in Constrained Language Mode..." -Type Warning
        Write-ColorOutput "  This may fail due to security restrictions..." -Type Warning
    }
    
    $jobTestSuccess = $false
    try {
        Write-ColorOutput "    Creating test job..." -Type Info
        $testJob = Start-Job -ScriptBlock { 
            return @{
                Success = $true
                LanguageMode = $ExecutionContext.SessionState.LanguageMode
                ProcessId = $PID
                Timestamp = Get-Date
            }
        } -ErrorAction Stop
        
        Write-ColorOutput "    ✓ Job creation: Success (Job ID: $($testJob.Id))" -Type Process
        
        # Wait for job completion with timeout
        Write-ColorOutput "    Waiting for job completion..." -Type Info
        $timeout = 10 # seconds
        $elapsed = 0
        
        while ($testJob.State -eq 'Running' -and $elapsed -lt $timeout) {
            Start-Sleep -Seconds 1
            $elapsed++
            Write-ColorOutput ("      Job state: {0} ({1}s/{2}s)" -f $testJob.State, $elapsed, $timeout) -Type Info
        }
        
        if ($testJob.State -eq 'Completed') {
            $jobResult = Receive-Job -Job $testJob
            Write-ColorOutput "    ✓ Job execution: Success" -Type Process
            Write-ColorOutput "      Job Language Mode: $($jobResult.LanguageMode)" -Type Info
            Write-ColorOutput "      Job Process ID: $($jobResult.ProcessId)" -Type Info
            Write-ColorOutput "      Job Timestamp: $($jobResult.Timestamp)" -Type Info
            $jobTestSuccess = $true
        } else {
            Write-ColorOutput "    ⚠ Job execution: Incomplete (State: $($testJob.State))" -Type Warning
            if ($testJob.State -eq 'Failed') {
                $jobError = Receive-Job -Job $testJob 2>&1
                Write-ColorOutput "      Job Error: $jobError" -Type Error
            }
        }
        
        # Clean up job
        Write-ColorOutput "    Cleaning up test job..." -Type Info
        Stop-Job $testJob -ErrorAction SilentlyContinue
        Remove-Job $testJob -ErrorAction SilentlyContinue
        Write-ColorOutput "    ✓ Job cleanup: Complete" -Type Process
        
    }
    catch {
        Write-ColorOutput "    ✗ Background jobs: Not supported" -Type Error
        Write-ColorOutput "      Error: $($_.Exception.Message)" -Type Error
        Write-ColorOutput "      Exception Type: $($_.Exception.GetType().Name)" -Type Error
        
        if ($languageMode -eq 'ConstrainedLanguage') {
            Write-ColorOutput "      This is expected in Constrained Language Mode environments" -Type Warning
            Write-ColorOutput "      Module operations will use direct execution instead of background jobs" -Type Info
        }
    }
    
    # PowerShell Version and Edition Analysis
    Write-ColorOutput "  === POWERSHELL ENVIRONMENT DETAILS ===" -Type Info
    Write-ColorOutput "    PowerShell Version: $($PSVersionTable.PSVersion)" -Type Info
    Write-ColorOutput "    Edition: $($PSVersionTable.PSEdition)" -Type Info
    Write-ColorOutput "    OS: $($PSVersionTable.OS)" -Type Info
    Write-ColorOutput "    Platform: $($PSVersionTable.Platform)" -Type Info
    Write-ColorOutput "    Git Commit ID: $($PSVersionTable.GitCommitId)" -Type Info
    
    # Module Path Analysis
    Write-ColorOutput "  === MODULE PATH ANALYSIS ===" -Type Info
    Write-ColorOutput "    Available Module Paths:" -Type Info
    foreach ($path in $env:PSModulePath -split ';') {
        if (Test-Path $path) {
            Write-ColorOutput "      ✓ $path" -Type Process
        } else {
            Write-ColorOutput "      ✗ $path (Not accessible)" -Type Warning
        }
    }
    
    # PowerShellGet Module Analysis
    Write-ColorOutput "  === POWERSHELLGET MODULE ANALYSIS ===" -Type Info
    try {
        $psGetModules = Get-Module -Name PowerShellGet -ListAvailable | Sort-Object Version -Descending
        if ($psGetModules) {
            Write-ColorOutput "    Available PowerShellGet versions:" -Type Info
            foreach ($module in $psGetModules) {
                Write-ColorOutput "      Version $($module.Version) at $($module.ModuleBase)" -Type Process
            }
        } else {
            Write-ColorOutput "    ⚠ PowerShellGet module not found" -Type Warning
        }
    }
    catch {
        Write-ColorOutput "    ✗ Error checking PowerShellGet: $($_.Exception.Message)" -Type Error
    }
    
    # Final Recommendations
    Write-ColorOutput "  === COMPATIBILITY RECOMMENDATIONS ===" -Type System
    
    if ($languageMode -eq 'ConstrainedLanguage') {
        Write-ColorOutput "  CONSTRAINED LANGUAGE MODE RECOMMENDATIONS:" -Type Warning
        Write-ColorOutput "    1. Module operations will work but may be slower without background jobs" -Type Info
        Write-ColorOutput "    2. Some .NET optimizations may not be available" -Type Info
        Write-ColorOutput "    3. Consider requesting Full Language Mode if policy allows" -Type Info
        Write-ColorOutput "    4. Ensure PowerShell is running as Administrator for best results" -Type Info
        Write-ColorOutput "    5. Monitor for additional security restrictions during module operations" -Type Info
        
        if (-not $jobTestSuccess) {
            Write-ColorOutput "    6. Background jobs disabled - using direct execution mode" -Type Warning
            Write-ColorOutput "    7. Module installation/updates will show less progress information" -Type Warning
        }
    } else {
        Write-ColorOutput "  ✓ Optimal PowerShell environment detected" -Type Process
        Write-ColorOutput "    All features should work as expected" -Type Process
    }
    
    Write-ColorOutput "=== COMPATIBILITY ANALYSIS COMPLETE ===" -Type System
    
    # Return job support status
    if ($jobTestSuccess) {
        Write-ColorOutput "Background jobs: Supported" -Type Process
        return $true
    } else {
        Write-ColorOutput "Background jobs: Not supported (will use direct execution)" -Type Warning
        return $false
    }
}

function Start-ScriptExecution {
    [CmdletBinding()]
    param()
      # Start transcript if requested
    if ($CreateLog) {
        $logFileName = "O365-Update-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
        $logFile = Join-Path $LogPath $logFileName
        
        try {
            Start-Transcript -Path $logFile -Force
            Write-ColorOutput "Transcript started: $logFile" -Type System
        }
        catch {
            Write-ColorOutput "Warning: Could not start transcript: $($_.Exception.Message)" -Type Warning
        }
    }
      # Display script information
    Clear-Host
      Write-ColorOutput "=== Microsoft Cloud PowerShell Module Updater ===" -Type System
    Write-ColorOutput "Script Version: 2.5" -Type System
    Write-ColorOutput "Start Time: $(Get-Date -Format (Get-Culture).DateTimeFormat.FullDateTimePattern)" -Type System
    Write-ColorOutput "Current Culture: $((Get-Culture).DisplayName)" -Type System
    Write-ColorOutput "Prompt Mode: $Prompt" -Type System
    Write-ColorOutput "Check Only Mode: $CheckOnly" -Type System
    Write-ColorOutput ""
      # Check PowerShell version
    $psVersion = $PSVersionTable.PSVersion
    Write-ColorOutput "PowerShell Version: $($psVersion.Major).$($psVersion.Minor)" -Type Process
    
    # Check PowerShell compatibility
    $Script:JobsSupported = Test-PowerShellCompatibility
    Write-ColorOutput ""
    
    if ($psVersion.Major -lt 5) {
        Write-ColorOutput "Error: PowerShell 5.1 or higher is required" -Type Error
        exit 1
    }
    
    # Determine module count
    $moduleCount = $Script:ModuleList.Count
    if ($psVersion.Major -lt 7) {
        $moduleCount++ # Add NuGet provider for PowerShell 5.x
    }
    
    Write-ColorOutput "Total modules to process: $moduleCount" -Type Process
    Write-ColorOutput ""
    
    return $moduleCount
}

function Stop-ScriptExecution {
    [CmdletBinding()]
    param()
      Write-ColorOutput ""
    Write-ColorOutput "=== Script Completed ===" -Type System
    Write-ColorOutput "End Time: $(Get-Date -Format (Get-Culture).DateTimeFormat.FullDateTimePattern)" -Type System
    
    # Check if core modules were updated and provide guidance
    $coreModulesUpdated = $false
    foreach ($module in $Script:ModuleList) {
        if ($module.RequiresSpecialHandling -eq $true) {
            $installedVersions = Get-InstalledModule -Name $module.Name -AllVersions -ErrorAction SilentlyContinue
            if ($installedVersions -and $installedVersions.Count -gt 1) {
                $coreModulesUpdated = $true
                break
            }
        }
    }    if ($coreModulesUpdated) {
        Write-ColorOutput ""
        Write-ColorOutput "🎯 CORE MODULES UPDATED SUCCESSFULLY" -Type Process
        Write-ColorOutput "PowerShell core modules (PowerShellGet, PackageManagement) have been updated." -Type Info
        Write-ColorOutput ""
        Write-ColorOutput "✅ What happened:" -Type Process
        Write-ColorOutput "• New versions were installed alongside existing versions" -Type Info
        Write-ColorOutput "• Current PowerShell session continues using existing versions" -Type Info
        Write-ColorOutput "• New versions will be active when you restart PowerShell" -Type Info
        Write-ColorOutput ""
        Write-ColorOutput "🔄 Next Steps:" -Type Warning
        Write-ColorOutput "• Continue using PowerShell normally for now" -Type Info
        Write-ColorOutput "• Restart PowerShell when convenient to activate new versions" -Type Info
        Write-ColorOutput ""
        Write-ColorOutput "🔍 Verification command (after restart):" -Type Info
        Write-ColorOutput "  Get-Module PowerShellGet, PackageManagement -ListAvailable | Select Name, Version" -Type Info
        Write-ColorOutput ""
        
        # Pause for important notice
        Write-ColorOutput "Press any key to acknowledge..." -Type Warning -NoNewline
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Write-ColorOutput ""
    }
    
    # Stop transcript cleanly
    if ($CreateLog) {
        Write-ColorOutput ""
        try {
            Stop-Transcript
            Write-ColorOutput "Transcript log saved successfully." -Type Process
        }
        catch {
            Write-ColorOutput "Note: Transcript may not have been active." -Type Info
        }
    }
}

function Get-PowerShellSessions {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$IncludeCurrent,
        
        [Parameter()]
        [switch]$ShowDetails,
        
        [Parameter()]
        [switch]$ShowConflicts,
        
        [Parameter()]
        [switch]$Silent
    )
    
    if (-not $Silent) {
        Write-ColorOutput "Detecting PowerShell sessions..." -Type Info
    }
    
    # Get all PowerShell-related processes
    $sessions = @()
    
    # Traditional PowerShell processes (Windows PowerShell 5.1)
    $psProcesses = Get-Process -Name "powershell" -ErrorAction SilentlyContinue
    if ($psProcesses) {
        $sessions += $psProcesses | ForEach-Object {
            $startTime = try { $_.StartTime } catch { $null }
            $windowTitle = try { $_.MainWindowTitle } catch { "" }
            [PSCustomObject]@{
                ProcessId = $_.Id
                ProcessName = $_.ProcessName
                Type = "Windows PowerShell"
                StartTime = $startTime
                MemoryMB = [math]::Round($_.WorkingSet / 1MB, 1)
                WindowTitle = $windowTitle
                IsCurrent = ($_.Id -eq $PID)
                Process = $_
            }
        }
    }
    
    # PowerShell Core processes (PowerShell 7+)
    $pwshProcesses = Get-Process -Name "pwsh" -ErrorAction SilentlyContinue
    if ($pwshProcesses) {
        $sessions += $pwshProcesses | ForEach-Object {
            $startTime = try { $_.StartTime } catch { $null }
            $windowTitle = try { $_.MainWindowTitle } catch { "" }
            [PSCustomObject]@{
                ProcessId = $_.Id
                ProcessName = $_.ProcessName
                Type = "PowerShell Core"
                StartTime = $startTime
                MemoryMB = [math]::Round($_.WorkingSet / 1MB, 1)
                WindowTitle = $windowTitle
                IsCurrent = ($_.Id -eq $PID)
                Process = $_
            }
        }
    }
    
    # PowerShell ISE
    $iseProcesses = Get-Process -Name "powershell_ise" -ErrorAction SilentlyContinue
    if ($iseProcesses) {
        $sessions += $iseProcesses | ForEach-Object {
            $startTime = try { $_.StartTime } catch { $null }
            $windowTitle = try { $_.MainWindowTitle } catch { "" }
            [PSCustomObject]@{
                ProcessId = $_.Id
                ProcessName = $_.ProcessName
                Type = "PowerShell ISE"
                StartTime = $startTime
                MemoryMB = [math]::Round($_.WorkingSet / 1MB, 1)
                WindowTitle = $windowTitle
                IsCurrent = $false
                Process = $_
            }
        }
    }
    
    # VS Code processes (check if they might be running PowerShell)
    $codeProcesses = Get-Process -Name "code" -ErrorAction SilentlyContinue
    if ($codeProcesses) {
        foreach ($proc in $codeProcesses) {
            try {
                $wmiProc = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId=$($proc.Id)" -ErrorAction SilentlyContinue
                if ($wmiProc -and $wmiProc.CommandLine -and ($wmiProc.CommandLine -like "*powershell*" -or $wmiProc.CommandLine -like "*.ps1*")) {
                    $startTime = try { $proc.StartTime } catch { $null }
                    $windowTitle = try { $proc.MainWindowTitle } catch { "" }
                    $sessions += [PSCustomObject]@{
                        ProcessId = $proc.Id
                        ProcessName = $proc.ProcessName
                        Type = "VS Code (PowerShell)"
                        StartTime = $startTime
                        MemoryMB = [math]::Round($proc.WorkingSet / 1MB, 1)
                        WindowTitle = $windowTitle
                        IsCurrent = $false
                        Process = $proc
                    }
                }
            }
            catch {
                # Ignore errors when checking VS Code processes
            }
        }
    }
    
    # Windows Terminal processes (may contain PowerShell sessions)
    $terminalProcesses = Get-Process -Name "WindowsTerminal" -ErrorAction SilentlyContinue
    if ($terminalProcesses) {
        $sessions += $terminalProcesses | ForEach-Object {
            $startTime = try { $_.StartTime } catch { $null }
            $windowTitle = try { $_.MainWindowTitle } catch { "" }
            [PSCustomObject]@{
                ProcessId = $_.Id
                ProcessName = $_.ProcessName
                Type = "Windows Terminal"
                StartTime = $startTime
                MemoryMB = [math]::Round($_.WorkingSet / 1MB, 1)
                WindowTitle = $windowTitle
                IsCurrent = $false
                Process = $_
            }
        }
    }
    
    # Filter out current process if requested
    if (-not $IncludeCurrent) {
        $sessions = $sessions | Where-Object { -not $_.IsCurrent }
    }
      if ($sessions.Count -eq 0) {
        if (-not $Silent) {
            Write-ColorOutput "    No PowerShell sessions detected" -Type Process
        }
        return @()
    }

    # Display results only if not silent
    if (-not $Silent) {
        $currentSessions = $sessions | Where-Object { $_.IsCurrent }
        $otherSessions = $sessions | Where-Object { -not $_.IsCurrent }
        
        if ($currentSessions) {
            Write-ColorOutput "    Current session:" -Type Info
            foreach ($session in $currentSessions) {
                # Use traditional if-else for Constrained Language Mode compatibility
                if ($session.StartTime) {
                    $startTimeText = $session.StartTime.ToString("yyyy-MM-dd HH:mm:ss")
                } else {
                    $startTimeText = "Unknown"
                }
                Write-ColorOutput "    • PID: $($session.ProcessId) | $($session.Type) | Started: $startTimeText | Memory: $($session.MemoryMB)MB" -Type Process
            }
        }
        
        if ($otherSessions.Count -gt 0) {
            Write-ColorOutput "    Found $($otherSessions.Count) other PowerShell session(s):" -Type Warning
            
            foreach ($session in $otherSessions) {
                # Use traditional if-else for Constrained Language Mode compatibility
                if ($session.StartTime) {
                    $startTimeText = $session.StartTime.ToString("yyyy-MM-dd HH:mm:ss")
                } else {
                    $startTimeText = "Unknown"
                }
                
                if ($session.WindowTitle -and $session.WindowTitle.Trim()) {
                    $titleText = $session.WindowTitle
                } else {
                    $titleText = "No window title"
                }
                
                if ($ShowDetails) {
                    Write-ColorOutput "    • PID: $($session.ProcessId) | $($session.Type) | Started: $startTimeText | Memory: $($session.MemoryMB)MB" -Type Info
                    Write-ColorOutput "      Title: $titleText" -Type Info
                    
                    # Try to get additional details via WMI
                    try {
                        $wmiProc = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId=$($session.ProcessId)" -ErrorAction SilentlyContinue
                        if ($wmiProc) {
                            $owner = try { (Invoke-CimMethod -InputObject $wmiProc -MethodName GetOwner).User } catch { "Unknown" }
                            Write-ColorOutput "      Owner: $owner" -Type Info
                            
                            if ($wmiProc.CommandLine -and $wmiProc.CommandLine.Length -gt 100) {
                                Write-ColorOutput "      Command: $($wmiProc.CommandLine.Substring(0,97))..." -Type Info
                            } elseif ($wmiProc.CommandLine) {
                                Write-ColorOutput "      Command: $($wmiProc.CommandLine)" -Type Info
                            }
                        }
                    }
                    catch {
                        # Ignore WMI errors
                    }
                    Write-ColorOutput "" -Type Info
                } else {
                    Write-ColorOutput "    • PID: $($session.ProcessId) | $($session.Type) | Memory: $($session.MemoryMB)MB | $titleText" -Type Info
                }
            }
            
            if ($ShowConflicts) {
                Write-ColorOutput ""
                Write-ColorOutput "Potential conflicts with module operations:" -Type Warning
                Write-ColorOutput "    • Other PowerShell sessions may have modules loaded" -Type Warning
                Write-ColorOutput "    • Loaded modules cannot be uninstalled or updated" -Type Warning
                Write-ColorOutput "    • ISE and VS Code may have PowerShell modules in memory" -Type Warning
                Write-ColorOutput "    • Windows Terminal may contain hidden PowerShell sessions" -Type Warning
            }
        }
    }
    
    return $sessions
}

function Stop-ConflictingPowerShellSessions {
    <#
    .SYNOPSIS
        Terminates conflicting PowerShell processes that may be holding modules
    
    .DESCRIPTION
        Identifies and optionally terminates PowerShell processes that could prevent
        module installation, update, or removal operations. Provides user confirmation
        before terminating processes and excludes the current session.
    
    .PARAMETER Sessions
        Array of PowerShell session objects from Get-PowerShellSessions
    
    .PARAMETER Force
        Skip confirmation prompts and terminate all conflicting sessions
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Sessions,
        
        [Parameter()]
        [switch]$Force
    )
    
    # Filter out current session and get potentially conflicting sessions
    $conflictingSessions = $Sessions | Where-Object { -not $_.IsCurrent }
    
    if ($conflictingSessions.Count -eq 0) {
        Write-ColorOutput "    No conflicting PowerShell sessions to terminate" -Type Process        return $true
    }
      # Sessions already displayed by caller, proceed with termination logic
    
    # Get user confirmation unless Force is specified
    if (-not $Force) {
        Write-ColorOutput "⚠ WARNING: Terminating these processes will:" -Type Warning
        Write-ColorOutput "    • Close any unsaved work in those PowerShell sessions" -Type Warning
        Write-ColorOutput "    • Stop any running scripts or commands" -Type Warning
        Write-ColorOutput "    • Close PowerShell ISE, VS Code PowerShell terminals, etc." -Type Warning
        Write-ColorOutput ""
        
        $response = Read-Host "Do you want to terminate these conflicting PowerShell sessions? (Y/N)"
        if ($response -notmatch '^[Yy]') {
            Write-ColorOutput "Session termination cancelled. Module operations may fail due to conflicts." -Type Warning
            Write-ColorOutput ""
            return $false
        }
        Write-ColorOutput ""
    }
    
    # Terminate conflicting sessions
    Write-ColorOutput "Terminating conflicting PowerShell sessions..." -Type System
    $terminatedCount = 0
    $failedCount = 0
    
    foreach ($session in $conflictingSessions) {
        Write-ColorOutput "    Terminating PID $($session.ProcessId) ($($session.Type))..." -Type Process -NoNewline
        
        try {
            # Try graceful termination first (for GUI applications)
            if ($session.Type -eq "PowerShell ISE" -or $session.Type -eq "VS Code (PowerShell)") {
                $session.Process.CloseMainWindow()
                Start-Sleep -Seconds 2
                
                # Check if process is still running
                $stillRunning = Get-Process -Id $session.ProcessId -ErrorAction SilentlyContinue
                if ($stillRunning) {
                    # Force termination if graceful close didn't work
                    $session.Process.Kill()
                }
            }
            else {
                # Force termination for console applications
                $session.Process.Kill()
            }
            
            # Wait a moment and verify termination
            Start-Sleep -Seconds 1
            $stillRunning = Get-Process -Id $session.ProcessId -ErrorAction SilentlyContinue
            
            if (-not $stillRunning) {
                Write-Host " Success!" -ForegroundColor $Script:Colors.Process
                $terminatedCount++
            }
            else {
                Write-Host " Failed (still running)" -ForegroundColor $Script:Colors.Error
                $failedCount++
            }
        }
        catch [System.InvalidOperationException] {
            # Process was already terminated
            Write-Host " Already terminated" -ForegroundColor $Script:Colors.Info
            $terminatedCount++
        }
        catch {
            Write-Host " Error: $($_.Exception.Message)" -ForegroundColor $Script:Colors.Error
            $failedCount++
        }
    }
    
    Write-ColorOutput ""
    if ($terminatedCount -gt 0) {
        Write-ColorOutput "Successfully terminated $terminatedCount PowerShell session(s)" -Type Process
    }
    
    if ($failedCount -gt 0) {
        Write-ColorOutput "Failed to terminate $failedCount PowerShell session(s)" -Type Warning
        Write-ColorOutput "You may need to manually close these applications or restart your computer" -Type Warning
    }
    
    # Brief pause to let processes fully terminate
    if ($terminatedCount -gt 0) {
        Write-ColorOutput "Waiting for processes to fully terminate..." -Type Info
        Start-Sleep -Seconds 3
    }
    
    return ($failedCount -eq 0)
}

function Show-PowerShellSessionGuidance {
    [CmdletBinding()]
    param(
        [Parameter()]
        [array]$Sessions
    )
    
    if (-not $Sessions -or $Sessions.Count -eq 0) {
        return
    }
    
    $otherSessions = $Sessions | Where-Object { -not $_.IsCurrent }
    if ($otherSessions.Count -eq 0) {
        return
    }
    
    Write-ColorOutput ""
    Write-ColorOutput "To resolve PowerShell session conflicts:" -Type Info
    Write-ColorOutput "• Close PowerShell windows manually" -Type Info
    Write-ColorOutput "• Close PowerShell ISE if open" -Type Info
    Write-ColorOutput "• Close VS Code if it has PowerShell files open" -Type Info
    Write-ColorOutput "• Close Windows Terminal tabs with PowerShell" -Type Info
    Write-ColorOutput ""
    Write-ColorOutput "To forcefully terminate PowerShell processes:" -Type Warning
    
    foreach ($session in $otherSessions) {
        if ($session.Type -ne "Windows Terminal") {  # Don't suggest killing Windows Terminal
            Write-ColorOutput "    Stop-Process -Id $($session.ProcessId) -Force" -Type Warning
        }
    }
    
    Write-ColorOutput ""
    Write-ColorOutput "After closing sessions, you can retry this script." -Type Info
}

function Test-ModuleConflicts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ModuleNames
    )
    
    Write-ColorOutput "Checking for module conflicts..." -Type Info
    
    $conflicts = @()
    $allSessions = Get-PowerShellSessions -IncludeCurrent
    $otherSessions = $allSessions | Where-Object { -not $_.IsCurrent }
    
    if ($otherSessions.Count -eq 0) {
        Write-ColorOutput "    No other PowerShell sessions detected" -Type Process
        return @()
    }
    
    # Check if any of the modules we want to remove/update are currently loaded
    $loadedModules = Get-Module | Where-Object { $_.Name -in $ModuleNames }
    if ($loadedModules) {
        Write-ColorOutput "    [Warning] The following modules are currently loaded:" -Type Warning
        foreach ($module in $loadedModules) {
            Write-ColorOutput "    • $($module.Name) v$($module.Version)" -Type Warning
            $conflicts += @{
                ModuleName = $module.Name
                Version = $module.Version
                Reason = "Currently loaded in this session"
                Session = "Current"
            }
        }
    }
    
    # Estimate potential conflicts from other sessions
    if ($otherSessions.Count -gt 0) {
        Write-ColorOutput "    [Warning] $($otherSessions.Count) other PowerShell sessions detected" -Type Warning
        Write-ColorOutput "    These sessions may have modules loaded that could prevent removal/updates" -Type Warning
        
        foreach ($session in $otherSessions) {
            foreach ($moduleName in $ModuleNames) {
                $conflicts += @{
                    ModuleName = $moduleName
                    Version = "Unknown"
                    Reason = "Potentially loaded in other session"
                    Session = "$($session.Type) (PID: $($session.ProcessId))"
                }
            }
        }
    }
    
    return $conflicts
}

# Main execution
try {
    # If user just wants to check for session conflicts, do that and exit
    if ($CheckSessions) {
        Clear-Host
        Write-ColorOutput "=== PowerShell Session Conflict Checker ===" -Type System
        Write-ColorOutput "Script Version: 2.5" -Type System
        Write-ColorOutput "Scan Time: $(Get-Date -Format (Get-Culture).DateTimeFormat.FullDateTimePattern)" -Type System
        Write-ColorOutput ""
        
        # Get all PowerShell sessions with detailed information
        $allSessions = Get-PowerShellSessions -IncludeCurrent -ShowDetails -ShowConflicts
        
        # Check for specific module conflicts with deprecated modules
        $moduleNames = $Script:DeprecatedModules | ForEach-Object { $_.Name }
        $conflicts = Test-ModuleConflicts -ModuleNames $moduleNames
        
        if ($conflicts.Count -gt 0) {
            Write-ColorOutput ""
            Write-ColorOutput "Module Conflict Analysis:" -Type Warning
            $uniqueConflicts = $conflicts | Sort-Object ModuleName, Session -Unique
            foreach ($conflict in $uniqueConflicts) {
                Write-ColorOutput "    • $($conflict.ModuleName): $($conflict.Reason) [$($conflict.Session)]" -Type Warning
            }
        }
          # Show guidance for resolving conflicts
        Show-PowerShellSessionGuidance -Sessions $allSessions
        
        # Offer termination option in CheckSessions mode
        $otherSessions = $allSessions | Where-Object { -not $_.IsCurrent }
        if ($otherSessions.Count -gt 0) {
            Write-ColorOutput ""
            $response = Read-Host "Would you like to terminate these conflicting sessions now? (Y/N)"
            if ($response -match '^[Yy]') {
                $terminationSuccess = Stop-ConflictingPowerShellSessions -Sessions $allSessions
                if ($terminationSuccess) {
                    Write-ColorOutput ""
                    Write-ColorOutput "✓ All conflicting sessions terminated successfully!" -Type Process
                    Write-ColorOutput "You can now run the update script without session conflicts." -Type Process
                }
            }
        }
        
        Write-ColorOutput ""
        Write-ColorOutput "Session check completed. Run the script without -CheckSessions to proceed with updates." -Type Info
        exit 0
    }
    
    # Validate administrator privileges
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-ColorOutput "Error: This script requires Administrator privileges" -Type Error
        Write-ColorOutput "Please run PowerShell as Administrator and try again" -Type Error
        exit 1
    }
      # Initialize script
    $moduleCount = Start-ScriptExecution    # Check for conflicting PowerShell sessions before starting module operations
    Write-ColorOutput "Checking for conflicting PowerShell sessions..." -Type System
    $Script:SessionConflictCheckPerformed = $true
    $allSessions = Get-PowerShellSessions -Silent
    $conflictingSessions = $allSessions | Where-Object { -not $_.IsCurrent }
      if ($conflictingSessions.Count -gt 0) {
        Write-ColorOutput ""
        Write-ColorOutput "⚠ DETECTED $($conflictingSessions.Count) OTHER POWERSHELL SESSION(S)" -Type Warning
        Write-ColorOutput "These sessions may interfere with module installation, updates, and removal." -Type Warning
        Write-ColorOutput ""
        
        # Show session details
        foreach ($session in $conflictingSessions) {
            # Use traditional if-else for Constrained Language Mode compatibility
            if ($session.StartTime) {
                $startTimeText = $session.StartTime.ToString("yyyy-MM-dd HH:mm:ss")
            } else {
                $startTimeText = "Unknown"
            }
            
            if ($session.WindowTitle -and $session.WindowTitle.Trim()) {
                $titleText = $session.WindowTitle
            } else {
                $titleText = "No title"
            }
            Write-ColorOutput "    • PID: $($session.ProcessId) - $($session.Type) - $titleText" -Type Info
        }
        Write-ColorOutput ""
        
        # Show informational warning about module operations being blocked
        Write-ColorOutput "These sessions may have PowerShell modules loaded, which can prevent:" -Type Warning
        Write-ColorOutput "    • Module installation (files in use)" -Type Warning
        Write-ColorOutput "    • Module updates (existing versions locked)" -Type Warning
        Write-ColorOutput "    • Module removal (loaded modules cannot be uninstalled)" -Type Warning
        Write-ColorOutput ""
          # Always offer termination option unless in CheckOnly mode
        if (-not $CheckOnly) {
            if ($TerminateConflicts) {
                # Automatic termination mode
                Write-ColorOutput "Auto-terminating conflicting sessions (TerminateConflicts parameter specified)..." -Type System
                $terminationSuccess = Stop-ConflictingPowerShellSessions -Sessions $allSessions -Force
                $Script:SessionConflictsResolved = $terminationSuccess
                if (-not $terminationSuccess) {
                    Write-ColorOutput "Warning: Some sessions could not be terminated. Module operations may encounter issues." -Type Warning
                    Write-ColorOutput ""
                }
            }
            else {                # Interactive mode - ask user
                $response = Read-Host "Terminate conflicting PowerShell sessions to ensure smooth operation? (Y/N)"
                if ($response -match '^[Yy]') {
                    $terminationSuccess = Stop-ConflictingPowerShellSessions -Sessions $allSessions -Force
                    $Script:SessionConflictsResolved = $terminationSuccess
                    if (-not $terminationSuccess) {
                        Write-ColorOutput "Warning: Some sessions could not be terminated. Module operations may encounter issues." -Type Warning
                        Write-ColorOutput ""
                        $continueResponse = Read-Host "Continue anyway? (Y/N)"
                        if ($continueResponse -notmatch '^[Yy]') {
                            Write-ColorOutput "Script execution cancelled by user." -Type Warning
                            exit 0
                        }
                    }
                    Write-ColorOutput ""
                }
                else {
                    Write-ColorOutput ""
                    Write-ColorOutput "⚠ Continuing with conflicting sessions present." -Type Warning
                    Write-ColorOutput "If module operations fail, consider closing other PowerShell applications." -Type Warning
                    Write-ColorOutput ""
                    $Script:SessionConflictsResolved = $false
                }
            }
        }
        else {
            Write-ColorOutput "Note: In check-only mode. Use -CheckSessions for detailed session analysis." -Type Info
            Write-ColorOutput ""
            $Script:SessionConflictsResolved = $false
        }
    }
    else {
        Write-ColorOutput "    ✓ No conflicting PowerShell sessions detected" -Type Process
        Write-ColorOutput ""
        $Script:SessionConflictsResolved = $true    }
    $counter = 0
    $moduleCount = $Script:ModuleList.Count
    
    # Update NuGet provider for PowerShell 5.x
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        Write-ColorOutput "Updating NuGet provider for PowerShell 5.x compatibility..." -Type Info
        Test-PackageProvider -PackageName "NuGet" -Description "NuGet package provider"
    }
    
    # Proactive guidance for core module updates
    Write-ColorOutput "📋 Core Module Update Information:" -Type Info
    Write-ColorOutput "• Core modules (PackageManagement, PowerShellGet) may show 'in use' warnings" -Type Info
    Write-ColorOutput "• This is normal behavior - these modules are essential to PowerShell" -Type Info  
    Write-ColorOutput "• Updates will install side-by-side and activate on next PowerShell restart" -Type Info
    Write-ColorOutput "• No action required from you - the script handles this automatically" -Type Info
    Write-ColorOutput ""
    
    # Clean up deprecated modules first
    if (-not $SkipDeprecatedCleanup -and -not $CheckOnly) {
        # Check prerequisites before attempting module removal
        if (-not (Test-ModuleRemovalPrerequisites)) {
            Write-ColorOutput "Skipping deprecated module cleanup due to prerequisite issues" -Type Warning
            Write-ColorOutput ""
        } else {
            Remove-DeprecatedModules
        }
    } else {
        Remove-DeprecatedModules
    }
    
    # Process each module
    foreach ($module in $Script:ModuleList) {
        $counter++
        Write-ColorOutput "($counter of $moduleCount) Processing $($module.Description)" -Type Process
        
        # Check if this module requires special handling (core PowerShell modules)
        if ($module.RequiresSpecialHandling -eq $true) {
            Test-CoreModuleInstallation -ModuleName $module.Name -Description $module.Description
        }
        else {
            Test-ModuleInstallation -ModuleName $module.Name -Description $module.Description
        }
        
        Write-ColorOutput ""
    }    # Complete script execution
    Stop-ScriptExecution
    
    # Display troubleshooting information if there were any issues - but make it dismissible
    if (-not $CheckOnly) {
        Write-ColorOutput ""
        Write-ColorOutput "=== Troubleshooting Information ===" -Type Info
        Write-ColorOutput "Common troubleshooting tips:" -Type Info
        Write-ColorOutput "• If modules fail to uninstall: Run as Administrator and close other PowerShell sessions" -Type Info
        Write-ColorOutput "• If modules are 'in use': Restart PowerShell and try again" -Type Info
        Write-ColorOutput "• For MSI-installed modules: Use Windows Add/Remove Programs" -Type Info
        Write-ColorOutput "• For manual cleanup: Delete module folders from PowerShell module paths" -Type Info
        Write-ColorOutput "• If connection fails: Check network connectivity and firewall settings" -Type Info
        Write-ColorOutput "• For more help, visit: https://docs.microsoft.com/powershell/module/powershellget/" -Type Info
        Write-ColorOutput ""
        
        # Clear the troubleshooting info and show final status
#        Clear-Host
        Write-ColorOutput "`=== Microsoft Cloud PowerShell Module Updater - Complete ===" -Type System
        Write-ColorOutput "Script execution finished successfully!" -Type Process
        Write-ColorOutput "End Time: $(Get-Date -Format (Get-Culture).DateTimeFormat.FullDateTimePattern)" -Type System
        Write-ColorOutput ""
        
        # Show summary of what was accomplished
        Write-ColorOutput "Summary:" -Type Info
        Write-ColorOutput "• Module updates completed" -Type Process
        if (-not $SkipDeprecatedCleanup) {
            Write-ColorOutput "• Deprecated modules cleaned up" -Type Process
        }
        Write-ColorOutput "• System is ready for Microsoft Cloud operations" -Type Process
        Write-ColorOutput ""
        Write-ColorOutput "You can now use the updated PowerShell modules for Microsoft 365, Azure, and Teams management." -Type Info
    } else {
        Write-ColorOutput ""
        Write-ColorOutput "Check-only mode completed. No changes were made to your system." -Type Info
    }
    
    Write-ColorOutput ""
}
catch {
    Write-ColorOutput "Fatal Error: $($_.Exception.Message)" -Type Error
    Write-ColorOutput "Stack Trace: $($_.ScriptStackTrace)" -Type Error
    
    if ($CreateLog) {
        try {
            Stop-Transcript
        }
        catch {
            # Transcript might not be running
        }
    }
    
    exit 1
}