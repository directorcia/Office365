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

.PARAMETER ForceVersionCleanup
    Forces a comprehensive cleanup of ALL old module versions across ALL installed modules
    (not just those in the current -ModuleScope), regardless of -SkipVersionCleanup or -Prompt.
    Modules are removed from the current session first so Uninstall-Module isn't blocked by
    "module in use". If Uninstall-Module still cannot remove a version (e.g. it was installed
    outside PowerShellGet by copying files into PSModulePath), the module folder is deleted
    directly from disk as a fallback. Bundled PS7 modules under $PSHOME\Modules are never touched.
    PowerShellGet and PackageManagement are cleaned up per module root so each root keeps its
    newest copy, preserving Windows PowerShell 5.1 bootstrap capability under WindowsPowerShell\Modules.

.PARAMETER CheckOnly
    Only checks versions without performing updates or installations

.PARAMETER CheckSessions
    Checks for PowerShell session conflicts and exits without performing updates

.PARAMETER TerminateConflicts
    Retained for backwards compatibility only. Terminating other sessions is not implemented;
    the switch is accepted but ignored with a warning. Use -CheckSessions to list sessions
    that may lock modules and close them manually.

.PARAMETER ModuleScope
    Limits processing to one or more module categories: All (default), Graph, Azure,
    Exchange, Teams, SharePoint, PowerApps. Core and Other modules (PowerShellGet,
    PackageManagement, Microsoft.WinGet.Client) are only processed with 'All'.

.PARAMETER SkipConnectivityCheck
    Skips the repository reachability check that normally runs before any updates.
    Use when outbound web requests are blocked but the repository is still usable
    (e.g. an internal file-share repository).

.PARAMETER Repository
    Name of a registered PowerShellGet repository to install from (default: PSGallery).
    Must already be registered with Register-PSRepository.

.PARAMETER TimeoutMinutes
    Maximum minutes allowed for each installation before its worker process is stopped
    (1-240, default 30). This limit is enforced in both background-job and child-process modes.
    Old versions are retained when an installation times out.

.PARAMETER PassThru
    Emits one result object per processed module to the pipeline so results can be
    filtered, exported or inspected by automation. Console output is unchanged.

.PARAMETER WhatIf
    Shows which modules would be installed, updated or uninstalled without making changes.
    Administrator rights are not required when -WhatIf is used.

.PARAMETER Confirm
    Prompts for confirmation before each install, update or uninstall.

.INPUTS
    None. This script does not accept pipeline input.

.OUTPUTS
    None by default (progress is written to the host/information stream).
    With -PassThru: PSCustomObject per module with ModuleName, Category, Action, Status,
    LocalVersion, OnlineVersion, Duration and ErrorMessage properties.

.EXAMPLE
    .\o365-update.ps1 -ModuleScope Graph, Exchange -WhatIf
    Shows what would be installed/updated for the Graph and Exchange modules without changing anything.

.EXAMPLE
    $results = .\o365-update.ps1 -CheckOnly -PassThru
    $results | Where-Object Status -eq 'UpdateAvailable' | Format-Table ModuleName, LocalVersion, OnlineVersion
    Captures check results as objects and lists modules with pending updates.

.EXAMPLE
    .\o365-update.ps1 -CreateLog -LogPath 'D:\Logs\Module Updates' -TimeoutMinutes 60
    Unattended run (e.g. scheduled task) with a transcript in a folder containing spaces and a longer timeout.
    Check $LASTEXITCODE: 0 = success, 1 = one or more operations failed or prerequisites were not met.

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
    .\o365-update.ps1 -SkipVersionCleanup
    Updates modules but skips automatic cleanup of old versions to preserve multiple versions

.EXAMPLE
    .\o365-update.ps1 -Prompt -SkipVersionCleanup
    Interactive mode with cleanup prompts. When multiple versions are detected,
    prompts user to choose whether to clean up old versions for each module.
    Also offers comprehensive cleanup at the end of processing.

.EXAMPLE
    .\o365-update.ps1 -ForceVersionCleanup
    Updates modules, then performs a comprehensive cleanup of every old version for
    every installed module (not just the ones being updated), including versions that
    Uninstall-Module normally can't remove because they're in use or were installed
    outside PowerShellGet.

.NOTES
    Author: CIAOPS    Version: 2.18
    Last Updated: November 2026
    Requires: PowerShell 7.X; Administrator privileges for installs/removals (not for
              -CheckOnly, -CheckSessions or -WhatIf); PowerShellGet 2.x; internet access
              to the PowerShell Gallery (or the repository given in -Repository)

    IMPORTANT: This version removes deprecated Azure AD and MSOnline modules in favor of Microsoft Graph PowerShell SDK

    Major Changes in v2.18:
    - Verifies effective write access to the AllUsers module directory before making changes
    - Enforces TimeoutMinutes in restricted environments by running installs in a child process
    - Reports failed final version cleanups separately and includes them in the failure exit code

    Major Changes in v2.17:
    - Added -WhatIf/-Confirm support for installs, updates, old-version cleanup and deprecated
      module removal; -WhatIf runs without elevation
    - Added -PassThru to emit one O365Update.ModuleResult object per module for automation
    - Set-StrictMode -Version Latest; fixed single-module -ModuleScope reporting the wrong count
    - Prerelease module versions (e.g. 1.2.3-nightly) no longer break version comparisons
    - TLS 1.2 is now added to (not substituted for) the enabled protocols, so TLS 1.3 stays on
    - The repository InstallationPolicy is restored after install instead of being left Trusted
    - Uses cached, name-filtered presence checks to avoid Get-InstalledModule calls for modules
      not present on disk, without performing a slow full module-path scan
    - Consolidated duplicated direct-install fallback code; removed unused functions
    - Fixed comment-based help so Get-Help works

    Major Changes in v2.16 (19 November 2026):
    - Made prerequisite failures and unreachable repositories fail fast instead of continuing
      into operations that cannot succeed
    - Prevented -CheckOnly from performing deprecated-module cleanup
    - Applied package-provider failures and forced-cleanup failures to the final exit code
    - Fixed AllowClobber conflict retries for update operations and detection of blocked jobs
    - Added clearer read-only mode reporting for administrator and session checks

    Major Changes in v2.15 (8 September 2026):
    - Added -ForceVersionCleanup switch for a comprehensive, unconditional cleanup of every
      old version of every installed module (not limited to -ModuleScope), overriding
      -SkipVersionCleanup and skipping -Prompt confirmations
    - New Remove-ModuleVersionForced helper unloads the target version from the current
      session (Remove-Module) before calling Uninstall-Module, avoiding "module is in use"
      failures for versions loaded earlier in the same session
    - Added a direct filesystem fallback: if Uninstall-Module still cannot remove a version
      (e.g. copied into PSModulePath outside PowerShellGet), the module version folder is
      removed directly, after re-confirming the folder isn't the active bundled PS7 module
    - Remove-AllOldModuleVersions and Remove-OldModuleVersions now also inspect
      Get-Module -ListAvailable -AllVersions so unregistered/manually-installed
      duplicate versions are detected and removed, not just PowerShellGet-registered ones
    - PowerShellGet and PackageManagement cleanup now keeps the newest version in each
      module root instead of keeping only the newest version globally, so Windows
      PowerShell 5.1 keeps a bootstrap copy even when PS7 has its bundled fallback

    Major Changes in v2.14 (31 July 2026):
    - Fixed banner showing v2.9 regardless of the actual script version
    - Fixed call to non-existent Update-PackageProvider cmdlet (now uses Install-PackageProvider)
    - Wired NuGet package provider check into main execution (Test-PackageProvider was never called)
    - Completed the broken administrator check in Test-Prerequisites (dead code produced no output)
    - Implemented -CheckSessions to list other PowerShell sessions and moved it before the
      network connectivity check; -TerminateConflicts now warns that it is ignored
    - Removed cross-process function calls from the background job script block
      (Remove-OldModuleVersions is not defined inside Start-Job child processes;
      cleanup is handled by the parent after job completion)
    - Fixed progress-file name mismatch between the background job and the monitor loop
      (job used its own $PID while the monitor looked for the job Id)
    - Renamed transcript log variable to avoid clobbering the -LogPath parameter
      (PowerShell variable names are case-insensitive)

    Major Changes in v2.13 (5 May 2026):
    - Added PS7 bundled-module detection to Get-ModuleVersionInfo
      Get-InstalledModule only sees modules registered by PowerShellGet; modules that
      ship inside the PS7 installation directory (e.g. PowerShellGet, PackageManagement)
      are never registered there and were incorrectly reported as 'not installed'.
      A second check via Get-Module -ListAvailable now detects these bundled modules by
      confirming their ModuleBase path falls under $PSHOME\Modules.
    - Added IsBundled flag to the version-info hashtable returned by Get-ModuleVersionInfo
      so that downstream logic can distinguish bundled modules from uninstalled ones.
    - Added skip-install guard in Start-ModuleProcessing
      When a module is flagged as bundled and the resolved action is 'Install', the
      script now skips the install entirely and reports the bundled version as current.
      This eliminates the misleading 'version currently in use' WARNING messages that
      appeared when attempting to install PowerShellGet / PackageManagement while those
      same modules were already loaded from $PSHOME\Modules by the PS7 runtime.
    - Fixed LogPath parameter ignored by implementation
      The -LogPath parameter (default: current directory) was documented and defaulted
      correctly but the log file creation code hardcoded $env:TEMP, meaning the parameter
      had no effect and logs always appeared in the system temp folder regardless of what
      the user specified. The implementation now uses $LogPath so logs are written to the
      documented location.
    - Restored ValidateScript on LogPath parameter
      The path validator was lost when the above fix was applied. It has been restored
      and the error message improved to include the supplied path value so users
      immediately see which path was rejected if an invalid directory is provided.

    Major Changes in v2.12:
    - Note: #Requires directive still shows system error in PS5 (unavoidable)

    Major Changes in v2.11:
    - Note: #Requires directive still shows system error in PS5 (unavoidable)

    Major Changes in v2.10:
    - Enforced PowerShell 7+ only: added `#Requires -Version 7.0` at top
    - Removed `#Requires -RunAsAdministrator` to allow friendly guidance
    - Added early, user-friendly admin check with clear steps and rationale
    - Printed ready-to-copy elevation command using safe array-form `-ArgumentList`
    - Tightened prerequisite check to require PS 7+ (was 5.1+)
    - Improved non-admin experience: actionable instructions instead of parser errors
    - Documentation updated to reflect PS7-only support and elevation workflow

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

#Requires -Version 7.0

#region Parameters
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(HelpMessage = "Prompt before installing missing modules")]
    [switch]$Prompt,

    [Parameter(HelpMessage = "Create a transcript log of all operations")]
    [switch]$CreateLog,

    [Parameter(HelpMessage = "Path for log file (default: current directory)")]
    [ValidateScript({
        # -LiteralPath so folder names containing [ ] are not treated as wildcards
        if (-not (Test-Path -LiteralPath $_ -PathType Container)) {
            throw "LogPath '$_' does not exist or is not a directory. Create it first or pass an existing folder."
        }
        return $true
    })]
    [string]$LogPath = (Get-Location).Path,

    [Parameter(HelpMessage = "Skip deprecated module cleanup")]
    [switch]$SkipDeprecatedCleanup,

    [Parameter(HelpMessage = "Skip automatic cleanup of old module versions after updates")]
    [switch]$SkipVersionCleanup,

    [Parameter(HelpMessage = "Force a comprehensive cleanup of ALL old module versions for ALL installed modules, regardless of ModuleScope, -Prompt or -SkipVersionCleanup, unloading modules from the session and falling back to direct folder removal when Uninstall-Module cannot complete the job")]
    [switch]$ForceVersionCleanup,

    [Parameter(HelpMessage = "Only check versions without updating")]
    [switch]$CheckOnly,

    [Parameter(HelpMessage = "Check for PowerShell session conflicts and exit")]
    [switch]$CheckSessions,

    [Parameter(HelpMessage = "Accepted for compatibility; session termination is not implemented and this switch is ignored")]
    [switch]$TerminateConflicts,

    [Parameter(HelpMessage = "Specify which modules to process (default: all)")]
    [ValidateSet('All', 'Graph', 'Azure', 'Exchange', 'Teams', 'SharePoint', 'PowerApps')]
    [string[]]$ModuleScope = @('All'),

    [Parameter(HelpMessage = "Skip network connectivity checks")]
    [switch]$SkipConnectivityCheck,

    [Parameter(HelpMessage = "Use alternative PowerShell Gallery repository")]
    [ValidateNotNullOrEmpty()]
    [string]$Repository = 'PSGallery',

    [Parameter(HelpMessage = "Maximum minutes per installation worker process")]
    [ValidateRange(1, 240)]
    [int]$TimeoutMinutes = 30,

    [Parameter(HelpMessage = "Emit a result object per module to the pipeline")]
    [switch]$PassThru
)
#endregion Parameters

#region Configuration
Set-StrictMode -Version Latest
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

# Run state (initialised up front so Set-StrictMode never trips on an unset variable)
$Script:SkipCleanupPrompts = $false
$Script:LastInstallationError = $null
$Script:ModuleOperationResults = @()
$Script:FilteredModuleList = @()
$Script:ProcessingFailureCount = 0
$Script:SkippedCount = 0
$Script:ModuleResults = [System.Collections.Generic.List[object]]::new()
# Per-name results; avoids the much slower full Get-Module -ListAvailable scan.
$Script:ModulePresenceCache = @{}

$Script:JobsSupported = $ExecutionContext.SessionState.LanguageMode -eq 'FullLanguage'
$Script:BootstrapModuleNames = @('PowerShellGet', 'PackageManagement')
#endregion Configuration

#region Helper functions

function ConvertTo-ModuleVersion {
    <#
    .SYNOPSIS
        Converts a module version (possibly a prerelease string) to [version] for comparison.
    .DESCRIPTION
        PowerShellGet reports prerelease modules as strings such as '2.4.0-nightly', which a
        plain [version] cast rejects. The prerelease label is stripped so those modules can
        still be compared and sorted instead of failing the whole module check.
    #>
    [CmdletBinding()]
    [OutputType([version])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$Version
    )

    if ($null -eq $Version) { return [version]'0.0' }
    if ($Version -is [version]) { return $Version }

    $text = ([string]$Version).Trim()
    # Drop SemVer prerelease/build metadata: '1.2.3-beta+abc' -> '1.2.3'
    $text = ($text -split '[-+]', 2)[0]
    $parsed = $null
    if ([version]::TryParse($text, [ref]$parsed)) { return $parsed }
    # Single-component versions ('5') are valid module versions but not valid [version] strings
    if ($text -match '^\d+$') { return [version]::new([int]$text, 0) }

    throw "Cannot interpret '$Version' as a module version."
}

function Enable-Tls12 {
    <#
    .SYNOPSIS
        Ensures TLS 1.2 is enabled for this process without disabling newer protocols.
    .DESCRIPTION
        Assigning Tls12 outright (the previous behaviour) replaced SystemDefault and so turned
        off TLS 1.3. OR-ing the flag in keeps whatever the OS already allows.
    #>
    [CmdletBinding()]
    param()

    try {
        $current = [Net.ServicePointManager]::SecurityProtocol
        if ($current -ne [Net.SecurityProtocolType]::SystemDefault -and
            -not ($current -band [Net.SecurityProtocolType]::Tls12)) {
            [Net.ServicePointManager]::SecurityProtocol = $current -bor [Net.SecurityProtocolType]::Tls12
        }
        [Net.ServicePointManager]::DefaultConnectionLimit = [Math]::Max([Net.ServicePointManager]::DefaultConnectionLimit, 12)
    }
    catch {
        # Expected under Constrained Language Mode; downloads still work with OS defaults.
        Write-Verbose "Could not adjust network settings: $($_.Exception.Message)"
    }
}

function Test-ModulePresentOnDisk {
    <#
    .SYNOPSIS
        Checks whether a module exists on PSModulePath, caching by module name.
    .DESCRIPTION
        A name-filtered Get-Module query is much cheaper than either Get-InstalledModule -AllVersions
        or a full unfiltered Get-Module scan on module trees with many OneDrive-backed entries.
        Call with -Refresh after changing this module's installation state.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter()]
        [switch]$Refresh
    )

    if ($Refresh -or -not $Script:ModulePresenceCache.ContainsKey($Name)) {
        Write-Verbose "Checking PSModulePath for '$Name'..."
        $availableModules = @(Get-Module -Name $Name -ListAvailable -ErrorAction SilentlyContinue)
        $Script:ModulePresenceCache[$Name] = ($availableModules.Count -gt 0)
    }

    return [bool]$Script:ModulePresenceCache[$Name]
}

function Add-ModuleResult {
    <#
    .SYNOPSIS
        Records a per-module outcome for the -PassThru output and the final summary.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable]$Module,
        [Parameter(Mandatory)] [string]$Status,
        [Parameter()] [AllowNull()] [string]$Action,
        [Parameter()] [AllowNull()] [object]$LocalVersion,
        [Parameter()] [AllowNull()] [object]$OnlineVersion,
        [Parameter()] [timespan]$Duration = [timespan]::Zero,
        [Parameter()] [AllowNull()] [string]$ErrorMessage
    )

    $Script:ModuleResults.Add([pscustomobject]@{
        PSTypeName    = 'O365Update.ModuleResult'
        ModuleName    = $Module.Name
        Category      = $Module.Category
        Action        = $Action
        Status        = $Status
        LocalVersion  = if ($null -ne $LocalVersion) { [string]$LocalVersion } else { $null }
        OnlineVersion = if ($null -ne $OnlineVersion) { [string]$OnlineVersion } else { $null }
        Duration      = $Duration
        ErrorMessage  = $ErrorMessage
    })
}

function Write-ColorOutput {
    <#
    .SYNOPSIS
        Writes coloured console status text.
    .NOTES
        Write-Host is deliberate: this is UI text, not data. In PS5+ it goes to the
        Information stream (6), so it can still be captured or silenced with 6>$null.
        Use -PassThru on the script for machine-readable output.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Coloured UI text; data is emitted via -PassThru')]
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
    <#
    .SYNOPSIS
        Ensures a package provider (normally NuGet) is installed and current.
    .OUTPUTS
        [bool] $true when the provider is usable (or would be, under -WhatIf); otherwise $false.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PackageName
    )

    # NuGet versions below this cannot talk to the current PowerShell Gallery API
    $minimumProviderVersion = [version]'2.8.5.201'

    try {
        Write-ColorOutput "    Checking package provider: $PackageName" -Type Info

        # Several provider versions can be loaded side by side; compare against the newest
        $installedProvider = @(Get-PackageProvider -Name $PackageName -ErrorAction SilentlyContinue) |
            Sort-Object { [version]$_.Version } -Descending | Select-Object -First 1

        if (-not $installedProvider) {
            Write-ColorOutput "    [Warning] Package provider '$PackageName' not found" -Type Warning

            if ($Prompt -and -not $CheckOnly) {
                $response = Read-Host "    Install package provider '$PackageName' (Y/N)?"
                if ($response -notmatch '^[Yy]') {
                    Write-ColorOutput "    Skipping installation of $PackageName - module installs need it. Re-run and answer Y, or run: Install-PackageProvider -Name $PackageName -Force" -Type Warning
                    return $false
                }
            }

            if (-not $CheckOnly) {
                if (-not $PSCmdlet.ShouldProcess($PackageName, 'Install package provider')) {
                    return [bool]$WhatIfPreference
                }
                Write-ColorOutput "    Installing package provider: $PackageName" -Type Process
                Install-PackageProvider -Name $PackageName -Force -Confirm:$false
                Write-ColorOutput "    Successfully installed $PackageName" -Type Process
            }
            return $true
        }

        if ([version]$installedProvider.Version -ge $minimumProviderVersion) {
            Write-Verbose "Package provider '$PackageName' v$($installedProvider.Version) meets the minimum version $minimumProviderVersion."
            return $true
        }

        # Check for updates
        $onlineProvider = Find-PackageProvider -Name $PackageName -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $onlineProvider) {
            Write-ColorOutput "    Cannot find online version of $PackageName" -Type Warning
            return $true
        }

        $localVersion = $installedProvider.Version
        $onlineVersion = $onlineProvider.Version

        if ([version]$localVersion -ge [version]$onlineVersion) {
            Write-ColorOutput "    Local provider $PackageName ($localVersion) is up to date" -Type Process
            return $true
        }

        Write-ColorOutput "    Local provider $PackageName ($localVersion) can be updated to ($onlineVersion)" -Type Warning
        if (-not $CheckOnly -and $PSCmdlet.ShouldProcess("$PackageName $localVersion -> $onlineVersion", 'Update package provider')) {
            Write-ColorOutput "    Updating package provider: $PackageName" -Type Process
            # There is no Update-PackageProvider cmdlet; Install-PackageProvider -Force installs the newer version side-by-side
            Install-PackageProvider -Name $PackageName -Force -Confirm:$false -WarningAction SilentlyContinue 2>$null
            Write-ColorOutput "    Successfully updated $PackageName" -Type Process
        }
        return $true
    }
    catch {
        Write-ColorOutput "    Error processing package provider '$PackageName' - $($PSItem.Exception.Message)" -Type Error
        Write-ColorOutput "    Fix: run 'Install-PackageProvider -Name $PackageName -Force' in an elevated PowerShell 7 session, then re-run." -Type Info
        return $false
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
    )

    # Modules that don't support -AllowClobber
    $noAllowClobberModules = @(
        'Microsoft.WinGet.Client',
        'Az'
        # REMOVED: Microsoft.Graph and Microsoft.Graph.Authentication
        # These modules NEED AllowClobber to resolve command conflicts
    )

    # Modules that don't support -AcceptLicense
    $noAcceptLicenseModules = @(
        'Microsoft.PowerApps.Administration.PowerShell',
        'Microsoft.PowerApps.PowerShell',
        'Microsoft.WinGet.Client',
        'MicrosoftTeams',
        'PnP.PowerShell',
        'Microsoft.Online.SharePoint.PowerShell'
        # REMOVED: Microsoft.Graph - it supports AcceptLicense
    )

    # Modules that don't support -SkipPublisherCheck
    $noSkipPublisherCheckModules = @(
        'Microsoft.PowerApps.Administration.PowerShell',
        'Microsoft.PowerApps.PowerShell',
        'Microsoft.WinGet.Client',
        'ExchangeOnlineManagement',
        'Az',
        'MicrosoftTeams',
        'Microsoft.Online.SharePoint.PowerShell'
        # REMOVED: Microsoft.Graph - it supports SkipPublisherCheck
        # PnP.PowerShell is allowed to use SkipPublisherCheck to avoid publisher-signature conflicts during updates
    )

    # Note: Microsoft.PowerApps modules were removed from the exclusion lists above
    # because they actually REQUIRE -AllowClobber, -AcceptLicense, and -SkipPublisherCheck
    # to install properly due to command name conflicts

    # Start with base parameters
    $moduleParams = $BaseParams.Clone()

    # Remove Install-only parameters for Update-Module operations
    # Update-Module doesn't support: AllowClobber, AcceptLicense, SkipPublisherCheck
    # PnP.PowerShell is handled through Install-Module instead because it needs SkipPublisherCheck
    # to update cleanly when the existing package is signed by Microsoft.
    if ($Operation -eq 'Update' -and $ModuleName -ne 'PnP.PowerShell') {
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

function Test-ModuleConflictError {
    <#
    .SYNOPSIS
        Returns $true when an Install-Module error can be fixed by retrying with -AllowClobber
        (or -SkipPublisherCheck for PnP.PowerShell, whose publisher certificate changed).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)] [string]$ModuleName,
        [Parameter(Mandatory)] [AllowEmptyString()] [string]$ErrorMessage
    )

    if ($ErrorMessage -like '*already available*' -or
        $ErrorMessage -like '*AllowClobber*' -or
        $ErrorMessage -like '*may override the existing commands*') {
        return $true
    }
    return ($ModuleName -eq 'PnP.PowerShell' -and
        ($ErrorMessage -like '*publisher*' -or $ErrorMessage -like '*SkipPublisherCheck*' -or $ErrorMessage -like '*signed module*'))
}

function Invoke-PowerShellCommandWithTimeout {
    <#
    .SYNOPSIS
        Executes PowerShell command text in a child process with a hard timeout.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$CommandText,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$RequestJson = '{}',

        [Parameter(Mandatory)]
        [ValidateRange(1, 14400)]
        [int]$TimeoutSeconds,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Activity
    )

    $workerDirectory = Join-Path $env:TEMP "O365Update-$([guid]::NewGuid().ToString('N'))"
    $workerScriptPath = Join-Path $workerDirectory 'worker.ps1'
    $requestPath = Join-Path $workerDirectory 'request.json'
    $stdoutPath = Join-Path $workerDirectory 'stdout.txt'
    $stderrPath = Join-Path $workerDirectory 'stderr.txt'
    $process = $null
    $result = [ordered]@{
        Success = $false
        TimedOut = $false
        ExitCode = $null
        StandardOutput = $null
        StandardError = $null
        ErrorMessage = $null
        Duration = [timespan]::Zero
    }
    $startedAt = Get-Date

    try {
        $null = New-Item -Path $workerDirectory -ItemType Directory -ErrorAction Stop
        Set-Content -LiteralPath $workerScriptPath -Value $CommandText -Encoding utf8 -ErrorAction Stop
        Set-Content -LiteralPath $requestPath -Value $RequestJson -Encoding utf8 -ErrorAction Stop
        $pwshPath = Join-Path $PSHOME 'pwsh.exe'
        $argumentList = "-NoLogo -NoProfile -NonInteractive -File `"$workerScriptPath`" -RequestPath `"$requestPath`""
        $process = Start-Process -FilePath $pwshPath `
            -ArgumentList $argumentList `
            -PassThru -WindowStyle Hidden `
            -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath `
            -ErrorAction Stop

        $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
        while (-not $process.HasExited) {
            if ((Get-Date) -ge $deadline) {
                Stop-Process -Id $process.Id -Force -ErrorAction Stop
                $result.TimedOut = $true
                $result.ErrorMessage = "The PowerShell worker exceeded the $TimeoutSeconds-second timeout and was terminated."
                break
            }

            $elapsed = (Get-Date) - $startedAt
            $percentComplete = [int][Math]::Min(99, ($elapsed.TotalSeconds / $TimeoutSeconds) * 100)
            Write-Progress -Activity $Activity -Status ("Child process running ({0:N0}s elapsed)" -f $elapsed.TotalSeconds) -PercentComplete $percentComplete
            Start-Sleep -Seconds 1
        }

        if (-not $result.TimedOut) {
            $result.ExitCode = $process.ExitCode
            $result.Success = $result.ExitCode -eq 0
        }

        if (Test-Path -LiteralPath $stdoutPath) {
            $result.StandardOutput = Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction Stop
        }
        if (Test-Path -LiteralPath $stderrPath) {
            $result.StandardError = Get-Content -LiteralPath $stderrPath -Raw -ErrorAction Stop
        }

        if (-not $result.Success -and -not $result.TimedOut) {
            $result.ErrorMessage = if (-not [string]::IsNullOrWhiteSpace($result.StandardError)) {
                $result.StandardError.Trim()
            } elseif (-not [string]::IsNullOrWhiteSpace($result.StandardOutput)) {
                $result.StandardOutput.Trim()
            } else {
                "The PowerShell worker exited with code $($result.ExitCode) without an error message."
            }
        }
    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
    }
    finally {
        $result.Duration = (Get-Date) - $startedAt
        Write-Progress -Activity $Activity -Completed

        if ($process -and -not $process.HasExited) {
            try {
                Stop-Process -Id $process.Id -Force -ErrorAction Stop
            }
            catch {
                Write-Verbose "Could not stop PowerShell worker PID $($process.Id): $($_.Exception.Message)"
            }
        }

        if (Test-Path -LiteralPath $workerDirectory) {
            try {
                Remove-Item -LiteralPath $workerDirectory -Recurse -Force -ErrorAction Stop
            }
            catch {
                Write-Verbose "Could not remove temporary worker directory '$workerDirectory': $($_.Exception.Message)"
            }
        }
    }

    return [pscustomobject]$result
}

function Get-ModuleInstallWorkerScript {
    <#
    .SYNOPSIS
        Creates a worker command that installs a module from a serialized request.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    return @"
param([Parameter(Mandatory)][string]`$RequestPath)
`$ErrorActionPreference = 'Stop'
try {
    try {
        `$protocol = [Net.ServicePointManager]::SecurityProtocol
        if (`$protocol -ne [Net.SecurityProtocolType]::SystemDefault -and -not (`$protocol -band [Net.SecurityProtocolType]::Tls12)) {
            [Net.ServicePointManager]::SecurityProtocol = `$protocol -bor [Net.SecurityProtocolType]::Tls12
        }
    }
    catch {
        Write-Warning "Could not adjust TLS settings in the worker process: `$(`$_.Exception.Message)"
    }

    `$request = Get-Content -LiteralPath `$RequestPath -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -ErrorAction Stop
    `$workerParams = `$request.InstallParams
    Install-Module -Name `$request.ModuleName @workerParams -ErrorAction Stop
    exit 0
}
catch {
    Write-Error -Message `$_.Exception.Message -ErrorAction Continue
    exit 1
}
"@
}

function Invoke-ModuleInstallDirect {
    <#
    .SYNOPSIS
        Runs Install-Module in a child process when background jobs are unavailable.
    .DESCRIPTION
        Uses a separate PowerShell process so the installation can be terminated when its
        timeout expires. Retries once with -AllowClobber / -SkipPublisherCheck for conflicts.
    .OUTPUTS
        [bool] success. On failure $Script:LastInstallationError holds the reason.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)] [string]$ModuleName,
        [Parameter(Mandatory)] [hashtable]$InstallParams,
        [Parameter(Mandatory)] [string]$Operation,
        [Parameter(Mandatory)] [datetime]$StartTime,
        [Parameter(Mandatory)] [ValidateRange(1, 240)] [int]$TimeoutMinutes
    )

    $isConstrainedLargeModule = ($ExecutionContext.SessionState.LanguageMode -eq 'ConstrainedLanguage' -and
        $ModuleName -in @('Az', 'Microsoft.Graph'))

    Enable-Tls12
    Write-ColorOutput "    Starting $Operation of $ModuleName..." -Type Process
    if ($isConstrainedLargeModule) {
        Write-ColorOutput "    📋 Phase 1: Download and Installation starting..." -Type Process
    }

    $workerCommand = Get-ModuleInstallWorkerScript
    $requestJson = @{
        ModuleName = $ModuleName
        InstallParams = $InstallParams
    } | ConvertTo-Json -Depth 10 -Compress

    $attempt = Invoke-PowerShellCommandWithTimeout -CommandText $workerCommand `
        -RequestJson $requestJson `
        -TimeoutSeconds ($TimeoutMinutes * 60) `
        -Activity "$Operation module: $ModuleName"

    if ($attempt.Success) {
        $elapsed = (Get-Date) - $StartTime
        Write-ColorOutput "    Successfully completed $Operation of $ModuleName" -Type Process
        Write-ColorOutput ("    Actual time: {0:N1} minutes" -f $elapsed.TotalMinutes) -Type Info
        if ($isConstrainedLargeModule) {
            Write-ColorOutput "    ✅ $ModuleName processing complete in Constrained Language Mode" -Type Process
            Write-ColorOutput "    The extended time was due to security verification requirements" -Type Info
        }
        return $true
    }
    $errorMessage = $attempt.ErrorMessage
    $Script:LastInstallationError = $errorMessage

    if ($attempt.TimedOut -or -not (Test-ModuleConflictError -ModuleName $ModuleName -ErrorMessage $errorMessage)) {
        Write-ColorOutput "    Error during $Operation of $ModuleName`: $errorMessage" -Type Error
        Write-ColorOutput ("    Time elapsed: {0:N1} minutes" -f ((Get-Date) - $StartTime).TotalMinutes) -Type Info
        return $false
    }

    Write-ColorOutput "    ⚠️ Command conflict detected for $ModuleName" -Type Warning
    Write-ColorOutput "    🔄 Retry attempt with -AllowClobber..." -Type Process
    $retryParams = $InstallParams.Clone()
    $retryParams.AllowClobber = $true
    if ($ModuleName -eq 'PnP.PowerShell') {
        $retryParams.SkipPublisherCheck = $true
    }
    $workerCommand = Get-ModuleInstallWorkerScript
    $retryRequestJson = @{
        ModuleName = $ModuleName
        InstallParams = $retryParams
    } | ConvertTo-Json -Depth 10 -Compress
    $attempt = Invoke-PowerShellCommandWithTimeout -CommandText $workerCommand `
        -RequestJson $retryRequestJson `
        -TimeoutSeconds ($TimeoutMinutes * 60) `
        -Activity "$Operation module: $ModuleName (retry)"

    if ($attempt.Success) {
        Write-ColorOutput "    ✅ Successfully resolved conflict and installed $ModuleName" -Type Process
        Write-ColorOutput ("    Total time: {0:N1} minutes" -f ((Get-Date) - $StartTime).TotalMinutes) -Type Info
        return $true
    }

    $Script:LastInstallationError = $attempt.ErrorMessage
    Write-ColorOutput "    ❌ Retry failed: $($attempt.ErrorMessage)" -Type Error
    Write-ColorOutput ("    Total time elapsed: {0:N1} minutes" -f ((Get-Date) - $StartTime).TotalMinutes) -Type Info
    return $false
}

function Install-ModuleWithProgress {
    <#
    .SYNOPSIS
        Installs or updates PowerShell modules with progress tracking

    .DESCRIPTION
        Executes module installation/update operations in a background job, or in a timed
        child process when jobs are unavailable, while displaying progress and time estimates.

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
        [ValidateSet('Install', 'Update')]
        [string]$Operation = 'Install'
    )

    $InstallParams = $InstallParams.Clone()
    $InstallParams.ErrorAction = 'Stop'
    $Script:LastInstallationError = $null
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
        param($ModuleName, $InstallParams, $Operation, $LanguageMode, $ProgressFilePath, $Repository)
        $ErrorActionPreference = 'Stop'
        $ProgressPreference = 'SilentlyContinue'
        $ConfirmPreference = 'None'
        $PSDefaultParameterValues['*:Confirm'] = $false
        $progressFile = $null
        $originalInstallationPolicy = $null

        try {
            # Job runs in a new process; enable TLS 1.2 without disabling TLS 1.3 (see Enable-Tls12)
            try {
                $currentProtocol = [Net.ServicePointManager]::SecurityProtocol
                if ($currentProtocol -ne [Net.SecurityProtocolType]::SystemDefault) {
                    [Net.ServicePointManager]::SecurityProtocol = $currentProtocol -bor [Net.SecurityProtocolType]::Tls12
                }
                [Net.ServicePointManager]::DefaultConnectionLimit = 12
            }
            catch {
                Write-Verbose "Could not adjust network settings in job: $($_.Exception.Message)"
            }

            # Start-Job runs in a separate process, so provider and repository
            # consent from the parent session is not available here.
            $nugetProvider = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
            if (-not $nugetProvider) {
                Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -Confirm:$false -ErrorAction Stop
            }
            # Set-PSRepository persists to the user's profile, so remember the original policy
            # and put it back in 'finally' rather than leaving the repository trusted forever.
            $originalInstallationPolicy = (Get-PSRepository -Name $Repository -ErrorAction Stop).InstallationPolicy
            if ($originalInstallationPolicy -ne 'Trusted') {
                Set-PSRepository -Name $Repository -InstallationPolicy Trusted -ErrorAction Stop
            }

            # Create progress tracking for Az and Graph modules in constrained mode
            $isConstrainedLargeModule = ($LanguageMode -eq 'ConstrainedLanguage' -and
                                       ($ModuleName -eq 'Az' -or $ModuleName -eq 'Microsoft.Graph'))

            if ($isConstrainedLargeModule) {
                $progressFile = $ProgressFilePath
                "Starting $Operation of $ModuleName at $(Get-Date -Format 'HH:mm:ss') | Phase: Download and Installation" | Out-File $progressFile -Force
            }

            # Execute the installation/update
            switch ($Operation) {
                'Install' {
                    if ($isConstrainedLargeModule) {
                        "Status: Installing $ModuleName... | Phase: Installation Starting" | Out-File $progressFile -Append
                    }
                    $null = Install-Module -Name $ModuleName @InstallParams
                    if ($isConstrainedLargeModule) {
                        "Phase: Installation Complete - Starting Verification | Status: Verification may take several minutes" | Out-File $progressFile -Append
                    }
                }
                'Update' {
                    if ($isConstrainedLargeModule) {
                        "Status: Updating $ModuleName... | Phase: Update Starting" | Out-File $progressFile -Append
                    }
                    $null = Install-Module -Name $ModuleName @InstallParams
                    if ($isConstrainedLargeModule) {
                        "Phase: Update Complete - Starting Verification | Status: Verification may take several minutes" | Out-File $progressFile -Append
                    }
                }
                default {
                    if ($isConstrainedLargeModule) {
                        "Status: Installing $ModuleName (default)... | Phase: Installation Starting" | Out-File $progressFile -Append
                    }
                    $null = Install-Module -Name $ModuleName @InstallParams
                    if ($isConstrainedLargeModule) {
                        "Phase: Installation Complete - Starting Verification | Status: Verification may take several minutes" | Out-File $progressFile -Append
                    }
                }
            }

            # Final cleanup phase
            if ($isConstrainedLargeModule) {
                "Phase: Verification Complete - Final Cleanup | Status: Complete at $(Get-Date -Format 'HH:mm:ss')" | Out-File $progressFile -Append
            }

            # Old-version cleanup is performed by the parent process after the job completes;
            # script functions like Remove-OldModuleVersions are not available inside this job.
            return @{ Success = $true; Message = "$Operation completed successfully" }
        }
        catch {
            if ($progressFile -and (Test-Path $progressFile)) {
                "Error: $($_.Exception.Message) at $(Get-Date -Format 'HH:mm:ss')" | Out-File $progressFile -Append
            }

            # Enhanced error handling for AllowClobber conflicts in background jobs
            $errorMessage = $_.Exception.Message
            if ($errorMessage -like "*already available*" -or
                $errorMessage -like "*AllowClobber*" -or
                $errorMessage -like "*may override the existing commands*" -or
                ($ModuleName -eq 'PnP.PowerShell' -and ($errorMessage -like "*publisher*" -or $errorMessage -like "*SkipPublisherCheck*" -or $errorMessage -like "*signed module*"))) {

                if ($progressFile -and (Test-Path $progressFile)) {
                    "Retry: Command conflict detected - retrying with AllowClobber at $(Get-Date -Format 'HH:mm:ss')" | Out-File $progressFile -Append
                }

                # Force AllowClobber for conflict resolution
                try {
                    if ($Operation -in @('Install', 'Update')) {
                        $retryParams = $InstallParams.Clone()
                        if ($ModuleName -eq 'PnP.PowerShell') {
                            $retryParams.SkipPublisherCheck = $true
                            $retryParams.AllowClobber = $true
                        } else {
                            $retryParams.AllowClobber = $true
                        }
                        $null = Install-Module -Name $ModuleName @retryParams

                        if ($progressFile -and (Test-Path $progressFile)) {
                            "Retry: Success with AllowClobber at $(Get-Date -Format 'HH:mm:ss')" | Out-File $progressFile -Append
                        }

                        return @{ Success = $true; Message = "$Operation completed successfully after retry" }
                    } else {
                        return @{ Success = $false; Message = "Unsupported operation '$Operation'." }
                    }
                }
                catch {
                    if ($progressFile -and (Test-Path $progressFile)) {
                        "Retry: Failed with error: $($_.Exception.Message)" | Out-File $progressFile -Append
                    }
                    return @{ Success = $false; Message = $_.Exception.Message }
                }
            }

            return @{ Success = $false; Message = $_.Exception.Message }
        }
        finally {
            if ($progressFile -and (Test-Path -LiteralPath $progressFile)) {
                Remove-Item -LiteralPath $progressFile -Force -ErrorAction SilentlyContinue
            }
            if ($originalInstallationPolicy -and $originalInstallationPolicy -ne 'Trusted') {
                Set-PSRepository -Name $Repository -InstallationPolicy $originalInstallationPolicy -ErrorAction SilentlyContinue
            }
        }
    }

    # Check if background jobs are supported in this environment
    if (-not $Script:JobsSupported) {
        Write-ColorOutput "    Background jobs not supported - using a timed PowerShell child process" -Type Warning

        # Constrained Language Mode prevents jobs; the child process preserves bounded execution.
        if ($languageMode -eq 'ConstrainedLanguage' -and ($ModuleName -eq 'Az' -or $ModuleName -eq 'Microsoft.Graph')) {
            Write-ColorOutput "    ⚠️  $ModuleName will run in a separate process under Constrained Language Mode" -Type Warning
            Write-ColorOutput "    Progress is monitored, and the worker is terminated after $TimeoutMinutes minute(s)" -Type Info
            Write-ColorOutput ""
        }

        return (Invoke-ModuleInstallDirect -ModuleName $ModuleName -InstallParams $InstallParams -Operation $Operation -StartTime $startTime -TimeoutMinutes $TimeoutMinutes)
    }

    # Background job implementation for environments that support it
    # Start the background job (progress file path is generated here and passed in so both sides agree)
    $progressFile = Join-Path $env:TEMP "ModuleProgress_$($ModuleName)_$([guid]::NewGuid().ToString('N')).txt"
    try {
        $job = Start-Job -ScriptBlock $jobScript -ArgumentList $ModuleName, $InstallParams, $Operation, $languageMode, $progressFile, $Repository
    }
    catch {
        Write-ColorOutput "    Background job creation failed - falling back to a timed PowerShell child process" -Type Warning
        Write-ColorOutput "    Error: $($_.Exception.Message)" -Type Warning
        $Script:JobsSupported = $false

        return (Invoke-ModuleInstallDirect -ModuleName $ModuleName -InstallParams $InstallParams -Operation $Operation -StartTime $startTime -TimeoutMinutes $TimeoutMinutes)
    }

    # Enhanced progress tracking with phase detection
    $progressParams = @{
        Activity = "$Operation module: $ModuleName"
        Status = "Downloading and installing..."
        PercentComplete = 0
    }

    $iteration = 0
    $maxIterations = [Math]::Max(10, [Math]::Ceiling($estimate.EstimatedTime / 3))
    $lastProgressUpdate = Get-Date
    $verificationPhaseStarted = $false
    $isConstrainedLargeModule = ($languageMode -eq 'ConstrainedLanguage' -and
                               ($ModuleName -eq 'Az' -or $ModuleName -eq 'Microsoft.Graph'))

    try {
    while ($job.State -in @('NotStarted', 'Running', 'Blocked')) {
        $elapsed = (Get-Date) - $startTime
        $elapsedSeconds = $elapsed.TotalSeconds
        $elapsedMinutes = $elapsed.TotalMinutes
        if ($elapsedMinutes -ge $TimeoutMinutes) {
            throw "Installation exceeded $TimeoutMinutes minutes. Old versions will be retained."
        }

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
                # The job may be mid-write; a missed poll is harmless, the next one catches up
                Write-Verbose "Progress file read skipped: $($_.Exception.Message)"
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

        # Wait-Job throws when the child is blocked on an interactive prompt.
        # The job is configured as non-interactive above, but surface a useful
        # error if a provider or module still requests input.
        if ($job.State -eq 'Blocked') {
            throw "Installation job is blocked waiting for user interaction. Ensure the NuGet provider and $($Repository) trust settings are available."
        }
        $null = Wait-Job -Job $job -Timeout 1 -ErrorAction Stop
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

    if ($job.State -ne 'Completed') {
        throw "Installation job ended in state '$($job.State)': $($job.ChildJobs[0].JobStateInfo.Reason)"
    }
    $jobResults = @(Receive-Job -Job $job -ErrorAction Stop)
    if ($jobResults.Count -ne 1 -or $jobResults[0] -isnot [System.Collections.IDictionary] -or -not $jobResults[0].Contains('Success')) {
        throw 'Installation job returned an invalid result.'
    }
    $result = $jobResults[0]
    }
    catch {
        $result = @{ Success = $false; Message = $_.Exception.Message }
    }
    finally {
        if ($job.State -notin @('Completed', 'Failed', 'Stopped')) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
        }
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $progressFile -Force -ErrorAction SilentlyContinue
        Write-Progress -Activity "$Operation module: $ModuleName" -Completed
    }

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
        $Script:LastInstallationError = $result.Message
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

    $filteredModules = [System.Collections.Generic.List[object]]::new()

    foreach ($scopeItem in $Scope) {
        # @() so a category with a single module still reports Count = 1, not the hashtable key count
        $modulesInScope = @($Script:ModuleList.Where({ $_.Category -eq $scopeItem }))
        if ($modulesInScope.Count -gt 0) {
            $filteredModules.AddRange($modulesInScope)
            Write-ColorOutput "Added $($modulesInScope.Count) module(s) from $scopeItem category" -Type Info
        } else {
            Write-ColorOutput "Warning: No modules found for scope '$scopeItem'" -Type Warning
        }
    }

    # Remove duplicates and sort by priority
    $uniqueModules = @($filteredModules | Sort-Object { $_.Name } -Unique | Sort-Object { $_.Priority }, { $_.Name })

    Write-ColorOutput "Total modules selected: $($uniqueModules.Count)" -Type Process
    # A single result unrolls to a bare hashtable; callers must wrap the call in @()
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

    try {
        $source = Get-PSRepository -Name $Repository -ErrorAction Stop
        $sourceUri = [uri]$source.SourceLocation
        if ($sourceUri.IsAbsoluteUri -and $sourceUri.Scheme -in @('https', 'http')) {
            # PSGallery rejects HEAD requests with HTTP 405. A lightweight GET
            # verifies the repository endpoint without relying on that method.
            $null = Invoke-WebRequest -Uri $sourceUri -Method Get -TimeoutSec 10 -MaximumRedirection 5 -ErrorAction Stop
        }
        elseif (-not (Test-Path -LiteralPath $source.SourceLocation -PathType Container)) {
            throw "Repository path '$($source.SourceLocation)' is unavailable."
        }
        Write-ColorOutput "  Repository '$Repository' is reachable." -Type Process
        return $true
    }
    catch {
        Write-ColorOutput "  Repository check failed: $($_.Exception.Message)" -Type Warning
        return $false
    }
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
        IsRegistered = $false
        IsBundled = $false
        LocalVersion = $null
        OnlineVersion = $null
        UpdateAvailable = $false
        UpdateRecommended = $false
        InstallRecommended = $false
        MultipleVersionsInstalled = $false
        InstalledVersions = @()
        ErrorMessage = $null
    }

    try {
        # Get-InstalledModule -AllVersions costs ~20 s per call on large module trees, so only
        # pay for it when the module actually exists on PSModulePath (one cached scan).
        $presentOnDisk = Test-ModulePresentOnDisk -Name $ModuleName

        # Check for versions installed via PowerShellGet / Install-Module
        $installedModules = @()
        if ($presentOnDisk) {
            $installedModules = @(Get-InstalledModule -Name $ModuleName -AllVersions -ErrorAction SilentlyContinue |
                Sort-Object { ConvertTo-ModuleVersion $_.Version } -Descending)
        }
        $versionInfo.InstalledVersions = $installedModules

        if ($installedModules) {
            $versionInfo.IsInstalled = $true
            $versionInfo.IsRegistered = $true
            $versionInfo.LocalVersion = $installedModules[0].Version
            $versionInfo.MultipleVersionsInstalled = ($installedModules | Measure-Object).Count -gt 1
        }

        # Get-InstalledModule only sees modules registered by PowerShellGet. A module can also
        # be installed by another tool or copied into PSModulePath, so inspect available modules
        # before deciding that an installation is missing.
        if (-not $versionInfo.IsInstalled -and $presentOnDisk) {
            $psHomeModulePath = Join-Path $PSHOME 'Modules'
            $availableModules = @(Get-Module -Name $ModuleName -ListAvailable -ErrorAction SilentlyContinue |
                Sort-Object { ConvertTo-ModuleVersion $_.Version } -Descending)
            $availableModule = $availableModules | Select-Object -First 1

            if ($availableModule) {
                $versionInfo.IsInstalled = $true
                $versionInfo.InstalledVersions = $availableModules
                $versionInfo.LocalVersion = $availableModule.Version
                $versionInfo.MultipleVersionsInstalled = $availableModules.Count -gt 1
                $versionInfo.IsBundled = $availableModule.ModuleBase -like "$psHomeModulePath*"
                if ($versionInfo.IsBundled) {
                    Write-Verbose "Module '$ModuleName' found as PS7 bundled module v$($availableModule.Version) at $($availableModule.ModuleBase)"
                } else {
                    Write-Verbose "Module '$ModuleName' found on PSModulePath at v$($availableModule.Version) ($($availableModule.ModuleBase))"
                }
            }
        }

        # Check online version
        $onlineModule = Find-Module -Name $ModuleName -Repository $Repository -ErrorAction Stop

        if ($onlineModule) {
            $versionInfo.OnlineVersion = $onlineModule.Version

            if ($versionInfo.IsInstalled) {
                $localVersion = ConvertTo-ModuleVersion $versionInfo.LocalVersion
                $onlineVersion = ConvertTo-ModuleVersion $versionInfo.OnlineVersion
                if ($onlineVersion -gt $localVersion) {
                    $versionInfo.UpdateAvailable = $true
                    # Update-Module requires a PowerShellGet registration. For modules
                    # discovered only on PSModulePath, install the newer copy instead.
                    if (-not $versionInfo.IsRegistered) {
                        $versionInfo.InstallRecommended = $true
                    }

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

function Get-ModuleVersionInstallPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$ModuleVersion
    )

    foreach ($propertyName in @('ModuleBase', 'InstalledLocation')) {
        $property = $ModuleVersion.PSObject.Properties[$propertyName]
        if ($property -and -not [string]::IsNullOrWhiteSpace($property.Value)) {
            return $property.Value
        }
    }

    return $null
}

function Get-ModuleInstallRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter()]
        [AllowNull()]
        [string]$ModuleBase
    )

    if ([string]::IsNullOrWhiteSpace($ModuleBase)) {
        return $null
    }

    $leaf = Split-Path -Path $ModuleBase -Leaf
    $parsedVersion = [version]'0.0'
    if ([version]::TryParse($leaf, [ref]$parsedVersion)) {
        return (Split-Path -Path $ModuleBase -Parent)
    }

    if ($leaf -ieq $ModuleName) {
        return $ModuleBase
    }

    return (Split-Path -Path $ModuleBase -Parent)
}

function Get-ModuleVersionCleanupEntries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter()]
        [switch]$IncludeAvailable
    )

    $entriesByKey = @{}
    $psHomeModulePath = Join-Path $PSHOME 'Modules'

    $addEntry = {
        param(
            [object]$ModuleVersion,
            [string]$ModuleBase
        )

        $version = ConvertTo-ModuleVersion $ModuleVersion.Version
        $moduleRoot = Get-ModuleInstallRoot -ModuleName $ModuleName -ModuleBase $ModuleBase
        $keyPath = if ($ModuleBase) { $ModuleBase } else { '<unknown>' }
        $key = '{0}|{1}' -f $version.ToString(), $keyPath

        if (-not $entriesByKey.ContainsKey($key)) {
            $entriesByKey[$key] = [pscustomobject]@{
                Version = $version
                ModuleBase = $ModuleBase
                ModuleRoot = $moduleRoot
            }
        }
    }

    $registeredVersions = @(Get-InstalledModule -Name $ModuleName -AllVersions -ErrorAction SilentlyContinue)
    foreach ($version in $registeredVersions) {
        & $addEntry $version (Get-ModuleVersionInstallPath -ModuleVersion $version)
    }

    if ($IncludeAvailable) {
        $availableVersions = @(Get-Module -Name $ModuleName -ListAvailable -ErrorAction SilentlyContinue |
            Where-Object { $_.ModuleBase -notlike "$psHomeModulePath*" })
        foreach ($version in $availableVersions) {
            & $addEntry $version $version.ModuleBase
        }
    }

    return @($entriesByKey.Values | Sort-Object { ConvertTo-ModuleVersion $_.Version } -Descending)
}

function Remove-ModuleVersionForced {
    <#
    .SYNOPSIS
        Forcibly removes a single specific version of a module from disk

    .DESCRIPTION
        Unloads the module version from the current session (so Uninstall-Module isn't
        blocked by "module is in use"), then attempts Uninstall-Module. If that still
        fails - typically because the version was installed outside PowerShellGet by
        copying files into a PSModulePath folder - falls back to deleting the module's
        version folder directly from disk. Refuses to touch anything under $PSHOME\Modules
        (bundled PS7 modules) as a safety guard.

    .PARAMETER ModuleName
        Name of the module

    .PARAMETER Version
        Specific version to remove

    .PARAMETER ModuleBase
        Optional known module base path (from Get-Module -ListAvailable) used for the
        filesystem fallback. If not supplied, the path is looked up.

    .RETURNS
        Hashtable: @{ Success = [bool]; Method = 'Uninstall-Module'|'Filesystem'|'WhatIf'|$null; ErrorMessage = [string] }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter(Mandatory)]
        [version]$Version,

        [Parameter()]
        [string]$ModuleBase
    )

    $psHomeModulePath = Join-Path $PSHOME 'Modules'
    $outcome = @{ Success = $false; Method = $null; ErrorMessage = $null }

    if (-not $PSCmdlet.ShouldProcess("$ModuleName v$Version", 'Remove module version')) {
        $outcome.Method = 'WhatIf'
        return $outcome
    }

    # Unload the exact version from this session first, otherwise both Uninstall-Module
    # and a filesystem delete can fail with "file in use" if it was ever imported.
    try {
        Get-Module -Name $ModuleName -ErrorAction SilentlyContinue |
            Where-Object { $_.Version -eq $Version } |
            Remove-Module -Force -ErrorAction Stop
    }
    catch {
        Write-Verbose "Could not unload $ModuleName v$Version from session: $($_.Exception.Message)"
    }

    # First attempt: standard, supported removal path
    try {
        Uninstall-Module -Name $ModuleName -RequiredVersion $Version -Force -Confirm:$false -ErrorAction Stop
        $outcome.Success = $true
        $outcome.Method = 'Uninstall-Module'
        return $outcome
    }
    catch {
        $uninstallError = $_.Exception.Message
        Write-Verbose "Uninstall-Module failed for $ModuleName v$Version`: $uninstallError"
    }

    # Fallback: locate and delete the version folder directly. This covers modules that
    # were copied into PSModulePath rather than installed via PowerShellGet, which
    # Uninstall-Module cannot see or remove.
    try {
        if (-not $ModuleBase) {
            $candidate = Get-Module -Name $ModuleName -ListAvailable -ErrorAction SilentlyContinue |
                Where-Object { $_.Version -eq $Version } |
                Select-Object -First 1
            $ModuleBase = $candidate.ModuleBase
        }

        if (-not $ModuleBase -or -not (Test-Path -LiteralPath $ModuleBase)) {
            throw "Could not locate an on-disk path for $ModuleName v$Version"
        }

        if ($ModuleBase -like "$psHomeModulePath*") {
            throw "Refusing to delete '$ModuleBase' - it is a PS7 bundled module"
        }

        Remove-Item -LiteralPath $ModuleBase -Recurse -Force -Confirm:$false -ErrorAction Stop
        $outcome.Success = $true
        $outcome.Method = 'Filesystem'
    }
    catch {
        $outcome.ErrorMessage = (@($uninstallError, $_.Exception.Message) | Where-Object { $_ }) -join '; '
    }

    return $outcome
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

    .PARAMETER Force
        Performs a comprehensive removal: also detects versions that are only visible
        via Get-Module -ListAvailable (not registered with PowerShellGet), unloads the
        module from the session before removing, and falls back to direct filesystem
        deletion when Uninstall-Module cannot complete the removal.

    .RETURNS
        Hashtable with cleanup results
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter()]
        [switch]$KeepLatestOnly,

        [Parameter()]
        [switch]$Force
    )

    $result = @{
        Success = $false
        RemovedVersions = @()
        KeptVersions = @()
        ErrorMessage = $null
        RemovedCount = 0
    }

    try {
        $installedVersions = @(Get-ModuleVersionCleanupEntries -ModuleName $ModuleName -IncludeAvailable:$Force)
        $keepNewestPerRoot = $ModuleName -in $Script:BootstrapModuleNames

        if (-not $installedVersions -or $installedVersions.Count -le 1) {
            Write-Verbose "No cleanup needed for $ModuleName - only one or no versions installed"
            $result.Success = $true
            if ($installedVersions) {
                $result.KeptVersions = @($installedVersions.Version.ToString())
            }
            return $result
        }

        if ($keepNewestPerRoot) {
            $versionsToRemove = @()
            foreach ($rootGroup in ($installedVersions | Group-Object -Property ModuleRoot)) {
                $rootVersions = @($rootGroup.Group | Sort-Object { ConvertTo-ModuleVersion $_.Version } -Descending)
                $versionsToRemove += @($rootVersions | Select-Object -Skip 1)
            }

            if ($versionsToRemove) {
                Write-ColorOutput "    Protecting $ModuleName bootstrap availability: keeping newest version per module root" -Type Info
            }
        } else {
            # Determine how many versions to keep (default to keeping only latest if no parameter specified)
            $versionsToKeep = if ($KeepLatestOnly -or $Force -or $PSBoundParameters.Count -eq 1) { 1 } else { 2 }
            $versionsToRemove = @($installedVersions | Select-Object -Skip $versionsToKeep)
        }

        if (-not $versionsToRemove) {
            Write-Verbose "No old versions to remove for $ModuleName"
            $result.Success = $true
            $result.KeptVersions = $installedVersions | ForEach-Object { $_.Version.ToString() }
            return $result
        }

        Write-ColorOutput "    🧹 Cleaning up old versions of $ModuleName..." -Type Info

        foreach ($versionToRemove in $versionsToRemove) {
            try {
                if (-not $PSCmdlet.ShouldProcess("$ModuleName v$($versionToRemove.Version)", 'Remove old module version')) {
                    continue
                }
                Write-Verbose "Removing $ModuleName version $($versionToRemove.Version)"

                if ($Force) {
                    $forcedResult = Remove-ModuleVersionForced -ModuleName $ModuleName -Version $versionToRemove.Version -ModuleBase $versionToRemove.ModuleBase -Confirm:$false
                    if (-not $forcedResult.Success) {
                        throw $forcedResult.ErrorMessage
                    }
                    Write-ColorOutput "      ✅ Removed version $($versionToRemove.Version) (via $($forcedResult.Method))" -Type Process
                } else {
                    try {
                        Uninstall-Module -Name $ModuleName -RequiredVersion $versionToRemove.Version -Confirm:$false -ErrorAction Stop
                        Write-ColorOutput "      ✅ Removed version $($versionToRemove.Version)" -Type Process
                    }
                    catch {
                        if ($_.Exception.Message -notmatch 'Administrator rights are required') {
                            throw
                        }
                        # PowerShellGet's own elevation check can report a false negative in an
                        # elevated session (observed under ConstrainedLanguage mode). The folder ACL
                        # is the real authority, so try deleting that exact version folder instead.
                        $uninstallError = $_.Exception.Message
                        Write-Verbose "Uninstall-Module reported missing admin rights for $ModuleName v$($versionToRemove.Version); trying direct folder removal"
                        $fallbackResult = Remove-ModuleVersionForced -ModuleName $ModuleName -Version $versionToRemove.Version -ModuleBase $versionToRemove.ModuleBase -Confirm:$false
                        if (-not $fallbackResult.Success) {
                            throw "$uninstallError Direct folder removal also failed: $($fallbackResult.ErrorMessage)"
                        }
                        Write-ColorOutput "      ✅ Removed version $($versionToRemove.Version) (via $($fallbackResult.Method))" -Type Process
                    }
                }

                $result.RemovedVersions += $versionToRemove.Version.ToString()
                $result.RemovedCount++
            }
            catch {
                $errorMsg = $_.Exception.Message
                $result.ErrorMessage = (@($result.ErrorMessage, $errorMsg) | Where-Object { $_ }) -join '; '
                Write-ColorOutput "      ⚠️ Could not remove version $($versionToRemove.Version): $errorMsg" -Type Warning

                # Don't fail the entire operation if we can't remove one version
                Write-Verbose "Continuing cleanup despite error removing version $($versionToRemove.Version)"
            }
        }

        # Record kept versions
        $result.KeptVersions = @($installedVersions | Where-Object { $_.Version.ToString() -notin $result.RemovedVersions } |
            ForEach-Object { $_.Version.ToString() })

        # Consider operation successful even if some versions couldn't be removed
        $result.Success = -not $result.ErrorMessage

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

    .PARAMETER Force
        Performs a comprehensive scan/removal that also includes modules only visible via
        Get-Module -ListAvailable (not registered with PowerShellGet), and uses
        Remove-ModuleVersionForced (session unload + filesystem fallback) for each removal.

    .RETURNS
        Hashtable with overall cleanup results
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string[]]$ExcludeModules = @(),

        [Parameter()]
        [switch]$KeepLatestOnly,

        [Parameter()]
        [switch]$Force
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
        # Get all modules registered via PowerShellGet, grouped by name
        $allInstalledModules = Get-InstalledModule -ErrorAction SilentlyContinue
        $moduleNames = @($allInstalledModules | Select-Object -ExpandProperty Name -Unique)

        if ($Force) {
            # Comprehensive scan: also include anything visible on PSModulePath, even if
            # PowerShellGet never registered it (e.g. copied in manually, or side-by-side
            # versions installed by a different tool). Bundled PS7 modules are excluded.
            $psHomeModulePath = Join-Path $PSHOME 'Modules'
            $availableModuleNames = @(Get-Module -ListAvailable -ErrorAction SilentlyContinue |
                Where-Object { $_.ModuleBase -notlike "$psHomeModulePath*" } |
                Select-Object -ExpandProperty Name -Unique)
            $moduleNames = @($moduleNames + $availableModuleNames | Select-Object -Unique)
        }

        # Find modules with multiple versions using the same combined logic that
        # Remove-OldModuleVersions applies, so counts here reflect what will actually happen
        $modulesWithMultipleVersions = foreach ($name in $moduleNames) {
            if ($name -in $ExcludeModules) { continue }

            $cleanupEntries = @(Get-ModuleVersionCleanupEntries -ModuleName $name -IncludeAvailable:$Force)
            if ($name -in $Script:BootstrapModuleNames) {
                $rootsWithMultipleVersions = @($cleanupEntries | Group-Object -Property ModuleRoot | Where-Object { $_.Count -gt 1 })
                $versionCount = if ($rootsWithMultipleVersions) {
                    ($rootsWithMultipleVersions | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
                } else {
                    0
                }
            } else {
                $versionCount = @($cleanupEntries.Version | Select-Object -Unique).Count
            }

            if ($versionCount -gt 1) {
                [pscustomobject]@{ Name = $name; Count = $versionCount }
            }
        }

        if (-not $modulesWithMultipleVersions) {
            Write-ColorOutput "✅ No modules with multiple versions found - cleanup not needed" -Type Process
            return $overallResult
        }

        Write-ColorOutput "Found $(@($modulesWithMultipleVersions).Count) module(s) with multiple versions" -Type Info

        foreach ($moduleInfo in $modulesWithMultipleVersions) {
            $moduleName = $moduleInfo.Name

            $overallResult.ProcessedModules++

            Write-ColorOutput "Processing: $moduleName ($($moduleInfo.Count) versions)" -Type Info

            # Perform cleanup for this module
            $cleanupResult = Remove-OldModuleVersions -ModuleName $moduleName -KeepLatestOnly:$KeepLatestOnly -Force:$Force

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

function Test-AllUsersModuleWriteAccess {
    <#
    .SYNOPSIS
        Verifies effective write access to the PowerShell 7 AllUsers module directory.
    .DESCRIPTION
        Tests the actual target directory rather than relying only on administrator group
        membership, which can disagree with the current UAC token or directory ACLs.
        Creates and removes a uniquely named, empty probe file before any module changes.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ModulePath = (Join-Path $env:ProgramFiles 'PowerShell\Modules')
    )

    $probePath = Join-Path $ModulePath ".o365-update-write-test-$([guid]::NewGuid().ToString('N')).tmp"
    $created = $false

    try {
        if (-not (Test-Path -LiteralPath $ModulePath -PathType Container)) {
            throw "AllUsers module directory '$ModulePath' does not exist."
        }

        if (-not $PSCmdlet.ShouldProcess($ModulePath, 'Create and remove write-access probe')) {
            return [pscustomobject]@{
                CanWrite = $null
                ModulePath = $ModulePath
                ErrorMessage = 'Write-access probe was skipped by ShouldProcess.'
            }
        }

        $null = New-Item -Path $probePath -ItemType File -ErrorAction Stop
        $created = $true
        Remove-Item -LiteralPath $probePath -Force -ErrorAction Stop
        $created = $false

        return [pscustomobject]@{
            CanWrite = $true
            ModulePath = $ModulePath
            ErrorMessage = $null
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($created -and (Test-Path -LiteralPath $probePath)) {
            try {
                Remove-Item -LiteralPath $probePath -Force -ErrorAction Stop
            }
            catch {
                Write-Verbose "Could not remove permission probe '$probePath': $($_.Exception.Message)"
            }
        }

        return [pscustomobject]@{
            CanWrite = $false
            ModulePath = $ModulePath
            ErrorMessage = $errorMessage
        }
    }
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
    Write-ColorOutput "PowerShell Version: $psVersion" -Type Info

    if ($psVersion.Major -lt 7) {
        Write-ColorOutput "❌ PowerShell 7.0 or higher is required" -Type Error
        $allPrerequisitesMet = $false
    } else {
        Write-ColorOutput "✅ PowerShell version is supported" -Type Success
    }

    # Check administrator privileges
    try {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

        if ($isAdmin) {
            Write-ColorOutput "✅ Running with administrator privileges" -Type Success
        } elseif ($CheckOnly -or $CheckSessions -or $WhatIfPreference) {
            Write-ColorOutput "ℹ️ Administrator privileges are not required for this read-only operation" -Type Info
        } else {
            Write-ColorOutput "❌ Administrator privileges are required (AllUsers scope installs)" -Type Error
            $allPrerequisitesMet = $false
        }

        if ($isAdmin -and -not $CheckOnly -and -not $CheckSessions -and -not $WhatIfPreference) {
            $writeAccess = Test-AllUsersModuleWriteAccess
            if ($writeAccess.CanWrite) {
                Write-ColorOutput "✅ Write access verified for $($writeAccess.ModulePath)" -Type Success
            } else {
                Write-ColorOutput "❌ Cannot write to the AllUsers module directory '$($writeAccess.ModulePath)'." -Type Error
                Write-ColorOutput "   No module changes were made. Open PowerShell 7 with 'Run as administrator' and try again." -Type Info
                Write-ColorOutput "   Permission check failed: $($writeAccess.ErrorMessage)" -Type Info
                $allPrerequisitesMet = $false
            }
        }
    }
    catch {
        Write-ColorOutput "⚠️ Could not verify administrator privileges: $($_.Exception.Message)" -Type Warning
        if (-not $CheckOnly -and -not $CheckSessions -and -not $WhatIfPreference) {
            $allPrerequisitesMet = $false
        }
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
    [CmdletBinding(SupportsShouldProcess)]
    param()

    $modules = @($Script:FilteredModuleList)
    $totalModules = $modules.Count
    Write-ProgressHeader "Module Processing" "Processing $totalModules modules"

    $processedCount = 0
    $successCount = 0
    $failureCount = 0
    $cleanupFailureCount = 0
    $cleanupCandidates = [System.Collections.Generic.List[object]]::new()
    $Script:ModuleOperationResults = @()

    foreach ($module in $modules) {
        $processedCount++

        Write-Progress -Activity 'Microsoft cloud module update' -Status "$($module.Name) ($processedCount of $totalModules)" -PercentComplete ([int](($processedCount - 1) / [Math]::Max($totalModules, 1) * 100))
        Write-ColorOutput "($processedCount of $totalModules) Processing: $($module.Description)" -Type Info

        # Get version information
        $versionInfo = Get-ModuleVersionInfo -ModuleName $module.Name

        if ($versionInfo.ErrorMessage) {
            Write-ColorOutput "  ❌ Error checking module: $($versionInfo.ErrorMessage)" -Type Error
            Write-ColorOutput "     Fix: check connectivity to repository '$Repository' (or use -Repository) and re-run with -Verbose." -Type Info
            Add-ModuleResult -Module $module -Status 'Failed' -Action 'Check' -LocalVersion $versionInfo.LocalVersion -ErrorMessage $versionInfo.ErrorMessage
            $failureCount++
            continue
        }

        # Display current status
        if ($versionInfo.IsInstalled) {
            Write-ColorOutput "  Current version: $($versionInfo.LocalVersion)" -Type Info
            if ($versionInfo.MultipleVersionsInstalled) {
                Write-ColorOutput "  ⚠️ Multiple versions installed" -Type Warning

                # Always inspect installed versions when duplicates are detected
                $allVersions = $versionInfo.InstalledVersions

                if ($allVersions -and $allVersions.Count -gt 1) {
                    Write-ColorOutput "  📋 Installed versions: $($allVersions.Version -join ', ')" -Type Info
                }

                # Default behavior: enforce latest-only maintenance immediately
                if (-not $CheckOnly -and -not $SkipVersionCleanup) {
                    Write-ColorOutput '  Old-version cleanup deferred until processing succeeds.' -Type Info
                }
                elseif (-not $CheckOnly -and $SkipVersionCleanup -and $Prompt -and -not $Script:SkipCleanupPrompts) {
                    Write-ColorOutput "  💡 Multiple versions detected for $($module.Name)" -Type Info

                    if ($allVersions -and $allVersions.Count -gt 1) {
                        # Interactive prompt for cleanup
                        $cleanupChoice = $null
                        Write-ColorOutput "  🧹 Would you like to clean up old versions now?" -Type Warning
                        Write-ColorOutput "     This will keep only the latest version ($($versionInfo.LocalVersion))" -Type Info

                        do {
                            $cleanupChoice = ([string](Read-Host "  Clean up old versions? (Y/N/S for Skip all prompts)")).Trim().ToUpper()

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
                } elseif ($CheckOnly) {
                    # In check-only mode, just show the versions
                    if ($allVersions -and $allVersions.Count -gt 1) {
                        Write-ColorOutput "  📋 Installed versions: $($allVersions.Version -join ', ')" -Type Info
                    }
                } else {
                    Write-ColorOutput "  ⏭️ Version cleanup skipped by parameter - multiple versions will be preserved" -Type Warning
                }
            }
        } else {
            Write-ColorOutput "  Module not installed" -Type Warning
        }

        Write-ColorOutput "  Latest version: $($versionInfo.OnlineVersion)" -Type Info

        if ($CheckOnly) {
            if ($versionInfo.UpdateAvailable) {
                Write-ColorOutput "  📋 Update available" -Type Warning
                $checkStatus = 'UpdateAvailable'
            } elseif ($versionInfo.InstallRecommended) {
                Write-ColorOutput "  📋 Installation recommended" -Type Warning
                $checkStatus = 'InstallRecommended'
            } else {
                Write-ColorOutput "  ✅ Up to date" -Type Process
                $checkStatus = 'UpToDate'
            }
            Add-ModuleResult -Module $module -Status $checkStatus -Action 'Check' -LocalVersion $versionInfo.LocalVersion -OnlineVersion $versionInfo.OnlineVersion
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

        # Skip installation for modules that are already bundled with PS7.
        # Bundled modules (e.g. PowerShellGet, PackageManagement) live inside $PSHOME\Modules
        # and are loaded automatically. Installing the same version just produces confusing
        # 'currently in use' warnings with no real benefit.
        if ($action -eq 'Install' -and $versionInfo.IsBundled) {
            $bundledVer = $versionInfo.LocalVersion
            Write-ColorOutput "  + Bundled with PS7 (v$bundledVer in $PSHOME\Modules) - skipping install" -Type Process
            Write-ColorOutput "    To get a standalone copy, run: Install-Module $($module.Name) -Force -AllowClobber -Scope AllUsers" -Type Info
            Add-ModuleResult -Module $module -Status 'Bundled' -Action 'None' -LocalVersion $versionInfo.LocalVersion -OnlineVersion $versionInfo.OnlineVersion
            $successCount++
            Write-ColorOutput ""
            continue
        }

        if ($action) {
            if ($Prompt) {
                $response = Read-Host "  $action $($module.Name)? (Y/N)"
                if ($response -notmatch '^[Yy]') {
                    Write-ColorOutput "  ⏭️ Skipped by user" -Type Warning
                    Add-ModuleResult -Module $module -Status 'Skipped' -Action $action -LocalVersion $versionInfo.LocalVersion -OnlineVersion $versionInfo.OnlineVersion
                    $Script:SkippedCount++
                    continue
                }
            }

            # Installs run in a separate job process that does not inherit -WhatIf, so gate here.
            if (-not $PSCmdlet.ShouldProcess("$($module.Name) $($versionInfo.LocalVersion) -> $($versionInfo.OnlineVersion) (AllUsers)", "$action module")) {
                Add-ModuleResult -Module $module -Status 'WhatIf' -Action $action -LocalVersion $versionInfo.LocalVersion -OnlineVersion $versionInfo.OnlineVersion
                $Script:SkippedCount++
                Write-ColorOutput ""
                continue
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
                    RequiredVersion = $versionInfo.OnlineVersion
                }

                # Call the existing function
                $success = Install-ModuleWithProgress -ModuleName $module.Name -InstallParams $installParams -Operation $action
                if ($success) {
                    $verified = @(Get-InstalledModule -Name $module.Name -RequiredVersion $versionInfo.OnlineVersion -ErrorAction Stop |
                        Where-Object { (ConvertTo-ModuleVersion $_.Version) -eq (ConvertTo-ModuleVersion $versionInfo.OnlineVersion) })
                    if ($verified.Count -eq 0) {
                        throw "Installation did not produce $($module.Name) v$($versionInfo.OnlineVersion)."
                    }
                }
                $result.Success = $success
                if (-not $success) { $result.ErrorMessage = $Script:LastInstallationError }
                $result.Duration = (Get-Date) - $startTime
            }
            catch {
                $result.ErrorMessage = $_.Exception.Message
                $result.Duration = (Get-Date) - $startTime
            }

            if ($result.Success) {
                $successCount++
                $cleanupCandidates.Add($module)
            } else {
                $failureCount++
            }
            $resultStatus = if ($result.Success) { 'Succeeded' } else { 'Failed' }
            Add-ModuleResult -Module $module -Status $resultStatus -Action $action -LocalVersion $versionInfo.LocalVersion -OnlineVersion $versionInfo.OnlineVersion -Duration $result.Duration -ErrorMessage $result.ErrorMessage

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
            Add-ModuleResult -Module $module -Status 'UpToDate' -Action 'None' -LocalVersion $versionInfo.LocalVersion -OnlineVersion $versionInfo.OnlineVersion
            $successCount++
            if ($versionInfo.MultipleVersionsInstalled) { $cleanupCandidates.Add($module) }
        }

        Write-ColorOutput ""
    }
    Write-Progress -Activity 'Microsoft cloud module update' -Completed

    # Final pass to ensure processed modules are maintained in a latest-only state
    if (-not $CheckOnly -and -not $SkipVersionCleanup) {
        Write-ColorOutput ""
        Write-ColorOutput "🧹 Final latest-only cleanup verification..." -Type System

        foreach ($module in $cleanupCandidates) {
            $finalCleanupResult = Remove-OldModuleVersions -ModuleName $module.Name -KeepLatestOnly
            if (-not $finalCleanupResult.Success) {
                $cleanupFailureCount++
                Write-ColorOutput "  ⚠️ Final cleanup could not fully complete for $($module.Name): $($finalCleanupResult.ErrorMessage)" -Type Warning
            }
        }
    }

    # Display summary
    Write-ProgressHeader "Processing Summary"
    Write-ColorOutput "Total modules processed: $processedCount" -Type Info
    Write-ColorOutput "Successful operations: $successCount" -Type Process
    Write-ColorOutput "Skipped (user/WhatIf): $($Script:SkippedCount)" -Type Info
    Write-ColorOutput "Failed operations: $failureCount" -Type Error
    Write-ColorOutput "Failed final cleanups: $cleanupFailureCount" -Type Error
    $Script:ProcessingFailureCount = $failureCount + $cleanupFailureCount

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
        foreach ($module in $modules) {
            $installedVersions = @(Get-InstalledModule -Name $module.Name -AllVersions -ErrorAction SilentlyContinue)
            if ($installedVersions.Count -gt 1) {
                $modulesWithMultipleVersions += @{
                    Name = $module.Name
                    Description = $module.Description
                    VersionCount = $installedVersions.Count
                    Versions = ($installedVersions | Sort-Object { ConvertTo-ModuleVersion $_.Version } -Descending | ForEach-Object { $_.Version.ToString() })
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
    [CmdletBinding(SupportsShouldProcess)]
    param()

    Write-ProgressHeader "Deprecated Module Cleanup"

    $removedCount = 0
    $failedCount = 0

    foreach ($deprecatedModule in $Script:DeprecatedModules) {
        try {
            # Most deprecated modules are absent; the cached on-disk check avoids a ~20 s
            # Get-InstalledModule call for each of them.
            if (-not (Test-ModulePresentOnDisk -Name $deprecatedModule.Name)) {
                Write-Verbose "Deprecated module $($deprecatedModule.Name) not present on PSModulePath"
                continue
            }

            $installedVersions = @(Get-InstalledModule -Name $deprecatedModule.Name -AllVersions -ErrorAction SilentlyContinue)

            if ($installedVersions.Count -gt 0) {
                Write-ColorOutput "Found deprecated module: $($deprecatedModule.Name)" -Type Warning
                Write-ColorOutput "  Reason: $($deprecatedModule.Reason)" -Type Info
                Write-ColorOutput "  Replacement: $($deprecatedModule.Replacement)" -Type Info
                Write-ColorOutput "  Versions installed: $($installedVersions.Count)" -Type Info

                foreach ($version in $installedVersions) {
                    if (-not $PSCmdlet.ShouldProcess("$($deprecatedModule.Name) v$($version.Version)", 'Uninstall deprecated module')) {
                        continue
                    }
                    try {
                        Write-ColorOutput "  Removing version $($version.Version)..." -Type Process
                        Uninstall-Module -Name $deprecatedModule.Name -RequiredVersion $version.Version -Force -Confirm:$false -ErrorAction Stop
                        Write-ColorOutput "    ✅ Removed successfully" -Type Process
                        $removedCount++
                    }
                    catch {
                        Write-ColorOutput "    ⚠️ Could not remove: $($_.Exception.Message)" -Type Warning
                        $failedCount++
                    }
                }
            }
        }
        catch {
            Write-ColorOutput "Error checking for $($deprecatedModule.Name)`: $($_.Exception.Message)" -Type Error
            $failedCount++
        }
    }

    if ($removedCount -gt 0) {
        Write-ColorOutput "✅ Removed $removedCount deprecated module version(s)" -Type Process
        # Removals change what is on disk; refresh the cache before version checks use it
        $null = Test-ModulePresentOnDisk -Name 'PowerShellGet' -Refresh
    } else {
        Write-ColorOutput "✅ No deprecated modules found or removal needed" -Type Process
    }

    return @{
        RemovedCount = $removedCount
        FailedCount = $failedCount
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

#endregion Helper functions

#region Main logic

try {
    # Early elevation check (friendlier than the generic #Requires message)
    try {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        $isAdmin = $false
    }

    # -WhatIf makes no changes, so like -CheckOnly it may run without elevation
    if (-not $isAdmin -and -not $CheckOnly -and -not $CheckSessions -and -not $WhatIfPreference) {
        $scriptPath = $PSCommandPath
        $argumentParts = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "'$($scriptPath -replace "'", "''")'")
        foreach ($boundName in $PSBoundParameters.Keys | Sort-Object) {
            $boundValue = $PSBoundParameters[$boundName]
            if ($boundValue -is [System.Management.Automation.SwitchParameter]) {
                if ($boundValue) {
                    $argumentParts += "-$boundName"
                }
            } elseif ($boundValue -is [array]) {
                $argumentParts += "-$boundName"
                foreach ($item in $boundValue) {
                    $argumentParts += "'$($item.ToString() -replace "'", "''")'"
                }
            } else {
                $argumentParts += "-$boundName"
                $argumentParts += "'$($boundValue.ToString() -replace "'", "''")'"
            }
        }
        $elevateCmd = "Start-Process -Verb RunAs pwsh -ArgumentList @($($argumentParts -join ', '))"

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

    # Start transcript logging if requested
    if ($CreateLog) {
        $logFile = Join-Path $LogPath "o365-update-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
        # Transcript cmdlets honour -WhatIf; logging must still happen during a dry run
        Start-Transcript -LiteralPath $logFile -Append -WhatIf:$false | Out-Null
        Write-ColorOutput "📋 Logging to: $logFile" -Type System
    }

    # Show header (Clear-Host removed to avoid console issues in some environments)
    Write-ProgressHeader "Microsoft Cloud PowerShell Module Updater v2.18" "Enhanced version with improved module management and version cleanup"

    # Display configuration
    Write-ColorOutput "Configuration:" -Type Info
    $displayMode = if ($CheckSessions) { 'Session Check' } elseif ($CheckOnly) { 'Check Only' } elseif ($WhatIfPreference) { 'Update (WhatIf - no changes)' } else { 'Update' }
    Write-ColorOutput "  Mode: $displayMode" -Type Info
    Write-ColorOutput "  Scope: $($ModuleScope -join ', ')" -Type Info
    Write-ColorOutput "  Repository: $Repository" -Type Info
    Write-ColorOutput "  Prompt: $Prompt" -Type Info
    Write-ColorOutput "  Skip Deprecated Cleanup: $SkipDeprecatedCleanup" -Type Info
    Write-ColorOutput "  Skip Version Cleanup: $SkipVersionCleanup" -Type Info
    Write-ColorOutput "  Force Version Cleanup: $ForceVersionCleanup" -Type Info
    Write-ColorOutput '  Installation mode: sequential (avoids shared dependency conflicts)' -Type Info
    Write-ColorOutput "  Timeout: $TimeoutMinutes minutes" -Type Info
    Write-ColorOutput ""

    if ($TerminateConflicts) {
        Write-ColorOutput "⚠️ -TerminateConflicts is not supported in this version and will be ignored" -Type Warning
        Write-ColorOutput "   Close other PowerShell sessions manually before updating" -Type Info
        Write-ColorOutput ""
    }

    # Check prerequisites
    Write-ColorOutput "🔍 Checking prerequisites..." -Type System
    $prerequisitesMet = Test-Prerequisites
    if (-not $prerequisitesMet) {
        Write-ColorOutput "❌ Prerequisites not met. Please resolve the issues above and try again." -Type Error
        exit 1
    }
    Write-ColorOutput "✅ Prerequisites check passed" -Type Process

    # Handle session check mode (before slower network checks)
    if ($CheckSessions) {
        Write-ProgressHeader "PowerShell Session Check Mode"
        $otherSessions = Get-Process -Name 'pwsh', 'powershell', 'powershell_ise' -ErrorAction SilentlyContinue |
                         Where-Object { $_.Id -ne $PID }
        if ($otherSessions) {
            Write-ColorOutput "⚠️ Other PowerShell sessions detected that may lock modules during updates:" -Type Warning
            foreach ($session in $otherSessions) {
                Write-ColorOutput "  • $($session.ProcessName) (PID $($session.Id))" -Type Info
            }
            Write-ColorOutput "Close these sessions before updating to avoid 'module in use' conflicts." -Type Info
        } else {
            Write-ColorOutput "✅ No other PowerShell sessions detected" -Type Process
        }
        Write-ColorOutput "Run the script without -CheckSessions to proceed with updates." -Type Info
        exit 0
    }

    # Test network connectivity unless skipped
    if (-not $SkipConnectivityCheck) {
        Write-ColorOutput "🌐 Testing network connectivity..." -Type System
        $connectivityOk = Test-NetworkConnectivity
        if (-not $connectivityOk) {
            if ($Prompt) {
                $continueResponse = Read-Host "Network connectivity issues detected. Continue anyway? (Y/N)"
                if ($continueResponse -notmatch '^[Yy]') {
                    Write-ColorOutput "Script execution cancelled due to connectivity issues." -Type Warning
                    exit 1
                }
            } else {
                Write-ColorOutput "❌ Repository '$Repository' is unreachable. Fix connectivity or use -SkipConnectivityCheck to continue anyway." -Type Error
                exit 1
            }
        }
        Write-ColorOutput "✅ Network connectivity check completed" -Type Process
    } else {
        Write-ColorOutput "⏭️ Network connectivity check skipped" -Type Warning
    }

    # Filter modules based on scope
    Write-ColorOutput "📋 Filtering modules based on scope..." -Type System
    $Script:FilteredModuleList = @(Get-FilteredModuleList -Scope $ModuleScope)

    if ($Script:FilteredModuleList.Count -eq 0) {
        Write-ColorOutput "❌ No modules selected for processing. Please check your ModuleScope parameter." -Type Error
        exit 1
    }

    Write-ColorOutput "✅ Found $($Script:FilteredModuleList.Count) module(s) to process" -Type Process

    # Ensure the NuGet package provider is available before any module installs
    if (-not $CheckOnly) {
        Write-ColorOutput "📦 Verifying package provider..." -Type System
        $packageProviderReady = Test-PackageProvider -PackageName 'NuGet'
        if (-not $packageProviderReady) {
            Write-ColorOutput "❌ NuGet package provider is unavailable; module installation cannot continue." -Type Error
            exit 1
        }
    }

    # Clean up deprecated modules first
    $Script:ProcessingFailureCount = 0
    $deprecatedCleanupResult = $null
    if ($CheckOnly) {
        Write-ColorOutput "⏭️ Deprecated module cleanup skipped in check-only mode" -Type Info
    } elseif (-not $SkipDeprecatedCleanup) {
        Write-ColorOutput "🧹 Checking for deprecated modules..." -Type System
        $deprecatedCleanupResult = Remove-DeprecatedModules
    } else {
        Write-ColorOutput "⏭️ Deprecated module cleanup skipped" -Type Warning
    }

    # Process modules
    Write-ColorOutput "🚀 Starting module processing..." -Type System
    Start-ModuleProcessing

    if ($deprecatedCleanupResult -and $deprecatedCleanupResult.FailedCount -gt 0) {
        Write-ColorOutput "⚠️ $($deprecatedCleanupResult.FailedCount) deprecated module version(s) could not be removed." -Type Warning
        $Script:ProcessingFailureCount = ($Script:ProcessingFailureCount + $deprecatedCleanupResult.FailedCount)
    }

    # Comprehensive forced cleanup: runs regardless of ModuleScope, -Prompt or
    # -SkipVersionCleanup, and independent of whether module processing had failures,
    # since removing stale old versions is a separate concern from update success.
    if ($ForceVersionCleanup -and -not $CheckOnly) {
        if ($SkipVersionCleanup) {
            Write-ColorOutput "ℹ️ -ForceVersionCleanup overrides -SkipVersionCleanup" -Type Info
        }
        Write-ProgressHeader "Forced Comprehensive Module Version Cleanup"
        Write-ColorOutput "🧹 Scanning ALL installed modules (not limited to -ModuleScope) for old versions..." -Type System
        $forcedCleanupResult = Remove-AllOldModuleVersions -Force -KeepLatestOnly
        if ($forcedCleanupResult.FailedCleanups -gt 0) {
            Write-ColorOutput "⚠️ $($forcedCleanupResult.FailedCleanups) module(s) could not be fully cleaned up. Close any sessions using them and re-run with -ForceVersionCleanup." -Type Warning
            $Script:ProcessingFailureCount += $forcedCleanupResult.FailedCleanups
        }
    }

    if ($Script:ProcessingFailureCount -gt 0) {
        Write-ProgressHeader 'Script Completed With Errors'
        Write-ColorOutput "$Script:ProcessingFailureCount module operation(s) failed. Review the errors above." -Type Error
        exit 1
    }

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

        if ($ForceVersionCleanup) {
            Write-ColorOutput "  • A comprehensive forced cleanup of old module versions was performed" -Type Info
        } elseif (-not $SkipVersionCleanup) {
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
    #region Cleanup
    Write-Progress -Activity 'Microsoft cloud module update' -Completed

    # Emitted here so results reach the pipeline on every exit path, including exit 1
    if ($PassThru) {
        $Script:ModuleResults.ToArray()
    }

    if ($CreateLog) {
        try {
            Stop-Transcript -WhatIf:$false | Out-Null
            Write-ColorOutput "📋 Log file saved successfully" -Type System
        }
        catch {
            Write-Verbose "Transcript was not running: $($_.Exception.Message)"
        }
    }
    #endregion Cleanup
}

#endregion Main logic