<#
.SYNOPSIS
    Collect and display Microsoft Defender (Windows Security) configuration, ASR rule status, and definition versions.

.DESCRIPTION
    Queries local Microsoft Defender preferences, Attack Surface Reduction (ASR) rule states, scan/update settings,
    remediation actions, signature / engine versions (optionally comparing to Microsoft published latest),
    and basic Device Guard / Credential Guard registry configuration.

    This script includes enhanced detection methods for enterprise security features like EDR in Block Mode,
    Tamper Protection, and Credential Guard that may show as "Unknown" in enterprise environments.

    Provides multiple output modes:
      - Console: Colored console summary (default)
      - Plain: Plain text (no colors)
      - Object: Returns rich PowerShell object to the pipeline
      - Json: Writes JSON file (and/or returns object)
      - Csv: Exports to CSV format
      - Html: Creates HTML report with styling
      - Xml: Exports to XML using Export-Clixml

.NOTES
    Manual Verification Commands (run in elevated PowerShell):
    
    # Check MDE onboarding status
    Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status' | Select-Object OnboardingState, LastConnected, OrgId, SenseVersion
    Get-Service -Name 'Sense', 'WdFilter' | Select-Object Name, Status, StartType
    
    # Check Tamper Protection directly
    Get-MpComputerStatus | Select-Object IsTamperProtected, TamperProtectionSource
    
    # Check Credential Guard status
    Get-CimInstance -ClassName Win32_DeviceGuard | Select-Object SecurityServicesConfigured, SecurityServicesRunning
    
    # Check VBS status
    Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\DeviceGuard' | Select-Object EnableVirtualizationBasedSecurity, RequirePlatformSecurityFeatures
    
    # Check Device Health Attestation
    Get-Service -Name 'DeviceHealthAttestationService' -ErrorAction SilentlyContinue | Select-Object Status, StartType
    Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceHealthAttestation' -ErrorAction SilentlyContinue
    Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\DeviceHealthAttestation' -ErrorAction SilentlyContinue

.PARAMETER Quiet
    Suppress standard informational output (only warnings/errors). Still returns object if OutputMode supports it.

.PARAMETER OutputMode
    One of: Console (default), Plain, Object, Json

.PARAMETER OutputPath
    File path for JSON export when OutputMode Json. Defaults to ./defender-status.json if not provided.

.PARAMETER SkipOnlineCheck
    Skip web lookup of latest Defender signature / engine / platform versions (offline or faster execution).

.PARAMETER IncludeRaw
    Include raw Get-MpPreference and Get-MpComputerStatus objects in returned structured result.

.PARAMETER SkipSlowChecks
    Skips potentially slow queries like Hyper-V feature state (DISM online), Windows Update hotfix history, and SMB share enumeration.
    Alias: -Fast

.EXAMPLE
    .\win10-def-get.ps1
    Runs with colored console output.

.EXAMPLE
    .\win10-def-get.ps1 -OutputMode Object | ConvertTo-Json -Depth 4

.EXAMPLE
    .\win10-def-get.ps1 -OutputMode Json -OutputPath c:\temp\defender.json

.NOTES
    Requires: Windows 10/11 with Defender, PowerShell 5.1+ or PowerShell 7+, appropriate permissions.
    Some registry keys require elevation. Web lookup requires internet connectivity unless -SkipOnlineCheck is used.

.LINK
    https://learn.microsoft.com/microsoft-365/security/defender-endpoint/attack-surface-reduction-rules-reference
    Guide: https://github.com/directorcia/Office365/wiki/Windows-Security-Audit-Script
#>

param(
    [switch]$Quiet,
    [ValidateSet('Console','Plain','Object','Json','Csv','Html','Xml')]
    [string]$OutputMode = 'Console',
    [string]$OutputPath,
    [switch]$SkipOnlineCheck,
    [switch]$IncludeRaw,
    [Alias('Fast')][switch]$SkipSlowChecks
)

$FastMode = [bool]$SkipSlowChecks

## Script version
$ScriptVersion = '1.20'

## Colors (used only when OutputMode Console and host supports)
$systemmessagecolor = 'Cyan'
$processmessagecolor = 'Green'
$errormessagecolor = 'Red'
$warningmessagecolor = 'Yellow'

# Fallbacks if earlier run left them undefined (e.g., prior malformed block)
if (-not $systemmessagecolor) { $systemmessagecolor = 'Cyan' }
if (-not $processmessagecolor) { $processmessagecolor = 'Green' }
if (-not $errormessagecolor) { $errormessagecolor = 'Red' }
if (-not $warningmessagecolor) { $warningmessagecolor = 'Yellow' }


## Helper: Determine if colored output should be used (guard for non-interactive hosts)
try { $null = $Host.UI.RawUI.ForegroundColor; $HasColor = $true } catch { $HasColor = $false }
$UseColor = ($OutputMode -eq 'Console') -and $HasColor

function Write-Info {
    param([string]$Message,[string]$Color='Gray')
    if ($Quiet) { return }
    if ($UseColor) { Write-Host -ForegroundColor $Color $Message }
    else { Write-Host $Message }
}
function Write-Section {
    param([string]$Title)
    try {
        $consoleWidth = [Console]::WindowWidth
        $bar = ('-' * [Math]::Min($consoleWidth, 80))
    } catch {
        $bar = ('-' * 60)
    }
    if (-not $bar) { $bar = ('-' * 60) }
    $header = "[ $Title ]"
    if ($UseColor) {
        Write-Host ''
        Write-Host -ForegroundColor Cyan $bar
        Write-Host -ForegroundColor White $header
        Write-Host -ForegroundColor Cyan $bar
    } else {
        Write-Host "`n$bar"
        Write-Host $header
        Write-Host $bar
    }
}
function Write-Status {
    param(
        [string]$Label,
        [object]$Value,
        [ValidateSet('Good','Warn','Bad','Neutral')]
        [string]$State='Neutral'
    )
    if ($Quiet -and $State -eq 'Neutral') { return }
    $color = switch ($State) {
        'Good' { $processmessagecolor }
        'Warn' { $warningmessagecolor }
        'Bad'  { $errormessagecolor }
        default { 'White' }
    }
    $msg = "{0} = {1}" -f $Label, $Value
    if ($UseColor) { Write-Host -ForegroundColor $color $msg } else { Write-Host $msg }
}

    if (-not (Get-Command Write-KV -ErrorAction SilentlyContinue)) {
        function Write-KV {
            param(
                [Parameter(Mandatory)] [string]$Label,
                [object]$Value,
                [ValidateSet('Good','Warn','Bad','Neutral')] [string]$State='Neutral'
            )
            if ($Quiet -and $State -eq 'Neutral') { return }
            $displayValue = if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) { '(Not Set)' } else { $Value }
            $color = switch ($State) {
                'Good' { $processmessagecolor }
                'Warn' { $warningmessagecolor }
                'Bad'  { $errormessagecolor }
                default { 'White' }
            }
            $labelText = $Label.TrimEnd(':') + ':'
            $pad = if ($LabelWidth -gt $labelText.Length) { ' ' * ($LabelWidth - $labelText.Length) } else { ' ' }
            $msg = "{0}{1}{2}" -f $labelText, $pad, $displayValue
            if ($UseColor) { Write-Host -ForegroundColor $color $msg } else { Write-Host $msg }
        }
    }

function Test-Admin {
    try { $current = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent(); return $current.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch { return $false }
}

# Check admin rights and print status
$isAdmin = Test-Admin
if ($isAdmin) {
    Write-Info 'Session is running as administrator.' 'Green'
} else {
    Write-Status -Label 'Warning' -Value 'Session is NOT running as administrator. Some settings may be unavailable.' -State Warn
}

Write-Info "Script version $ScriptVersion starting`n"

## Clear-host

write-host -foregroundcolor $systemmessagecolor "Script started`n"

## ASR Rule Definitions (GUID -> Name)
$AsrRuleMap = @{
    'BE9BA2D9-53EA-4CDC-84E5-9B1EEEE46550' = 'Block executable content from email client and webmail'
    'D4F940AB-401B-4EFC-AADC-AD5F3C50688A' = 'Block all Office applications from creating child processes'
    '3B576869-A4EC-4529-8536-B80A7769E899' = 'Block Office applications from creating executable content'
    '75668C1F-73B5-4CF0-BB93-3ECF5CB7CC84' = 'Block Office applications from injecting code into other processes'
    'D3E037E1-3EB8-44C8-A917-57927947596D' = 'Block JavaScript or VBScript from launching downloaded executable content'
    '5BEB7EFE-FD9A-4556-801D-275E5FFC04CC' = 'Block execution of potentially obfuscated scripts'
    '92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B' = 'Block Win32 API calls from Office macros'
    '01443614-cd74-433a-b99e-2ecdc07bfc25' = 'Block executable files from running unless they meet criteria'
    'c1db55ab-c21a-4637-bb3f-a12568109d35' = 'Use advanced protection against ransomware'
    '9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2' = 'Block credential stealing from LSASS'
    'd1e49aac-8f56-4280-b9ba-993a6d77406c' = 'Block process creations from PSExec and WMI'
    'b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4' = 'Block untrusted / unsigned processes from USB'
    '26190899-1602-49e8-8b27-eb1d0a1ce869' = 'Block Office communication app from creating child processes'
    '7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c' = 'Block Adobe Reader from creating child processes'
    'e6db77e5-3df2-4cf1-b95a-636979351e5b' = 'Block persistence through WMI event subscription'
    'A8F5898E-1DC8-49A9-9878-85004B8A61E6' = 'Block Webshell creation for Servers'
    '56A863A9-875E-4185-98A7-B882C64B5CE5' = 'Block abuse of exploited vulnerable signed drivers'
    '33DDEDF1-C6E0-47CB-833E-DE6133960387' = 'Block rebooting machine in Safe Mode'
    'C0033C00-D16D-4114-A5A0-DC9B3A7D2CEB' = 'Block use of copied or impersonated system tools'
}

$asrrules = foreach ($k in $AsrRuleMap.Keys) { [PSCustomObject]@{ GUID = $k; Name = $AsrRuleMap[$k] } }

$enabledvalues = @{ 0='Not Enabled'; 1='Enabled'; 2='Audit'; 6='Warn' }
$displaycolor = $errormessagecolor, $processmessagecolor, $warningmessagecolor

$results = Get-MpPreference
Write-Section 'Attack Surface Reduction Rules'
$count = 0 

if ($results.AttackSurfaceReductionRules_ids) {
    $asrMap = @{}
    for ($i=0; $i -lt $results.AttackSurfaceReductionRules_ids.Count; $i++) {
        $rawId = $results.AttackSurfaceReductionRules_ids[$i]
        if ([string]::IsNullOrWhiteSpace($rawId)) { continue }
        $normId = ($rawId -replace '[{}]','').ToUpper()
        $stateRaw = $results.AttackSurfaceReductionRules_Actions[$i]
        # Normalize ASR state to integer code: 0=Not Enabled, 1=Enabled(Block), 2=Audit, 6=Warn
        $stateCode = 0
        if ($null -ne $stateRaw) {
            if ($stateRaw -is [int]) { $stateCode = [int]$stateRaw }
            elseif ($stateRaw -is [string]) {
                $s = $stateRaw.Trim().ToLower()
                switch ($s) {
                    '0' { $stateCode = 0 }
                    '1' { $stateCode = 1 }
                    '2' { $stateCode = 2 }
                    '6' { $stateCode = 6 }
                    'disabled' { $stateCode = 0 }
                    'not configured' { $stateCode = 0 }
                    'notconfigured' { $stateCode = 0 }
                    'not enabled' { $stateCode = 0 }
                    'enabled' { $stateCode = 1 }
                    'block' { $stateCode = 1 }
                    'audit' { $stateCode = 2 }
                    'auditmode' { $stateCode = 2 }
                    'warn' { $stateCode = 6 }
                    default {
                        try { $stateCode = [int]$stateRaw } catch { $stateCode = 0 }
                    }
                }
            } else {
                try { $stateCode = [int]$stateRaw } catch { $stateCode = 0 }
            }
        }
        $stateText = if ($enabledvalues.ContainsKey($stateCode)) { $enabledvalues[$stateCode] } else { "Unknown($stateCode)" }
        if ($asrMap.ContainsKey($normId)) {
            $prev = $asrMap[$normId]
            # Precedence: Enabled (1) > Warn (6) > Audit (2) > Not Enabled (0)
            $precedence = @{ 1=4; 6=3; 2=2; 0=1 }
            $prevPrec = if ($precedence.ContainsKey($prev.StateCode)) { $precedence[$prev.StateCode] } else { 0 }
            $newPrec  = if ($precedence.ContainsKey($stateCode)) { $precedence[$stateCode] } else { 0 }
            if ($newPrec -gt $prevPrec) {
                $prev.StateCode = $stateCode
                $prev.State = $stateText
            }
            $asrMap[$normId] = $prev
        } else {
            $name = $AsrRuleMap[$normId]
            if ([string]::IsNullOrWhiteSpace($name)) { $name = 'Unknown ASR rule' }
            $asrMap[$normId] = [PSCustomObject]@{
                GUID = $normId
                Name = $name
                StateCode = $stateCode
                State = $stateText
            }
        }
    }
    $AsrStatus = $asrMap.Values
    foreach ($r in $AsrStatus) {
        $stateCat = switch ($r.StateCode) { 1 { 'Good' } 6 { 'Warn' } 2 { 'Warn' } default { 'Bad' } }
        Write-Status -Label $r.Name -Value $r.State -State $stateCat
    }
} else {
    foreach ($r in $asrrules) { Write-Status -Label $r.Name -Value 'Not Enabled' -State Bad }
    $AsrStatus = $null
}
Write-Section 'Defender Settings'
$PUAState = if($results.PUAProtection -eq 1){'Good'} elseif($results.PUAProtection -eq 2){'Warn'} else {'Bad'}
Write-KV -Label 'Potentially Unwanted Application Protection' -Value $enabledvalues[$results.puaprotection] -State $PUAState
$ArchiveScanValue = if($results.DisableArchiveScanning){'Disabled'} else {'Enabled'}
$ArchiveScanState = if($results.DisableArchiveScanning){'Bad'} else {'Good'}
Write-KV -Label 'Scan archive files (.zip/.cab) for malware' -Value $ArchiveScanValue -State $ArchiveScanState
$AutoExclValue = if($results.DisableAutoExclusions){'Disabled'} else {'Enabled'}
$AutoExclState = if($results.DisableAutoExclusions){'Bad'} else {'Good'}
Write-KV -Label 'Automatic Exclusions feature for server' -Value $AutoExclValue -State $AutoExclState
$BehMonValue = if($results.DisableBehaviorMonitoring){'Disabled'} else {'Enabled'}
$BehMonState = if($results.DisableBehaviorMonitoring){'Bad'} else {'Good'}
Write-KV -Label 'Behavior Monitoring' -Value $BehMonValue -State $BehMonState
$BlockFirstSeenValue = if($results.DisableBlockAtFirstSeen){'Disabled'} else {'Enabled'}
$BlockFirstSeenState = if($results.DisableBlockAtFirstSeen){'Bad'} else {'Good'}
Write-KV -Label 'Block at first seen' -Value $BlockFirstSeenValue -State $BlockFirstSeenState
$PrivacyModeValue = if($results.DisablePrivacyMode){'Disabled'} else {'Enabled'}
$PrivacyModeState = if($results.DisablePrivacyMode){'Bad'} else {'Good'}
Write-KV -Label 'Privacy mode (hide threat history to non-admins)' -Value $PrivacyModeValue -State $PrivacyModeState
$UILockValue = if($results.UILockdown){'Disabled'} else {'Enabled'}
$UILockState = if($results.UILockdown){'Bad'} else {'Good'}
Write-KV -Label 'UI Lockdown Mode' -Value $UILockValue -State $UILockState
$netprotect = "Off", "On", "Audit"
$npState = switch ($results.EnableNetworkProtection) { 1 { 'Good' } 2 { 'Warn' } default { 'Bad' } }
Write-KV -Label 'Network protection' -Value $netprotect[$results.EnableNetworkProtection] -State $npState
Write-KV -Label 'Cmdlet throttle limit' -Value $results.ThrottleLimit -State Neutral
$submitconsent = "Always Prompt", "Send Safe sample automatically", "Never send", "Send all samples automatically"
Write-KV -Label 'Sample submission consent' -Value $submitconsent[$results.SubmitSamplesConsent] -State Neutral
if ($results.ThreatIDDefaultAction_Ids) {
    $tidIds = $results.ThreatIDDefaultAction_Ids -join ','
} else { $tidIds = $null }
if ($results.ThreatIDDefaultAction_Actions) {
    $tidActions = $results.ThreatIDDefaultAction_Actions -join ','
} else { $tidActions = $null }
Write-KV -Label 'Threat ID default actions (IDs)' -Value $tidIds -State Neutral
Write-KV -Label 'Threat ID default actions (Actions)' -Value $tidActions -State Neutral


<#  Scan settings   #>
Write-Section 'Scanning Settings'
${fsCatchVal} = if($results.DisableCatchupFullScan){'Disabled'} else {'Enabled'}; ${fsCatchState}= if($results.DisableCatchupFullScan){'Bad'} else {'Good'}
${qsCatchVal} = if($results.DisableCatchupQuickScan){'Disabled'} else {'Enabled'}; ${qsCatchState}= if($results.DisableCatchupQuickScan){'Bad'} else {'Good'}
${idleThrottleVal}= if($results.DisableCpuThrottleOnIdleScans){'Disabled'} else {'Enabled'}; ${idleThrottleState}= if($results.DisableCpuThrottleOnIdleScans){'Bad'} else {'Good'}
${emailScanVal}= if($results.DisableEmailScanning){'Disabled'} else {'Enabled'}; ${emailScanState}= if($results.DisableEmailScanning){'Bad'} else {'Good'}
${exploitProtVal}= if($results.DisableIntrusionPreventionSystem){'Disabled'} else {'Enabled'}; ${exploitProtState}= if($results.DisableIntrusionPreventionSystem){'Bad'} else {'Good'}
${ioavVal}= if($results.DisableIOAVProtection){'Disabled'} else {'Enabled'}; ${ioavState}= if($results.DisableIOAVProtection){'Bad'} else {'Good'}
${rtProtVal}= if($results.DisableRealtimeMonitoring){'Disabled'} else {'Enabled'}; ${rtProtState}= if($results.DisableRealtimeMonitoring){'Bad'} else {'Good'}
${remDriveVal}= if($results.DisableRemovableDriveScanning){'Disabled'} else {'Enabled'}; ${remDriveState}= if($results.DisableRemovableDriveScanning){'Bad'} else {'Good'}
${restorePointVal}= if($results.DisableRestorePoint){'Disabled'} else {'Enabled'}; ${restorePointState}= if($results.DisableRestorePoint){'Bad'} else {'Good'}
${mappedNetVal}= if($results.DisableScanningMappedNetworkDrivesForFullScan){'Disabled'} else {'Enabled'}; ${mappedNetState}= if($results.DisableScanningMappedNetworkDrivesForFullScan){'Bad'} else {'Good'}
${netFilesVal}= if($results.DisableScanningNetworkFiles){'Disabled'} else {'Enabled'}; ${netFilesState}= if($results.DisableScanningNetworkFiles){'Good'} else {'Warn'}
${scriptScanVal}= if($results.DisableScriptScanning){'Disabled'} else {'Enabled'}; ${scriptScanState}= if($results.DisableScriptScanning){'Bad'} else {'Good'}
Write-KV -Label 'Full scan catch up' -Value $fsCatchVal -State $fsCatchState
Write-KV -Label 'Quick scan catch up' -Value $qsCatchVal -State $qsCatchState
Write-KV -Label 'Throttle CPU on idle scans' -Value $idleThrottleVal -State $idleThrottleState
Write-KV -Label 'Email scanning' -Value $emailScanVal -State $emailScanState
Write-KV -Label 'Exploit network protection' -Value $exploitProtVal -State $exploitProtState
Write-KV -Label 'Downloaded files & attachments scan' -Value $ioavVal -State $ioavState
Write-KV -Label 'Real-time protection' -Value $rtProtVal -State $rtProtState
Write-KV -Label 'Removable drive scanning (full scan)' -Value $remDriveVal -State $remDriveState
Write-KV -Label 'Restore point scanning' -Value $restorePointVal -State $restorePointState
Write-KV -Label 'Mapped network drives scan' -Value $mappedNetVal -State $mappedNetState
Write-KV -Label 'Network files scan (discouraged)' -Value $netFilesVal -State $netFilesState
Write-KV -Label 'Script scanning' -Value $scriptScanVal -State $scriptScanState
Write-KV -Label 'Excluded extensions' -Value ($results.ExclusionExtension -join ',') -State Neutral
Write-KV -Label 'Excluded paths' -Value ($results.ExclusionPath -join ';') -State Neutral
Write-KV -Label 'Excluded processes' -Value ($results.Exclusionprocess -join ',') -State Neutral
$mapsrep = "Disabled","Basic Membership", "Advanced Membership"
Write-KV -Label 'MAPS membership' -Value $mapsrep[$results.mapsreporting] -State Neutral
Write-KV -Label 'Quarantine retention (days)' -Value $results.QuarantinePurgeItemsAfterDelay -State Neutral
${randTaskVal}= if($results.RandomizeScheduleTaskTimes){'Enabled'} else {'Disabled'}
Write-KV -Label 'Randomize scheduled task times' -Value $randTaskVal -State Neutral
$scandir = "Both", "Incoming", "Outgoing"
Write-KV -Label 'Real-time scan direction' -Value $scandir[$results.RealTimeScanDirection] -State Neutral
$weekday = "Everyday", "Sunday", "Monday", "Tuesday", "Wednesday","Thursday","Friday","Saturday","Never"
Write-KV -Label 'Remediation scheduled full scan day' -Value $weekday[$results.RealTimeScanDirection] -State Neutral
Write-KV -Label 'Remediation scheduled scan time (mins after midnight)' -Value $results.RemediationScheduleTime -State Neutral
Write-KV -Label 'Reporting additional action timeout (mins)' -Value $results.ReportingAdditionalActionTimeOut -State Neutral
Write-KV -Label 'Reporting critical failure timeout (mins)' -Value $results.ReportingCriticalFailureTimeOut -State Neutral
Write-KV -Label 'Reporting non-critical timeout (mins)' -Value $results.ReportingNonCriticalTimeOut -State Neutral
Write-KV -Label 'Max CPU percent for scan' -Value $results.ScanAvgCPULoadFactor -State Neutral
${idleOnlyVal}= if($results.ScanOnlyIfIdleEnabled){'Enabled'} else {'Disabled'}; ${idleOnlyState}= if($results.ScanOnlyIfIdleEnabled){'Good'} else {'Warn'}
Write-KV -Label 'Scheduled scan only if idle' -Value $idleOnlyVal -State $idleOnlyState
$scantype = "Quick scan", "Full scan"
Write-KV -Label 'Scheduled scan type' -Value $scantype[$results.ScanParameters] -State Neutral
Write-KV -Label 'Scan history retention (days)' -Value $results.ScanPurgeItemsAfterDelay -State Neutral
Write-KV -Label 'Scheduled scan day' -Value $weekday[$results.ScanScheduleDay] -State Neutral
Write-KV -Label 'Scheduled quick scan time (mins after midnight)' -Value $results.ScanScheduleQuickScanTime -State Neutral
Write-KV -Label 'Signature grace period (mins)' -Value $results.SignatureAuGracePeriod -State Neutral
Write-KV -Label 'Definition update file share sources' -Value ($results.SignatureDefinitionUpdateFileSharesSources -join ',') -State Neutral

<#          Threat Level        #>
Write-Section 'Remediation Actions'
$hta = "NoAction", "Clean", "Quarantine", "Remove", "Allow", "UserDefined", "Block"
$remediationvalue = @(
    'Apply action based on SIU (default)',
    'Clean the detected threat',
    'Quarantine the detected threat',
    'Remove the detected threat',
    'Allow (not recommended)',
    'User defined',
    'Block the detected threat'
)
Write-KV -Label 'Low threat default action' -Value ("{0} ({1})" -f $hta[$results.LowThreatDefaultAction],$remediationvalue[$results.LowThreatDefaultAction]) -State Neutral
Write-KV -Label 'Moderate threat default action' -Value ("{0} ({1})" -f $hta[$results.ModerateThreatDefaultAction],$remediationvalue[$results.ModerateThreatDefaultAction]) -State Neutral
Write-KV -Label 'High threat default action' -Value ("{0} ({1})" -f $hta[$results.HighThreatDefaultAction],$remediationvalue[$results.HighThreatDefaultAction]) -State Neutral
Write-KV -Label 'Severe threat default action' -Value ("{0} ({1})" -f $hta[$results.SevereThreatDefaultAction],$remediationvalue[$results.SevereThreatDefaultAction]) -State Neutral
Write-KV -Label 'Unknown threat default action' -Value ("{0} ({1})" -f $hta[$results.UnknownThreatDefaultAction],$remediationvalue[$results.UnknownThreatDefaultAction]) -State Neutral

<#      Updates         #>
Write-Section 'Update Settings'
${updOnStartVal}= if($results.SignatureDisableUpdateOnStartupWithoutEngine){'Disabled'} else {'Enabled'}; ${updOnStartState}= if($results.SignatureDisableUpdateOnStartupWithoutEngine){'Bad'} else {'Good'}
Write-KV -Label 'Update on startup without engine' -Value $updOnStartVal -State $updOnStartState
Write-KV -Label 'Definition source fallback order' -Value $results.SignatureFallbackOrder -State Neutral
Write-KV -Label 'Definition first grace period (mins)' -Value $results.SignatureFirstAuGracePeriod -State Neutral
Write-KV -Label 'Definition update schedule day' -Value $weekday[$results.SignatureScheduleDay] -State Neutral
Write-KV -Label 'Definition update schedule time (mins after midnight)' -Value $results.SignatureScheduleTime -State Neutral
Write-KV -Label 'Definition catch-up interval (days)' -Value $results.SignatureUpdateCatchupInterval -State Neutral
Write-KV -Label 'Definition update interval (hours)' -Value $results.SignatureUpdateInterval -State Neutral

Write-Section 'Latest Signature / Engine Versions'
## https://docs.microsoft.com/en-us/previous-versions/windows/desktop/defender/msft-mpcomputerstatus#properties
$localdefender = Get-MpComputerStatus
write-host -foregroundcolor $processmessagecolor "Look up web site"
if (-not $SkipOnlineCheck) {
    try {
        Write-Info 'Querying Microsoft for latest signature versions...'
        $info = Invoke-WebRequest -Uri 'https://www.microsoft.com/en-us/wdsi/defenderupdates' -UseBasicParsing -ErrorAction Stop
        $raw = $info.RawContent
        $null = $raw -match '<li>Version: <span>.*'
        $LatestSig = ($Matches.Values -replace '<li>Version: <span>','' -replace '</span></li>','').Trim()
        $null = $raw -match '<li>Engine version: <span>.*'
        $LatestEngine = ($Matches.Values -replace '<li>Engine Version: <span>','' -replace '</span></li>','').Trim()
        $null = $raw -match '<li>Platform version: <span>.*'
        $LatestPlatform = ($Matches.Values -replace '<li>Platform Version: <span>','' -replace '</span></li>','').Trim()
        $null = $raw -match '<li>Released: <span id=.*'
        $LatestReleaseDate = ($Matches.Values -replace '<li>Released: <span id="dateofrelease">','' -replace '</span></li>','').Trim()
    } catch {
        Write-Status -Label 'Online Version Lookup' -Value "Failed: $($_.Exception.Message)" -State Warn
    }
} else {
    Write-Info 'Skipping online version lookup (-SkipOnlineCheck)'
}

if ($LatestSig) {
    $SigState = if ($localdefender.AntispywareSignatureVersion -eq $LatestSig) { 'Good' } else { 'Warn' }
    Write-Status -Label 'Signature Version' -Value ("{0} (Latest: {1})" -f $localdefender.AntispywareSignatureVersion,$LatestSig) -State $SigState
} else { Write-Status -Label 'Signature Version' -Value $localdefender.AntispywareSignatureVersion }
if ($LatestEngine) {
    $EngineState = if ($localdefender.AMEngineVersion -eq $LatestEngine) { 'Good' } else { 'Warn' }
    Write-Status -Label 'Engine Version' -Value ("{0} (Latest: {1})" -f $localdefender.AMEngineVersion,$LatestEngine) -State $EngineState
} else { Write-Status -Label 'Engine Version' -Value $localdefender.AMEngineVersion }
if ($LatestPlatform) {
    $PlatformState = if ($localdefender.AMServiceVersion -eq $LatestPlatform) { 'Good' } else { 'Warn' }
    Write-Status -Label 'Platform Version' -Value ("{0} (Latest: {1})" -f $localdefender.AMServiceVersion,$LatestPlatform) -State $PlatformState
} else { Write-Status -Label 'Platform Version' -Value $localdefender.AMServiceVersion }
if ($LatestReleaseDate) { Write-Status -Label 'Latest Release Date' -Value $LatestReleaseDate }
Write-KV -Label 'Last signature update time' -Value $localdefender.AntivirusSignatureLastUpdated -State Neutral
Write-KV -Label 'Anti-Malware mode' -Value $localdefender.AMRunningMode -State Neutral
${amSvcState}= if($localdefender.AMServiceEnabled){'Good'} else {'Bad'}
${asSvcState}= if($localdefender.AntispywareEnabled){'Good'} else {'Bad'}
${avSvcState}= if($localdefender.AntivirusEnabled){'Good'} else {'Bad'}
${behMonState}= if($localdefender.BehaviorMonitorEnabled){'Good'} else {'Bad'}
${ioavState2}= if($localdefender.IoavProtectionEnabled){'Good'} else {'Bad'}
${tamperState}= if($localdefender.IsTamperProtected){'Good'} else {'Bad'}
${nriState}= if($localdefender.NISEnabled){'Good'} else {'Bad'}
${onAccessState}= if($localdefender.OnAccessProtectionEnabled){'Good'} else {'Bad'}
${rtState}= if($localdefender.RealTimeProtectionEnabled){'Good'} else {'Bad'}
Write-KV -Label 'Anti-Malware service enabled' -Value $localdefender.AMServiceEnabled -State $amSvcState
Write-KV -Label 'Anti-Spyware service enabled' -Value $localdefender.AntispywareEnabled -State $asSvcState
Write-KV -Label 'Anti-Virus service enabled' -Value $localdefender.AntivirusEnabled -State $avSvcState
Write-KV -Label 'Behavior monitoring enabled' -Value $localdefender.BehaviorMonitorEnabled -State $behMonState
Write-KV -Label 'Download/attachment scanning enabled' -Value $localdefender.IoavProtectionEnabled -State $ioavState2
Write-KV -Label 'Tamper protection enabled' -Value $localdefender.IsTamperProtected -State $tamperState
Write-KV -Label 'NRI engine enabled' -Value $localdefender.NISEnabled -State $nriState
Write-KV -Label 'On-access protection enabled' -Value $localdefender.OnAccessProtectionEnabled -State $onAccessState
Write-KV -Label 'Real-time protection enabled' -Value $localdefender.RealTimeProtectionEnabled -State $rtState

<#      Device Guard        #>

Write-Section 'Defender & Security Features'
# Firewall status
try {
    $fwProfiles = Get-NetFirewallProfile -ErrorAction Stop
    foreach ($profile in $fwProfiles) {
        $fwState = if ($profile.Enabled) { 'Enabled' } else { 'Disabled' }
        $fwStateCat = if ($profile.Enabled) { 'Good' } else { 'Warn' }
        Write-KV -Label ("Firewall ({0})" -f $profile.Name) -Value $fwState -State $fwStateCat
    }
} catch { Write-KV -Label 'Firewall Status' -Value 'Unknown' -State Warn }

# Enhanced Exploit Protection Detection
try {
    $epStatus = 'Unknown'
    $epDetails = @()
    
    # Method 1: Check system-wide process mitigation settings
    try {
        $epSettings = Get-ProcessMitigation -System -ErrorAction Stop
        $mitigations = @()
        
        # Check DEP (Data Execution Prevention)
        if ($epSettings.Dep -and $epSettings.Dep.Enable) {
            $mitigations += 'DEP'
        }
        
        # Check ASLR (Address Space Layout Randomization)
        if ($epSettings.Aslr -and $epSettings.Aslr.ForceRelocateImages) {
            $mitigations += 'ASLR'
        }
        
        # Check CFG (Control Flow Guard)
        if ($epSettings.CFG -and $epSettings.CFG.Enable) {
            $mitigations += 'CFG'
        }
        
        # Check SEHOP (Structured Exception Handler Overwrite Protection)
        if ($epSettings.SEHOP -and $epSettings.SEHOP.Enable) {
            $mitigations += 'SEHOP'
        }
        
        # Check Heap Protection
        if ($epSettings.Heap -and $epSettings.Heap.TerminateOnHeapErrors) {
            $mitigations += 'Heap'
        }
        
        if ($mitigations.Count -gt 0) {
            $epStatus = 'Enabled'
            $epDetails += "Mitigations: $($mitigations -join ', ')"
        } else {
            $epStatus = 'Limited'
            $epDetails += 'System: Basic protections only'
        }
    } catch {
        # Fallback to registry check
        $epReg = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\kernel' -ErrorAction SilentlyContinue
        if ($epReg -and $epReg.PSObject.Properties.Name -contains 'DisableExceptionChainValidation') {
            $sehop = if ($epReg.DisableExceptionChainValidation -eq 0) { 'Enabled' } else { 'Disabled' }
            $epDetails += "SEHOP: $sehop"
            $epStatus = if ($sehop -eq 'Enabled') { 'Partial' } else { 'Limited' }
        }
    }
    
    # Method 2: Check Windows Defender Exploit Guard settings
    try {
        $egReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Windows Defender Exploit Guard\Exploit Protection\Settings' -ErrorAction SilentlyContinue
        if ($egReg) {
            $epDetails += 'Defender EG: Configured'
            if ($epStatus -eq 'Unknown') {
                $epStatus = 'Configured'
            }
        }
    } catch { }
    
    # Method 3: Check process-specific mitigations
    try {
        $processCount = (Get-ProcessMitigation -RegistryConfigFilePath $null -ErrorAction SilentlyContinue | Measure-Object).Count
        if ($processCount -gt 0) {
            $epDetails += "Process Rules: $processCount configured"
        }
    } catch { }
    
    # Display result
    $epState = switch ($epStatus) {
        'Enabled' { 'Good' }
        'Configured' { 'Good' }
        'Partial' { 'Warn' }
        'Limited' { 'Warn' }
        default { 'Warn' }
    }
    
    # Main status line
    Write-KV -Label 'Exploit Protection' -Value $epStatus -State $epState
    
    # Display details as sub-items
    if ($epDetails.Count -gt 0) {
        foreach ($detail in $epDetails) {
            $detailParts = $detail -split ': ', 2
            if ($detailParts.Count -eq 2) {
                $detailLabel = "  └─ $($detailParts[0])"
                $detailValue = $detailParts[1]
                Write-KV -Label $detailLabel -Value $detailValue -State Neutral
            }
        }
    }
} catch { 
    Write-KV -Label 'Exploit Protection' -Value "Error: $($_.Exception.Message)" -State Warn 
}

# Enhanced Controlled Folder Access Detection
try {
    $cfaStatus = 'Disabled'
    $cfaDetails = @()
    
    # Method 1: Get-MpPreference (primary method)
    $prefs = Get-MpPreference
    $cfa = $null
    if ($prefs.PSObject.Properties.Name -contains 'EnableControlledFolderAccess') {
        $cfa = $prefs.EnableControlledFolderAccess
    } elseif ($prefs.PSObject.Properties.Name -contains 'ControlledFolderAccess') {
        $cfa = $prefs.ControlledFolderAccess
    }
    
    $cfaStatus = switch ($cfa) { 
        1 { 'Enabled' } 
        2 { 'Audit Mode' } 
        default { 'Disabled' } 
    }
    
    # Method 2: Check protected folders count
    if ($cfa -and $cfa -gt 0) {
        try {
            $protectedFolders = $prefs.ControlledFolderAccessProtectedFolders
            if ($protectedFolders) {
                $folderCount = ($protectedFolders | Measure-Object).Count
                $cfaDetails += "Protected Folders: $folderCount"
            } else {
                $cfaDetails += 'Protected Folders: Default system folders'
            }
        } catch { }
        
        # Check allowed applications
        try {
            $allowedApps = $prefs.ControlledFolderAccessAllowedApplications
            if ($allowedApps) {
                $appCount = ($allowedApps | Measure-Object).Count
                $cfaDetails += "Allowed Apps: $appCount"
            }
        } catch { }
    }
    
    # Method 3: Registry fallback check
    if ($cfa -eq $null) {
        try {
            $cfaReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Windows Defender Exploit Guard\Controlled Folder Access' -ErrorAction SilentlyContinue
            if ($cfaReg -and $cfaReg.PSObject.Properties.Name -contains 'EnableControlledFolderAccess') {
                $regValue = $cfaReg.EnableControlledFolderAccess
                $cfaStatus = switch ($regValue) { 
                    1 { 'Enabled' } 
                    2 { 'Audit Mode' } 
                    default { 'Disabled' } 
                }
                $cfaDetails += 'Source: Registry'
            }
        } catch { }
    }
    
    # Method 4: Check policy configuration
    try {
        $cfaPolicy = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\Controlled Folder Access' -ErrorAction SilentlyContinue
        if ($cfaPolicy) {
            $cfaDetails += 'Policy: Configured'
        }
    } catch { }
    
    # Display result
    $cfaState = switch ($cfaStatus) { 
        'Enabled' { 'Good' } 
        'Audit Mode' { 'Warn' } 
        default { 'Warn' } 
    }
    
    # Main status line
    Write-KV -Label 'Controlled Folder Access' -Value $cfaStatus -State $cfaState
    
    # Display details as sub-items
    if ($cfaDetails.Count -gt 0) {
        foreach ($detail in $cfaDetails) {
            $detailParts = $detail -split ': ', 2
            if ($detailParts.Count -eq 2) {
                $detailLabel = "  └─ $($detailParts[0])"
                $detailValue = $detailParts[1]
                Write-KV -Label $detailLabel -Value $detailValue -State Neutral
            }
        }
    }
} catch { 
    Write-KV -Label 'Controlled Folder Access' -Value "Error: $($_.Exception.Message)" -State Warn 
}

# Enhanced Ransomware Protection Detection
try {
    $ransomStatus = 'Disabled'
    $ransomDetails = @()
    
    # Method 1: Controlled Folder Access (primary ransomware protection)
    $prefs = Get-MpPreference
    $cfa = $null
    if ($prefs.PSObject.Properties.Name -contains 'EnableControlledFolderAccess') {
        $cfa = $prefs.EnableControlledFolderAccess
    } elseif ($prefs.PSObject.Properties.Name -contains 'ControlledFolderAccess') {
        $cfa = $prefs.ControlledFolderAccess
    }
    
    $cfaStatus = switch ($cfa) { 
        1 { 'Enabled' } 
        2 { 'Audit Mode' } 
        default { 'Disabled' } 
    }
    
    if ($cfa -and $cfa -gt 0) {
        $ransomStatus = $cfaStatus
        $ransomDetails += "Controlled Folder Access: $cfaStatus"
    }
    
    # Method 2: Check ASR rules specifically targeting ransomware
    $ransomwareAsrRules = @{
        'c1db55ab-c21a-4637-bb3f-a12568109d35' = 'Advanced Ransomware Protection'
        '9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2' = 'Credential Stealing Protection'
        'd1e49aac-8f56-4280-b9ba-993a6d77406c' = 'Process Creation Protection'
    }
    
    $asrRansomProtection = @()
    if ($results.AttackSurfaceReductionRules_ids) {
        for ($i = 0; $i -lt $results.AttackSurfaceReductionRules_ids.Count; $i++) {
            $ruleId = ($results.AttackSurfaceReductionRules_ids[$i] -replace '[{}]', '').ToUpper()
            $ruleAction = $results.AttackSurfaceReductionRules_Actions[$i]
            
            if ($ransomwareAsrRules.ContainsKey($ruleId.ToLower()) -and $ruleAction -eq 1) {
                $asrRansomProtection += $ransomwareAsrRules[$ruleId.ToLower()]
            }
        }
    }
    
    if ($asrRansomProtection.Count -gt 0) {
        $ransomDetails += "ASR Anti-Ransomware: $($asrRansomProtection.Count) rules enabled"
        if ($ransomStatus -eq 'Disabled') {
            $ransomStatus = 'Partial (ASR Only)'
        }
    }
    
    # Method 3: Check OneDrive Files Restore capability
    try {
        $oneDriveReg = Get-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\OneDrive\Accounts\*' -ErrorAction SilentlyContinue
        if ($oneDriveReg) {
            $ransomDetails += 'OneDrive: Files Restore available'
        }
    } catch { }
    
    # Method 4: Check Windows Backup status
    try {
        $backupReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsBackup' -ErrorAction SilentlyContinue
        if ($backupReg) {
            $ransomDetails += 'Windows Backup: Configured'
        }
    } catch { }
    
    # Method 5: Check File History
    try {
        $fileHistoryReg = Get-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\FileHistory' -ErrorAction SilentlyContinue
        if ($fileHistoryReg -and $fileHistoryReg.PSObject.Properties.Name -contains 'Enabled' -and $fileHistoryReg.Enabled -eq 1) {
            $ransomDetails += 'File History: Enabled'
        }
    } catch { }
    
    # Overall status determination
    if ($ransomStatus -eq 'Disabled' -and $ransomDetails.Count -gt 1) {
        $ransomStatus = 'Basic Protection'
    }
    
    # Display result
    $ransomState = switch ($ransomStatus) { 
        'Enabled' { 'Good' }
        'Audit Mode' { 'Warn' }  
        'Partial (ASR Only)' { 'Warn' }
        'Basic Protection' { 'Warn' }
        default { 'Bad' } 
    }
    
    # Main status line
    Write-KV -Label 'Ransomware Protection' -Value $ransomStatus -State $ransomState
    
    # Display details as sub-items
    if ($ransomDetails.Count -gt 0) {
        foreach ($detail in $ransomDetails) {
            $detailParts = $detail -split ': ', 2
            if ($detailParts.Count -eq 2) {
                $detailLabel = "  └─ $($detailParts[0])"
                $detailValue = $detailParts[1]
                Write-KV -Label $detailLabel -Value $detailValue -State Neutral
            }
        }
    }
} catch { 
    Write-KV -Label 'Ransomware Protection' -Value "Error: $($_.Exception.Message)" -State Warn 
}

# Enhanced Cloud-delivered Protection Detection
try {
    $cloudStatus = 'Disabled'
    $cloudDetails = @()
    
    # Method 1: MAPS Reporting level
    try {
        $cloud = Get-MpPreference | Select-Object -ExpandProperty MAPSReporting -ErrorAction Stop
        $cloudStatus = switch ($cloud) { 
            2 { 'Advanced' } 
            1 { 'Basic' } 
            default { 'Disabled' } 
        }
        $cloudDetails += "MAPS Reporting: $cloudStatus"
    } catch {
        # Fallback to direct property check
        $prefs = Get-MpPreference
        if ($prefs.PSObject.Properties.Name -contains 'MAPSReporting') {
            $cloud = $prefs.MAPSReporting
            $cloudStatus = switch ($cloud) { 
                2 { 'Advanced' } 
                1 { 'Basic' } 
                default { 'Disabled' } 
            }
            $cloudDetails += "MAPS Reporting: $cloudStatus"
        }
    }
    
    # Method 2: Cloud block level
    try {
        $prefs = Get-MpPreference
        if ($prefs.PSObject.Properties.Name -contains 'CloudBlockLevel') {
            $blockLevel = $prefs.CloudBlockLevel
            $blockLevelText = switch ($blockLevel) {
                0 { 'Default' }
                2 { 'High' }
                4 { 'High+' }
                6 { 'Zero Tolerance' }
                default { "Level $blockLevel" }
            }
            $cloudDetails += "Block Level: $blockLevelText"
        }
    } catch { }
    
    # Method 3: Cloud extended timeout
    try {
        $prefs = Get-MpPreference
        if ($prefs.PSObject.Properties.Name -contains 'CloudExtendedTimeout') {
            $timeout = $prefs.CloudExtendedTimeout
            if ($timeout -gt 0) {
                $cloudDetails += "Extended Timeout: ${timeout}s"
            }
        }
    } catch { }
    
    # Method 4: First seen file blocking
    try {
        $prefs = Get-MpPreference
        if ($prefs.PSObject.Properties.Name -contains 'DisableBlockAtFirstSeen') {
            $blockFirstSeen = -not $prefs.DisableBlockAtFirstSeen
            if ($blockFirstSeen) {
                $cloudDetails += 'Block at First Seen: Enabled'
            }
        }
    } catch { }
    
    # Method 5: Check registry for additional cloud settings
    try {
        $cloudReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Spynet' -ErrorAction SilentlyContinue
        if ($cloudReg) {
            if ($cloudReg.PSObject.Properties.Name -contains 'SpynetReporting') {
                $spynetLevel = $cloudReg.SpynetReporting
                if ($spynetLevel -gt 0 -and $cloudStatus -eq 'Disabled') {
                    $cloudStatus = 'Basic (Registry)'
                    $cloudDetails += 'Source: Registry configuration'
                }
            }
        }
    } catch { }
    
    # Method 6: Check connectivity to Microsoft cloud
    try {
        $connectivityReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Features' -ErrorAction SilentlyContinue
        if ($connectivityReg -and $connectivityReg.PSObject.Properties.Name -contains 'TamperProtectionSource' -and $connectivityReg.TamperProtectionSource -eq 1) {
            $cloudDetails += 'Cloud Connectivity: Active'
        }
    } catch { }
    
    # Display result
    $cloudState = switch ($cloudStatus) { 
        'Advanced' { 'Good' }
        'Basic' { 'Good' }
        'Basic (Registry)' { 'Good' }
        default { 'Warn' } 
    }
    
    # Main status line
    Write-KV -Label 'Cloud-delivered Protection' -Value $cloudStatus -State $cloudState
    
    # Display details as sub-items
    if ($cloudDetails.Count -gt 0) {
        foreach ($detail in $cloudDetails) {
            $detailParts = $detail -split ': ', 2
            if ($detailParts.Count -eq 2) {
                $detailLabel = "  └─ $($detailParts[0])"
                $detailValue = $detailParts[1]
                Write-KV -Label $detailLabel -Value $detailValue -State Neutral
            }
        }
    }
} catch { 
    Write-KV -Label 'Cloud-delivered Protection' -Value "Error: $($_.Exception.Message)" -State Warn 
}

# Automatic Sample Submission
try {
    $sample = Get-MpPreference | Select-Object -ExpandProperty SubmitSamplesConsent
    $sampleLevels = @( 'Always Prompt', 'Send Safe samples automatically', 'Never send', 'Send all samples automatically' )
    $sampleState = $sampleLevels[$sample]
    Write-KV -Label 'Automatic Sample Submission' -Value $sampleState -State Neutral
} catch { Write-KV -Label 'Automatic Sample Submission' -Value 'Unknown' -State Warn }

# Security Intelligence Channel
try {
    $prefs = Get-MpPreference
    $sigChannel = $null
    if ($prefs.PSObject.Properties.Name -contains 'SignatureUpdateChannel') {
        $sigChannel = $prefs.SignatureUpdateChannel
        $channelState = switch ($sigChannel) { 0 { 'Broad' } 1 { 'Current' } 2 { 'Staged' } default { $sigChannel } }
        Write-KV -Label 'Security Intelligence Channel' -Value $channelState -State Neutral
    } else {
        Write-KV -Label 'Security Intelligence Channel' -Value '(Not Set)' -State Neutral
    }
} catch { Write-KV -Label 'Security Intelligence Channel' -Value '(Not Set)' -State Neutral }

# Threat History (last 5 events)
try {
    $threats = Get-MpThreatDetection -ErrorAction SilentlyContinue | Sort-Object -Property InitialDetectionTime -Descending | Select-Object -First 5
    foreach ($t in $threats) {
        Write-KV -Label ("Threat: {0}" -f $t.ThreatName) -Value $t.ActionTaken -State (if($t.ActionTaken -eq 'Blocked'){'Good'}else{'Warn'})
    }
} catch { Write-KV -Label 'Threat History' -Value 'Unavailable' -State Neutral }

# Quarantine Items (count)
try {
    $quarantine = Get-MpThreatDetection -ErrorAction SilentlyContinue | Where-Object { $_.ActionTaken -eq 'Quarantined' }
    $qCount = ($quarantine | Measure-Object).Count
    Write-KV -Label 'Quarantine Items' -Value $qCount -State (if($qCount -gt 0){'Warn'}else{'Good'})
} catch { Write-KV -Label 'Quarantine Items' -Value 'Unknown' -State Warn }

Write-Section 'Platform Security'
$dg = Get-ItemProperty -Path HKLM:\System\CurrentControlSet\Control\deviceguard
$dgEnabled = $false
if (-not [string]::IsNullOrEmpty($dg.EnableVirtualizationBasedSecurity)) {
    if ($dg.EnableVirtualizationBasedSecurity -eq 1) { $dgEnabled = $true }
}
${dgVbsVal}= if($dgEnabled){'Enabled'} else {'Disabled'}; ${dgVbsState}= if($dgEnabled){'Good'} else {'Warn'}
Write-KV -Label 'Device Guard virtualization-based security' -Value $dgVbsVal -State $dgVbsState

# Actual Secure Boot state (independent of VBS requirements)
$sbText = $null; $sbState = 'Warn'
try {
    $sb = Confirm-SecureBootUEFI -ErrorAction Stop
    if ($sb -is [bool]) {
        $sbText = if ($sb) { 'On' } else { 'Off' }
        $sbState = if ($sb) { 'Good' } else { 'Warn' }
    }
} catch {
    try {
        $sbReg = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State' -ErrorAction Stop
        $sbVal = $sbReg.UEFISecureBootEnabled
        if ($null -ne $sbVal) {
            $sbText = if ($sbVal -eq 1) { 'On' } else { 'Off' }
            $sbState = if ($sbVal -eq 1) { 'Good' } else { 'Warn' }
        }
    } catch { }
}
Write-KV -Label 'Secure Boot state' -Value $sbText -State $sbState

# BIOS Mode (UEFI or Legacy)
$biosMode = $null
try {
    $fw = (Get-ItemProperty -Path 'HKLM:hardware\description\system' -ErrorAction Stop)
    if ($fw.PSObject.Properties.Name -contains 'PEFirmwareType') {
        switch ([int]$fw.PEFirmwareType) {
            0 { $biosMode = 'Unknown' }
            1 { $biosMode = 'Legacy' }
            2 { $biosMode = 'UEFI' }
            3 { $biosMode = 'UEFI-CSM' }
            default { $biosMode = "Unknown ($($fw.PEFirmwareType))" }
        }
    }
} catch { }
if (-not $biosMode) {
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $biosMode = if ($cs.BootupState -match 'EFI|UEFI') { 'UEFI' } else { 'Legacy' }
    } catch { $biosMode = 'Unknown' }
}
Write-KV -Label 'BIOS Mode' -Value $biosMode -State Neutral

# PCR7 Configuration (best-effort to mirror msinfo32 wording)
$pcr7 = 'Not Supported'
try {
    $isUefi = ($biosMode -eq 'UEFI' -or $biosMode -eq 'UEFI-CSM')
    $hasTpm = $false
    try {
        $tpm = Get-Tpm -ErrorAction Stop
        if ($tpm -and $tpm.TpmPresent) { $hasTpm = $true }
    } catch { }
    $isGpt = $false
    try {
        $sysDrive = Get-Partition -DriveLetter ((Get-Location).Path.Substring(0,1)) -ErrorAction SilentlyContinue
        if (-not $sysDrive) { $sysDrive = Get-Partition | Where-Object { $_.IsBoot -or $_.IsSystem } | Select-Object -First 1 }
        if ($sysDrive) { $isGpt = ($sysDrive.GptType -ne $null -or $sysDrive.Guid -ne $null) }
    } catch { }
    if (-not $isUefi) { $pcr7 = 'Binding Not Possible' }
    elseif (-not $hasTpm) { $pcr7 = 'Binding Not Possible' }
    elseif ($sbText -ne 'On') { $pcr7 = 'Binding Not Possible' }
    else { $pcr7 = 'Bound' }
} catch { $pcr7 = 'Elevation Required to View' }
Write-KV -Label 'PCR7 Configuration' -Value $pcr7 -State Neutral

# TPM details
$tpmPresent = $false
$tpmSpec = $null
try {
    $tpm = Get-Tpm -ErrorAction Stop
    if ($tpm) {
        $tpmPresent = [bool]$tpm.TpmPresent
        # Try to infer spec version: use TpmVersion if present, otherwise TPM 2.0 if Ready and PhysicalPresenceVersionInfo >= 2
        $tpmSpec = $null
        if ($tpm.PSObject.Properties.Name -contains 'SpecVersion') { $tpmSpec = $tpm.SpecVersion }
        elseif ($tpm.PSObject.Properties.Name -contains 'ManufacturerVersion') { $tpmSpec = $tpm.ManufacturerVersion }
        elseif ($tpm.TpmPresent) { $tpmSpec = if ($tpm.ManagedAuthLevel -or $tpm.IsActivated_InitialValue) { '2.0 (inferred)' } else { '1.2/Unknown' } }
    }
} catch {
    try {
        $tpmWmi = Get-WmiObject -Namespace 'root\\CIMV2\\Security\\MicrosoftTpm' -Class Win32_Tpm -ErrorAction Stop
        if ($tpmWmi) {
            $tpmPresent = [bool]$tpmWmi.IsEnabled_InitialValue -or [bool]$tpmWmi.IsActivated_InitialValue -or [bool]$tpmWmi.IsOwned_InitialValue
            $tpmSpec = if ($tpmWmi.SpecVersion) { $tpmWmi.SpecVersion } else { if ($tpmPresent) { 'Unknown' } else { $null } }
        }
    } catch { }
}
${tpmPresentState} = if($tpmPresent){'Good'} else {'Warn'}
Write-KV -Label 'TPM Present' -Value $tpmPresent -State $tpmPresentState
Write-KV -Label 'TPM Spec Version' -Value $tpmSpec -State Neutral

# TCG event log availability (measured boot log)
$tcgAvailable = $false
try {
    $paths = @(
        '$env:windir\\Logs\\MeasuredBoot\\',
        '$env:windir\\System32\\Logs\\MeasuredBoot\\'
    )
    foreach ($p in $paths) {
        $expanded = Invoke-Expression -Command ("`"$p`"")
        if (Test-Path -Path $expanded) {
            $files = Get-ChildItem -Path $expanded -File -ErrorAction SilentlyContinue | Where-Object { $_.Length -gt 0 }
            if ($files) { $tcgAvailable = $true; break }
        }
    }
} catch { }
${tcgLogState} = if($tcgAvailable){'Good'} else {'Warn'}
Write-KV -Label 'TCG Log Available' -Value $tcgAvailable -State $tcgLogState

# BitLocker Status
try {
    $bitlocker = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
    $blStatus = $bitlocker.ProtectionStatus
    $blEncMethod = $bitlocker.EncryptionMethod
    $blState = switch ($blStatus) { 'On' { 'Good' } 'Off' { 'Warn' } default { 'Neutral' } }
    Write-KV -Label 'BitLocker Protection' -Value $blStatus -State $blState
    Write-KV -Label 'BitLocker Encryption Method' -Value $blEncMethod -State Neutral
} catch { Write-KV -Label 'BitLocker Status' -Value 'Unknown' -State Warn }

# Windows Hello Status  
try {
    $helloReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Biometrics' -ErrorAction SilentlyContinue
    $helloEnabled = if ($helloReg -and $helloReg.Enabled -eq 0) { 'Disabled' } else { 'Available' }
    Write-KV -Label 'Windows Hello' -Value $helloEnabled -State Neutral
} catch { Write-KV -Label 'Windows Hello' -Value 'Unknown' -State Neutral }

# Virtualization Status
if ($FastMode) {
    if (-not $Quiet) { Write-Info 'Skipping Hyper-V feature check due to -Fast/-SkipSlowChecks' }
    Write-KV -Label 'Hyper-V Feature' -Value '(Skipped in Fast mode)' -State Neutral
} else {
    $hyperVStateValue = $null
    $hypervisorPresent = $null
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $hypervisorPresent = [bool]$cs.HypervisorPresent
    } catch { }
    try {
        if (-not $Quiet) { Write-Info 'Checking Hyper-V feature (this can take a few seconds)...' }
        # First try CIM which works without elevation and is fast
        $opt = Get-CimInstance -ClassName Win32_OptionalFeature -Filter "Name='Microsoft-Hyper-V-All'" -ErrorAction Stop
        if ($null -ne $opt) {
            switch ($opt.InstallState) {
                1 { $hyperVStateValue = 'Enabled' }
                2 { $hyperVStateValue = 'Disabled' }
                default { $hyperVStateValue = 'Unknown' }
            }
        }
    } catch {
        # Ignore and try DISM fallback below
    }
    if ($null -eq $hyperVStateValue) {
        try {
            if (-not $Quiet) { Write-Info 'Fallback to DISM online feature query for Hyper-V (can take 10–60 seconds)...' }
            $featureNames = @('Microsoft-Hyper-V','Microsoft-Hyper-V-All','Microsoft-Hyper-V-Tools-All','Microsoft-Hyper-V-Hypervisor')
            $found = $false
            foreach ($fname in $featureNames) {
                try {
                    $f = Get-WindowsOptionalFeature -Online -FeatureName $fname -ErrorAction Stop
                    if ($null -ne $f) {
                        $found = $true
                        $hyperVStateValue = $f.State
                        break
                    }
                } catch { }
            }
            if (-not $found -and $null -eq $hyperVStateValue) { $hyperVStateValue = 'Not Applicable' }
        } catch {
            $hyperVStateValue = 'Unknown'
            if (-not $IsElevated) { $hyperVStateValue += ' (try running as administrator)' }
        }
    }
    if ($hypervisorPresent -eq $true -and ($hyperVStateValue -eq 'Disabled' -or $hyperVStateValue -eq 'Not Applicable' -or $null -eq $hyperVStateValue)) {
        $hyperVStateValue = 'Enabled (Hypervisor Present)'
    }
    $hvStateClass = if ($hyperVStateValue -like 'Enabled*') { 'Good' } elseif ($hyperVStateValue -eq 'Disabled') { 'Warn' } elseif ($hyperVStateValue -eq 'Not Applicable') { 'Neutral' } else { 'Neutral' }
    Write-KV -Label 'Hyper-V Feature' -Value $hyperVStateValue -State $hvStateClass
}

# Report Hypervisor presence explicitly (fast, safe)
try {
    if ($null -eq $hypervisorPresent) {
        $cs2 = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $hypervisorPresent = [bool]$cs2.HypervisorPresent
    }
    $hvPresentValue = if ($hypervisorPresent) { 'Yes' } else { 'No' }
    $hvPresentState = if ($hypervisorPresent) { 'Good' } else { 'Neutral' }
    Write-KV -Label 'Hypervisor Present' -Value $hvPresentValue -State $hvPresentState
} catch { Write-KV -Label 'Hypervisor Present' -Value 'Unknown' -State Neutral }
# Windows Update Status
if ($FastMode) {
    if (-not $Quiet) { Write-Info 'Skipping Windows Update history due to -Fast/-SkipSlowChecks' }
    Write-KV -Label 'Last Update Installed' -Value '(Skipped in Fast mode)' -State Neutral
    Write-KV -Label 'Last Update KB' -Value '(Skipped in Fast mode)' -State Neutral
} else {
    try {
        if (-not $Quiet) { Write-Info 'Querying Windows Update history (can take 10–30 seconds)...' }
        $wu = Get-HotFix | Sort-Object -Property InstalledOn -Descending | Select-Object -First 1
        Write-KV -Label 'Last Update Installed' -Value $wu.InstalledOn -State Neutral
        Write-KV -Label 'Last Update KB' -Value $wu.HotFixID -State Neutral
    } catch { Write-KV -Label 'Windows Update Status' -Value 'Unknown' -State Warn }
}

# UAC Level
try {
    $uac = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ErrorAction Stop
    $uacLevel = switch ($uac.ConsentPromptBehaviorAdmin) {
        0 { 'Never notify' }
        1 { 'Prompt for credentials on secure desktop' }
        2 { 'Prompt for consent on secure desktop' }
        3 { 'Prompt for credentials' }
        4 { 'Prompt for consent' }
        5 { 'Prompt for consent for non-Windows binaries' }
        default { 'Unknown' }
    }
    $uacState = if ($uac.ConsentPromptBehaviorAdmin -ge 2) { 'Good' } else { 'Warn' }
    Write-KV -Label 'UAC Level' -Value $uacLevel -State $uacState
} catch { Write-KV -Label 'UAC Level' -Value 'Unknown' -State Warn }

# Local Admin Accounts
try {
    $localAdmins = Get-LocalGroupMember -Group 'Administrators' -ErrorAction Stop | Where-Object { $_.ObjectClass -eq 'User' -and $_.PrincipalSource -eq 'Local' }
    $adminCount = ($localAdmins | Measure-Object).Count
    Write-KV -Label 'Local Admin Accounts' -Value $adminCount -State (if($adminCount -gt 2){'Warn'}else{'Good'})
} catch { Write-KV -Label 'Local Admin Accounts' -Value 'Unknown' -State Warn }

# RDP Status
try {
    $rdp = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -ErrorAction Stop
    $rdpEnabled = if ($rdp.fDenyTSConnections -eq 0) { 'Enabled' } else { 'Disabled' }
    $rdpState = if ($rdp.fDenyTSConnections -eq 0) { 'Warn' } else { 'Good' }
    Write-KV -Label 'Remote Desktop' -Value $rdpEnabled -State $rdpState
} catch { Write-KV -Label 'Remote Desktop' -Value 'Unknown' -State Warn }

# Network Shares
if ($FastMode) {
    if (-not $Quiet) { Write-Info 'Skipping SMB share enumeration due to -Fast/-SkipSlowChecks' }
    Write-Section 'Network Shares (non-admin)' -Value '(Skipped in Fast mode)' -State Neutral
} else {
    try {
        if (-not $Quiet) { Write-Info 'Enumerating SMB shares (may take a few seconds)...' }
        $shares = Get-SmbShare | Where-Object { $_.Name -ne 'IPC$' -and $_.Name -ne 'ADMIN$' -and $_.Name -notlike 'C$' }
        $shareCount = ($shares | Measure-Object).Count
        Write-Section 'Network Shares (non-admin)' -Value $shareCount -State (if($shareCount -gt 0){'Warn'}else{'Good'})
    } catch { Write-KV -Label 'Network Shares' -Value 'Unknown' -State Warn }
}

# Advanced Defender Features
# Enhanced EDR in Block Mode Detection
try {
    $edrBlock = $null
    
    # Method 1: Check MpComputerStatus
    if ($null -ne $localdefender -and ($localdefender.PSObject.Properties.Name -contains 'EDRInBlockMode')) {
        $edrBlock = [bool]$localdefender.EDRInBlockMode
    }
    
    # Method 2: Check registry (requires specific permissions)
    if ($null -eq $edrBlock) {
        $mdeReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status' -ErrorAction SilentlyContinue
        if ($mdeReg -and $mdeReg.PSObject.Properties.Name -contains 'EDRInBlockMode') {
            $edrBlock = [bool]$mdeReg.EDRInBlockMode
        }
    }
    
    # Method 3: Check MDE service status
    if ($null -eq $edrBlock) {
        $mdeService = Get-Service -Name 'Sense' -ErrorAction SilentlyContinue
        $onboardingState = if ($mdeReg) { $mdeReg.OnboardingState } else { $null }
        
        if ($mdeService -and $mdeService.Status -eq 'Running' -and $onboardingState -eq 1) {
            # If MDE is running and onboarded, try PowerShell method
            try {
                $edrStatus = Get-MpPreference | Select-Object -ExpandProperty 'CloudBlockLevel' -ErrorAction SilentlyContinue
                $edrBlock = $edrStatus -gt 0
            } catch { }
        }
    }
    
    # Display result
    if ($null -eq $edrBlock) {
        $mdeOnboard = if ($mdeReg -and $mdeReg.OnboardingState -eq 1) { 'Yes' } else { 'No' }
        Write-KV -Label 'EDR in Block Mode' -Value "Unknown (MDE Onboarded: $mdeOnboard)" -State Neutral
    } else {
        $edrText = if ($edrBlock) { 'Enabled' } else { 'Disabled' }
        $edrState = if ($edrBlock) { 'Good' } else { 'Warn' }
        Write-KV -Label 'EDR in Block Mode' -Value $edrText -State $edrState
    }
} catch {
    Write-KV -Label 'EDR in Block Mode' -Value "Error: $($_.Exception.Message)" -State Warn
}

# Enhanced Tamper Protection Detection
try {
    $tp = $null
    
    # Method 1: Get-MpComputerStatus (most reliable)
    if ($null -ne $localdefender -and ($localdefender.PSObject.Properties.Name -contains 'IsTamperProtected')) {
        $tp = [bool]$localdefender.IsTamperProtected
    }
    
    # Method 2: Registry check (fallback)
    if ($null -eq $tp) {
        $tpReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Features' -Name 'TamperProtection' -ErrorAction SilentlyContinue
        if ($tpReg -and $null -ne $tpReg.TamperProtection) {
            $tp = ([int]$tpReg.TamperProtection) -eq 1
        }
    }
    
    # Method 3: Alternative registry location
    if ($null -eq $tp) {
        $tpReg2 = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Real-Time Protection' -Name 'TamperProtection' -ErrorAction SilentlyContinue
        if ($tpReg2 -and $null -ne $tpReg2.TamperProtection) {
            $tp = ([int]$tpReg2.TamperProtection) -eq 1
        }
    }
    
    # Method 4: Check if cloud-managed
    if ($null -eq $tp) {
        $cloudManaged = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Features' -Name 'TamperProtectionSource' -ErrorAction SilentlyContinue
        if ($cloudManaged -and $cloudManaged.TamperProtectionSource -eq 1) {
            Write-KV -Label 'Tamper Protection' -Value 'Cloud Managed (Status Unknown)' -State Neutral
            $tp = 'CloudManaged'
        }
    }
    
    # Display result
    if ($null -eq $tp) {
        Write-KV -Label 'Tamper Protection' -Value 'Unknown (Check Defender Security Center)' -State Warn
    } elseif ($tp -eq 'CloudManaged') {
        # Already displayed above
    } else {
        $tpText = if ($tp) { 'Enabled' } else { 'Disabled' }
        $tpState = if ($tp) { 'Good' } else { 'Warn' }
        Write-KV -Label 'Tamper Protection' -Value $tpText -State $tpState
    }
} catch {
    Write-KV -Label 'Tamper Protection' -Value "Error: $($_.Exception.Message)" -State Warn
}

# Enhanced Device Health Attestation Detection
try {
    $dhaStatus = 'Not Available'
    $dhaDetails = @()
    
    # Method 1: Check service configuration
    $dhaService = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\DeviceHealthAttestation' -ErrorAction SilentlyContinue
    if ($dhaService) {
        $startType = switch ($dhaService.Start) {
            0 { 'Boot' }
            1 { 'System' }
            2 { 'Automatic' }
            3 { 'Manual' }
            4 { 'Disabled' }
            default { "Unknown($($dhaService.Start))" }
        }
        $dhaDetails += "Service: $startType"
        
        if ($dhaService.Start -le 2) {
            $dhaStatus = 'Available'
        }
    }
    
    # Method 2: Check actual service status
    try {
        $dhaSvc = Get-Service -Name 'DeviceHealthAttestationService' -ErrorAction SilentlyContinue
        if ($dhaSvc) {
            $dhaDetails += "Status: $($dhaSvc.Status)"
            if ($dhaSvc.Status -eq 'Running' -or $dhaSvc.StartType -eq 'Automatic') {
                $dhaStatus = 'Available'
            }
        }
    } catch { }
    
    # Method 3: Check registry configuration
    $dhaReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceHealthAttestation' -ErrorAction SilentlyContinue
    if ($dhaReg) {
        $dhaDetails += 'Registry: Configured'
        if ($dhaReg.PSObject.Properties.Name -contains 'Enabled' -and $dhaReg.Enabled -eq 1) {
            $dhaStatus = 'Available'
        }
    }
    
    # Method 4: Check enterprise policy
    $dhaPol = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\DeviceHealthAttestation' -ErrorAction SilentlyContinue
    if ($dhaPol) {
        $dhaDetails += 'Policy: Configured'
        if ($dhaPol.PSObject.Properties.Name -contains 'EnableDeviceHealthAttestation' -and $dhaPol.EnableDeviceHealthAttestation -eq 1) {
            $dhaStatus = 'Available'
        }
    }
    
    # Method 5: Check TPM requirement (DHA requires TPM)
    if ($tpmPresent -eq $false -and $dhaStatus -eq 'Available') {
        $dhaStatus = 'Not Available (No TPM)'
    }
    
    # Display result with better formatting
    $dhaState = switch ($dhaStatus) {
        'Available' { 'Good' }
        'Not Available (No TPM)' { 'Warn' }
        default { 'Neutral' }
    }
    
    # Main status line
    Write-KV -Label 'Device Health Attestation' -Value $dhaStatus -State $dhaState
    
    # Display detailed information as sub-items if available
    if ($dhaDetails.Count -gt 0) {
        foreach ($detail in $dhaDetails) {
            $detailParts = $detail -split ': ', 2
            if ($detailParts.Count -eq 2) {
                $detailLabel = "  └─ $($detailParts[0])"
                $detailValue = $detailParts[1]
                Write-KV -Label $detailLabel -Value $detailValue -State Neutral
            }
        }
    }
} catch { 
    Write-KV -Label 'Device Health Attestation' -Value "Error: $($_.Exception.Message)" -State Warn 
}

# Enhanced Defender for Endpoint Detection
try {
    $mdeStatus = 'Not Onboarded'
    $mdeDetails = @()
    
    # Method 1: Check primary onboarding registry
    $mdeReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status' -ErrorAction SilentlyContinue
    if ($mdeReg) {
        if ($mdeReg.PSObject.Properties.Name -contains 'OnboardingState') {
            $onboardingState = $mdeReg.OnboardingState
            $mdeDetails += "OnboardingState: $onboardingState"
            
            if ($onboardingState -eq 1) {
                $mdeStatus = 'Onboarded'
                
                # Add last connected info if available
                if ($mdeReg.PSObject.Properties.Name -contains 'LastConnected') {
                    $lastConnected = $mdeReg.LastConnected
                    if ($lastConnected) {
                        try {
                            $lastConnectedDate = [DateTime]::FromFileTime($lastConnected).ToString('yyyy-MM-dd HH:mm')
                            $mdeDetails += "Last Connected: $lastConnectedDate"
                        } catch { 
                            $mdeDetails += "Last Connected: $lastConnected"
                        }
                    }
                }
                
                # Check org ID
                if ($mdeReg.PSObject.Properties.Name -contains 'OrgId') {
                    $orgId = $mdeReg.OrgId
                    if ($orgId) {
                        $mdeDetails += "OrgId: $($orgId.Substring(0, [Math]::Min(8, $orgId.Length)))..."
                    }
                }
            }
        }
        
        # Check sense version
        if ($mdeReg.PSObject.Properties.Name -contains 'SenseVersion') {
            $senseVer = $mdeReg.SenseVersion
            if ($senseVer) {
                $mdeDetails += "Version: $senseVer"
            }
        }
    }
    
    # Method 2: Check Sense service status
    try {
        $senseService = Get-Service -Name 'Sense' -ErrorAction SilentlyContinue
        if ($senseService) {
            $mdeDetails += "Sense Service: $($senseService.Status)"
            if ($senseService.Status -eq 'Running' -and $mdeStatus -eq 'Not Onboarded') {
                $mdeStatus = 'Service Running (Check Onboarding)'
            }
        } else {
            $mdeDetails += 'Sense Service: Not Installed'
        }
    } catch { }
    
    # Method 3: Check WdFilter driver
    try {
        $wdFilter = Get-Service -Name 'WdFilter' -ErrorAction SilentlyContinue
        if ($wdFilter -and $wdFilter.Status -eq 'Running') {
            $mdeDetails += 'WdFilter: Running'
        }
    } catch { }
    
    # Method 4: Check MDE platform version
    try {
        $mdeVer = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection' -ErrorAction SilentlyContinue
        if ($mdeVer -and $mdeVer.PSObject.Properties.Name -contains 'InstallTime') {
            $installTime = $mdeVer.InstallTime
            if ($installTime) {
                try {
                    $installDate = [DateTime]::FromFileTime($installTime).ToString('yyyy-MM-dd')
                    $mdeDetails += "Installed: $installDate"
                } catch { }
            }
        }
    } catch { }
    
    # Method 5: Check cloud connectivity
    if ($mdeStatus -eq 'Onboarded') {
        try {
            $connectivity = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status' -Name 'OnboardingInfo' -ErrorAction SilentlyContinue
            if ($connectivity -and $connectivity.OnboardingInfo) {
                $mdeDetails += 'Cloud: Connected'
            }
        } catch { }
    }
    
    # Display result with better formatting
    $mdeState = switch ($mdeStatus) {
        'Onboarded' { 'Good' }
        'Service Running (Check Onboarding)' { 'Warn' }
        'Not Onboarded' { 'Neutral' }
        default { 'Neutral' }
    }
    
    # Main status line
    Write-KV -Label 'Defender for Endpoint' -Value $mdeStatus -State $mdeState
    
    # Display detailed information as sub-items if onboarded
    if ($mdeStatus -eq 'Onboarded' -and $mdeDetails.Count -gt 0) {
        foreach ($detail in $mdeDetails) {
            $detailParts = $detail -split ': ', 2
            if ($detailParts.Count -eq 2) {
                $detailLabel = "  └─ $($detailParts[0])"
                $detailValue = $detailParts[1]
                Write-KV -Label $detailLabel -Value $detailValue -State Neutral
            }
        }
    } elseif ($mdeDetails.Count -gt 0) {
        # For non-onboarded status, show key details inline
        $keyDetails = $mdeDetails | Where-Object { $_ -match '^(Sense Service|OnboardingState):' } | Select-Object -First 2
        if ($keyDetails) {
            Write-KV -Label '  └─ Details' -Value ($keyDetails -join ', ') -State Neutral
        }
    }
} catch { 
    Write-KV -Label 'Defender for Endpoint' -Value "Error: $($_.Exception.Message)" -State Warn 
}

# Enhanced Credential Guard Detection
try {
    $cg = $null
    $cgRunning = $null
    
    # Method 1: Check DeviceGuard registry
    $dgReg = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' -ErrorAction SilentlyContinue
    if ($dgReg -and $null -ne $dgReg.EnableVirtualizationBasedSecurity) {
        $vbsEnabled = [int]$dgReg.EnableVirtualizationBasedSecurity -eq 1
    } else { $vbsEnabled = $false }
    
    # Method 2: Check LSA configuration
    $lsaReg = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' -ErrorAction SilentlyContinue
    if ($lsaReg -and $null -ne $lsaReg.LsaCfgFlags) {
        $cgConfigured = [int]$lsaReg.LsaCfgFlags -gt 0
    } else { $cgConfigured = $false }
    
    # Method 3: Check if Credential Guard is running
    try {
        $cgService = Get-CimInstance -ClassName Win32_DeviceGuard -ErrorAction SilentlyContinue
        if ($cgService -and $cgService.SecurityServicesRunning -contains 1) {
            $cgRunning = $true
        }
    } catch { }
    
    # Determine status
    if ($cgRunning -eq $true) {
        Write-KV -Label 'Credential Guard' -Value 'Running' -State Good
    } elseif ($vbsEnabled -and $cgConfigured) {
        Write-KV -Label 'Credential Guard' -Value 'Configured (VBS Required)' -State Good
    } elseif ($cgConfigured) {
        Write-KV -Label 'Credential Guard' -Value 'Configured (VBS Disabled)' -State Warn
    } elseif ($vbsEnabled) {
        Write-KV -Label 'Credential Guard' -Value 'VBS Enabled (CG Not Configured)' -State Warn
    } else {
        Write-KV -Label 'Credential Guard' -Value 'Not Configured' -State Neutral
    }
} catch {
    Write-KV -Label 'Credential Guard' -Value "Error: $($_.Exception.Message)" -State Warn
}

Write-Section 'Report Summary'
# Calculate compliance score based on security settings
$goodSettings = 0
$totalSettings = 0

# Count security settings with good/bad states
foreach ($key in @('RealTimeProtectionEnabled', 'BehaviorMonitorEnabled', 'IoavProtectionEnabled', 'NetworkProtectionEnabled', 'PUAProtectionEnabled')) {
    if ($results.$key -ne $null) {
        $totalSettings++
        if ($results.$key -eq 'True' -or $results.$key -eq 'Enabled') { $goodSettings++ }
    }
}

# Add ASR rules count
if ($AsrStatus -and $AsrStatus.Count -gt 0) {
    $asrEnabled = ($AsrStatus | Where-Object { $_.State -eq 'Enabled' }).Count
    $goodSettings += $asrEnabled
    $totalSettings += $AsrStatus.Count
}

# Calculate percentage
$compliancePercent = if ($totalSettings -gt 0) { [math]::Round(($goodSettings / $totalSettings) * 100, 1) } else { 0 }
$complianceState = if ($compliancePercent -ge 80) { 'Good' } elseif ($compliancePercent -ge 60) { 'Warn' } else { 'Bad' }

Write-KV -Label 'Security Compliance' -Value "$compliancePercent% ($goodSettings/$totalSettings)" -State $complianceState
Write-KV -Label 'Report Generated' -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -State Neutral

Write-Section 'Script Information'

Write-Info "`nScript Finished"

## Build structured output
$ResultObject = [PSCustomObject]@{
    ScriptVersion = $ScriptVersion
    Timestamp = (Get-Date)
    ASR = $AsrStatus
    Preferences = if($IncludeRaw){ $results } else { $null }
    ComputerStatus = if($IncludeRaw){ $localdefender } else { $null }
    Versions = [PSCustomObject]@{
        LocalSignature = $localdefender.AntispywareSignatureVersion
        LatestSignature = $LatestSig
        LocalEngine = $localdefender.AMEngineVersion
        LatestEngine = $LatestEngine
        LocalPlatform = $localdefender.AMServiceVersion
        LatestPlatform = $LatestPlatform
        LatestReleaseDate = $LatestReleaseDate
    }
    DeviceGuard = if ($dg) { $dg } else { $null }
    CredentialGuard = if ($cg) { $cg } else { $null }
}

switch ($OutputMode) {
    'Object' { return $ResultObject }
    'Json' {
        if (-not $OutputPath) { $OutputPath = Join-Path -Path (Get-Location) -ChildPath 'defender-status.json' }
        $ResultObject | ConvertTo-Json -Depth 6 | Set-Content -Path $OutputPath -Encoding UTF8
        Write-Info "JSON written to $OutputPath"
        return $ResultObject
    }
    'Csv' {
        if (-not $OutputPath) { $OutputPath = Join-Path -Path (Get-Location) -ChildPath "defender-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv" }
        $flatResults = @()
        foreach ($prop in $ResultObject.PSObject.Properties) {
            if ($prop.Value -is [hashtable] -or $prop.Value -is [PSCustomObject]) {
                if ($prop.Value -is [hashtable]) {
                    foreach ($item in $prop.Value.GetEnumerator()) {
                        $flatResults += [PSCustomObject]@{
                            Section = $prop.Name
                            Setting = $item.Key
                            Value = $item.Value
                        }
                    }
                } else {
                    foreach ($subProp in $prop.Value.PSObject.Properties) {
                        $flatResults += [PSCustomObject]@{
                            Section = $prop.Name
                            Setting = $subProp.Name
                            Value = $subProp.Value
                        }
                    }
                }
            } else {
                $flatResults += [PSCustomObject]@{
                    Section = 'General'
                    Setting = $prop.Name
                    Value = $prop.Value
                }
            }
        }
        $flatResults | Export-Csv -Path $OutputPath -NoTypeInformation
        Write-Info "CSV written to $OutputPath"
        return
    }
    'Html' {
        if (-not $OutputPath) { $OutputPath = Join-Path -Path (Get-Location) -ChildPath "defender-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html" }
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows Defender Security Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0078d4; }
        h2 { color: #323130; border-bottom: 1px solid #ddd; }
        table { border-collapse: collapse; width: 100%; margin: 10px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f8f9fa; }
        .good { color: #107c10; }
        .warn { color: #d83b01; }
        .neutral { color: #605e5c; }
    </style>
</head>
<body>
    <h1>Windows Defender Security Report</h1>
    <p>Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
"@
        foreach ($prop in $ResultObject.PSObject.Properties) {
            if ($prop.Name -ne 'Timestamp') {
                $html += "<h2>$($prop.Name)</h2><table><tr><th>Setting</th><th>Value</th></tr>"
                if ($prop.Value -is [hashtable] -or $prop.Value -is [PSCustomObject]) {
                    if ($prop.Value -is [hashtable]) {
                        foreach ($item in $prop.Value.GetEnumerator()) {
                            $html += "<tr><td>$($item.Key)</td><td>$($item.Value)</td></tr>"
                        }
                    } else {
                        foreach ($subProp in $prop.Value.PSObject.Properties) {
                            $html += "<tr><td>$($subProp.Name)</td><td>$($subProp.Value)</td></tr>"
                        }
                    }
                } else {
                    $html += "<tr><td>$($prop.Name)</td><td>$($prop.Value)</td></tr>"
                }
                $html += "</table>"
            }
        }
        $html += "</body></html>"
        $html | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Info "HTML written to $OutputPath"
        return
    }
    'Xml' {
        if (-not $OutputPath) { $OutputPath = Join-Path -Path (Get-Location) -ChildPath "defender-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').xml" }
        $ResultObject | Export-Clixml -Path $OutputPath
        Write-Info "XML written to $OutputPath"
        return
    }
    'Plain' { # Already wrote plain host output
        return
    }
    default { return }
}
