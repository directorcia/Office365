#requires -Version 7.0
<#
.SYNOPSIS
    Offline checks for SharePoint SharingCapability normalisation.
.EXAMPLE
    pwsh -NoProfile -File ./tests/m365-ai-ready-export.Tests.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'AI-Ready/m365-ai-ready-export.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }

foreach ($definition in $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
}, $false)) {
    . ([scriptblock]::Create($definition.Extent.Text))
}

# Same values as Microsoft.SharePoint.Client.SharingCapabilities. Defined locally so the
# test does not need the SharePoint CSOM assembly.
enum SharingCapabilities {
    Disabled = 0
    ExternalUserSharingOnly = 1
    ExternalUserAndGuestSharing = 2
    ExistingExternalUserSharingOnly = 3
}

function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    if ($Actual -ne $Expected) {
        throw "Assertion failed: $Message. Expected '$Expected' but got '$Actual'."
    }
}

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

function Get-SharingSiteSample {
    param([object]$Capability, [int]$Count)
    for ($i = 0; $i -lt $Count; $i++) {
        [pscustomobject]@{ SharingCapability = $Capability }
    }
}

function Get-SharePointExposureSample {
    param([object[]]$Sites, [switch]$AsJson)
    $section = [ordered]@{
        Status               = 'reported'
        ExternalSharingLevel = 'Disabled'
        Sites                = @($Sites)
    }
    if ($AsJson) {
        $section = ($section | ConvertTo-Json -Depth 6 | ConvertFrom-Json)
    }
    $script:Report = [ordered]@{ SharePointOneDrive = $section }
    return Get-SharePointExposureAdvanced
}

# Customer distribution, 18 Nov 2026: 17 disabled, 232 external-user, 115 anyone, 2 existing-external.
$customerSites = @(
    Get-SharingSiteSample -Capability 0 -Count 17
    Get-SharingSiteSample -Capability 1 -Count 232
    Get-SharingSiteSample -Capability 2 -Count 115
    Get-SharingSiteSample -Capability 3 -Count 2
)
Assert-Equal @($customerSites).Count 366 'Customer sample has 366 sites.'

$numeric = Get-SharePointExposureSample -Sites $customerSites -AsJson
Assert-Equal $numeric.AnyoneLinkSites 115 'JSON integer SharingCapability values count anyone-link sites (capability 2).'
Assert-Equal $numeric.ExternalSharingSites 349 'JSON integer SharingCapability values count external-sharing sites (1, 2, or 3).'
$jsonCapability = ([pscustomobject]@{ SharingCapability = 2 } | ConvertTo-Json | ConvertFrom-Json).SharingCapability
Assert-Equal (ConvertTo-SharingCapabilityName -Value $jsonCapability) 'ExternalUserAndGuestSharing' 'JSON-deserialised ints normalise to ExternalUserAndGuestSharing.'
Assert-Equal $jsonCapability.GetType().FullName 'System.Int64' 'ConvertFrom-Json preserves SharingCapability as Int64.'

$nameSites = @(
    Get-SharingSiteSample -Capability 'Disabled' -Count 17
    Get-SharingSiteSample -Capability 'ExternalUserSharingOnly' -Count 232
    Get-SharingSiteSample -Capability 'externaluserandguestsharing' -Count 115
    Get-SharingSiteSample -Capability 'ExistingExternalUserSharingOnly' -Count 2
)
$named = Get-SharePointExposureSample -Sites $nameSites
Assert-Equal $named.AnyoneLinkSites 115 'Enum names count anyone-link sites.'
Assert-Equal $named.ExternalSharingSites 349 'Enum names count external-sharing sites, including ExistingExternalUserSharingOnly.'

$enumSites = @(
    Get-SharingSiteSample -Capability ([SharingCapabilities]::Disabled) -Count 17
    Get-SharingSiteSample -Capability ([SharingCapabilities]::ExternalUserSharingOnly) -Count 232
    Get-SharingSiteSample -Capability ([SharingCapabilities]::ExternalUserAndGuestSharing) -Count 115
    Get-SharingSiteSample -Capability ([SharingCapabilities]::ExistingExternalUserSharingOnly) -Count 2
)
$enums = Get-SharePointExposureSample -Sites $enumSites
Assert-Equal $enums.AnyoneLinkSites 115 'PowerShell enum objects count anyone-link sites.'
Assert-Equal $enums.ExternalSharingSites 349 'PowerShell enum objects count external-sharing sites.'

$mixed = Get-SharePointExposureSample -Sites @(
    Get-SharingSiteSample -Capability '0' -Count 17
    Get-SharingSiteSample -Capability ([int64]1) -Count 232
    Get-SharingSiteSample -Capability '2' -Count 100
    Get-SharingSiteSample -Capability ([SharingCapabilities]::ExternalUserAndGuestSharing) -Count 15
    Get-SharingSiteSample -Capability 'ExistingExternalUserSharingOnly' -Count 2
    [pscustomobject]@{ SharingCapability = 'AnonymousAccess' }
)
Assert-Equal $mixed.AnyoneLinkSites 115 'Mixed numeric strings, ints, and enums count anyone-link sites.'
Assert-Equal $mixed.ExternalSharingSites 349 'AnonymousAccess is not counted as external sharing.'

Assert-Equal (ConvertTo-SharingCapabilityName -Value $null) $null 'Null capability stays null.'
Assert-Equal (ConvertTo-SharingCapabilityName -Value '') $null 'Blank capability stays null.'
Assert-Equal (ConvertTo-SharingCapabilityName -Value '  3  ') 'ExistingExternalUserSharingOnly' 'Numeric strings are trimmed.'
Assert-Equal (ConvertTo-SharingCapabilityName -Value 'AnonymousAccess') 'AnonymousAccess' 'AnonymousAccess is not rewritten to a real enum name.'
Assert-Equal (ConvertTo-SharingCapabilityName -Value 'not collected') 'not collected' 'Unrecognised text is preserved.'
Assert-Equal (ConvertTo-SharingCapabilityName -Value 4) '4' 'Unknown integers are preserved as text.'

# Per-site AnyoneLinkEnabled and tenant ExternalSharingLevel go through the same helper.
$script:MockSharingCapability = 2
$script:MockSites = @(
    [pscustomobject]@{ Url = 'https://contoso.sharepoint.com/sites/disabled'; StorageQuota = 1; StorageUsageCurrent = 0; SharingCapability = 0; LastContentModifiedDate = [datetime]::UtcNow }
    [pscustomobject]@{ Url = 'https://contoso.sharepoint.com/sites/external'; StorageQuota = 1; StorageUsageCurrent = 0; SharingCapability = '1'; LastContentModifiedDate = [datetime]::UtcNow }
    [pscustomobject]@{ Url = 'https://contoso.sharepoint.com/sites/anyone'; StorageQuota = 1; StorageUsageCurrent = 0; SharingCapability = ([SharingCapabilities]::ExternalUserAndGuestSharing); LastContentModifiedDate = [datetime]::UtcNow }
    [pscustomobject]@{ Url = 'https://contoso.sharepoint.com/sites/existing'; StorageQuota = 1; StorageUsageCurrent = 0; SharingCapability = 3; LastContentModifiedDate = [datetime]::UtcNow }
    [pscustomobject]@{ Url = 'https://contoso.sharepoint.com/sites/bogus'; StorageQuota = 1; StorageUsageCurrent = 0; SharingCapability = 'AnonymousAccess'; LastContentModifiedDate = [datetime]::UtcNow }
)

function Get-PnPTenant {
    [CmdletBinding()]
    param()
    [pscustomobject]@{
        SharingCapability = $script:MockSharingCapability
        DefaultSharingLinkType = 'Internal'
        RestrictedSharePointSearch = $false
    }
}

function Get-PnPTenantSite {
    [CmdletBinding()]
    param([switch]$IncludeOneDriveSites)
    # The collector always passes -IncludeOneDriveSites. Accept that switch and
    # return the prepared sample whether or not it is set.
    $sites = @($script:MockSites)
    if ($IncludeOneDriveSites) { return $sites }
    return $sites
}

$local = Get-SharePointOneDriveReadinessLocal
Assert-Equal $local.ExternalSharingLevel 'ExternalUserAndGuestSharing' 'Tenant integer 2 is emitted as ExternalUserAndGuestSharing.'
Assert-Equal ([string]$local.Sites[0].SharingCapability) '0' 'Per-site SharingCapability is still the collected value.'
Assert-Equal $local.Sites[0].AnyoneLinkEnabled $false 'Disabled does not enable anyone links.'
Assert-Equal $local.Sites[1].AnyoneLinkEnabled $false 'ExternalUserSharingOnly does not enable anyone links.'
Assert-Equal $local.Sites[2].AnyoneLinkEnabled $true 'ExternalUserAndGuestSharing enables anyone links.'
Assert-Equal $local.Sites[3].AnyoneLinkEnabled $false 'ExistingExternalUserSharingOnly does not enable anyone links.'
Assert-Equal $local.Sites[4].AnyoneLinkEnabled $false 'AnonymousAccess does not enable anyone links.'

$script:MockSharingCapability = [SharingCapabilities]::ExistingExternalUserSharingOnly
Assert-Equal (Get-SharePointOneDriveReadinessLocal).ExternalSharingLevel 'ExistingExternalUserSharingOnly' 'Tenant enum is emitted as its canonical name.'

$script:MockSharingCapability = 'externalusersharingonly'
Assert-Equal (Get-SharePointOneDriveReadinessLocal).ExternalSharingLevel 'ExternalUserSharingOnly' 'Tenant name is canonicalised.'

$script:MockSharingCapability = $null
Assert-Equal (Get-SharePointOneDriveReadinessLocal).ExternalSharingLevel 'not collected' 'A missing tenant capability stays not collected.'

# Scoring and recommendations consume AnyoneLinkSites; they do not re-read SharingCapability.
$script:Report = [ordered]@{
    ReadinessFlags = [ordered]@{
        MfaRegisteredPercent = 'not collected'
        CaPoliciesEnabled = 'not collected'
        SensitivityLabelsPublished = 'not collected'
        OneDrivePersonalSiteSharePercent = 'not collected'
        RestrictedSearchEnabled = 'not collected'
        ActiveUsersInReportPeriodPercent = 'not collected'
    }
    IdentityAccessAdvanced = [ordered]@{}
    DataProtectionAdvanced = [ordered]@{}
    SharePointExposureAdvanced = [ordered]@{ AnyoneLinkSites = 0 }
    AppGovernanceAdvanced = [ordered]@{}
    EndpointAppAdvanced = [ordered]@{}
    AdoptionAdvanced = [ordered]@{}
    ReadinessEvidence = [ordered]@{ OverallScore = 'not collected' }
}
$noLinks = Get-ReadinessEvidence
Assert-Equal $noLinks.DomainScores.ContentExposure 100 'Zero anyone-link sites receive the full content-exposure weight.'
$noLinkRecommendations = Get-Recommendations
Assert-True (-not (@($noLinkRecommendations.TopRisks) -join "`n" -match 'anonymous links')) 'Zero anyone-link sites do not raise the anyone-link recommendation.'

$script:Report.SharePointExposureAdvanced = [ordered]@{ AnyoneLinkSites = 115 }
$withLinks = Get-ReadinessEvidence
Assert-Equal $withLinks.DomainScores.ContentExposure 0 'Anyone-link sites receive the present weight of 0.'
Assert-Equal $withLinks.RiskSignals.AnyoneLinkSites 115 'Risk signals keep the anyone-link site count.'
$linkRecommendations = Get-Recommendations
Assert-True ((@($linkRecommendations.TopRisks) -join "`n") -match '115 sampled SharePoint/OneDrive sites allow anonymous links') 'Recommendations follow the corrected anyone-link count.'

Write-Output 'm365-ai-ready-export SharingCapability tests passed.'
