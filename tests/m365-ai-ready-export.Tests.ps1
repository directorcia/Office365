#requires -Version 7.0
<#
.SYNOPSIS
    Offline Pester and standalone checks for m365-ai-ready-export.ps1.
.EXAMPLE
    Invoke-Pester -Path ./tests/m365-ai-ready-export.Tests.ps1
.EXAMPLE
    pwsh -NoProfile -File ./tests/m365-ai-ready-export.Tests.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if (-not (Get-Command Describe -ErrorAction SilentlyContinue)) {
    function Describe { param([string]$Name, [scriptblock]$Fixture) & $Fixture }
    function Context { param([string]$Name, [scriptblock]$Fixture) & $Fixture }
    function It { param([string]$Name, [scriptblock]$Test) & $Test }
}

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

Describe 'm365-ai-ready-export' {
    Context 'SharingCapability normalisation and exposure scoring' {
        It 'counts anyone-link and external-sharing sites across ints, strings, and enums' {
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
        }

        It 'distinguishes empty SharePoint site inventories from failed collections' {
            $script:MockSharingCapability = 2
            $script:MockTenantCollectionError = $null
            $script:MockTenantCollectionNull = $false
            $script:MockSiteCollectionError = $null
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
                if ($script:MockTenantCollectionError) { throw $script:MockTenantCollectionError }
                if ($script:MockTenantCollectionNull) { return }
                [pscustomobject]@{
                    SharingCapability = $script:MockSharingCapability
                    DefaultSharingLinkType = 'Internal'
                    RestrictedSharePointSearch = $false
                }
            }

            function Get-PnPTenantSite {
                [CmdletBinding()]
                param([switch]$IncludeOneDriveSites)
                if ($script:MockSiteCollectionError) { throw $script:MockSiteCollectionError }
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

            $script:MockSharingCapability = 2
            $script:MockSites = @()
            $emptyInventory = Get-SharePointOneDriveReadinessLocal
            Assert-Equal $emptyInventory.Status 'reported' 'A successfully collected empty site inventory remains reported.'
            Assert-Equal $emptyInventory.SiteInventoryStatus 'reported' 'An empty site inventory is distinct from a failed collection.'
            Assert-Equal $emptyInventory.TotalSiteCount 0 'A successful empty site inventory reports zero sites.'

            $script:MockSiteCollectionError = 'Access denied while listing sites.'
            $failedInventory = Get-SharePointOneDriveReadinessLocal
            Assert-Equal $failedInventory.Status 'partial' 'A site inventory failure keeps successfully collected tenant settings.'
            Assert-Equal $failedInventory.SiteInventoryStatus 'not collected' 'A site inventory failure is explicit.'
            Assert-Equal $failedInventory.TotalSiteCount 'not collected' 'A site inventory failure is not reported as zero sites.'
            Assert-Equal $failedInventory.OneDrivePersonalSiteCount 'not collected' 'A site inventory failure does not fabricate a OneDrive count.'
            Assert-True ($failedInventory.SiteInventoryReason -match 'Access denied') 'The site inventory error is preserved for diagnosis.'
            $script:Report = [ordered]@{ SharePointOneDrive = $failedInventory }
            $failedExposure = Get-SharePointExposureAdvanced
            Assert-Equal $failedExposure.AnyoneLinkSites 'not collected' 'Failed site collection does not become zero anonymous-link sites.'

            $script:MockSiteCollectionError = $null
            $script:MockTenantCollectionError = 'Tenant settings unavailable.'
            $script:MockSites = @([pscustomobject]@{
                Url = 'https://contoso.sharepoint.com/sites/available'
                StorageQuota = 1
                StorageUsageCurrent = 0
                SharingCapability = 1
                LastContentModifiedDate = [datetime]::UtcNow
            })
            $tenantFailure = Get-SharePointOneDriveReadinessLocal
            Assert-Equal $tenantFailure.Status 'partial' 'Site data remains available when tenant settings fail.'
            Assert-Equal $tenantFailure.ExternalSharingLevel 'not collected' 'Failed tenant settings remain unknown.'
            Assert-Equal $tenantFailure.TotalSiteCount 1 'A successful site inventory is retained when tenant settings fail.'
            $script:MockTenantCollectionError = $null
            $script:MockTenantCollectionNull = $true
            $nullTenant = Get-SharePointOneDriveReadinessLocal
            Assert-Equal $nullTenant.Status 'partial' 'A null tenant result is not reported as complete collection.'
            Assert-True ($nullTenant.Reason -match 'returned no data') 'A null tenant result explains which collection returned no data.'
            $script:MockTenantCollectionNull = $false
        }

        It 'propagates AnyoneLinkSites into readiness scoring and recommendations' {
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
        }
    }

    Context 'Exchange, user caching, licensing, and parameter-set behaviour' {
        It 'records Exchange mailbox query errors instead of fabricating zero mailboxes' {
            function Get-ConnectionInformation {
                [CmdletBinding()]
                param()
                [pscustomobject]@{ Name = 'ExchangeOnline' }
            }

            function Get-EXOMailbox {
                [CmdletBinding()]
                param($ResultSize, $Properties)
                throw 'Mailbox query access denied.'
            }

            function Test-GraphPermissionsAvailable {
                param([string[]]$RequiredPermissions)
                return $false
            }

            $exchangeFailure = Get-ExchangeReadiness
            Assert-Equal $exchangeFailure.Status 'not collected' 'A failed Exchange query with no Graph fallback is not reported as an empty mailbox list.'
            Assert-Equal $exchangeFailure.TotalMailboxes 'not collected' 'A failed Exchange query does not fabricate zero mailboxes.'
            Assert-True ($exchangeFailure.Reason -match 'Mailbox query access denied') 'The Exchange query failure is included in the report reason.'
        }

        It 'merges sign-in activity by user ID and evaluates Copilot prerequisite plans accurately' {
            $script:GraphUsersCache = @{}
            $mockBaseUsers = @(
                [pscustomobject]@{
                    Id = 'user-ready'
                    userType = 'Member'
                    assignedPlans = @(
                        [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'MicrosoftOffice' }
                        [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'exchange' }
                        [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'SharePointEnterprise' }
                        [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'TeamspaceAPI' }
                    )
                }
                [pscustomobject]@{
                    Id = 'user-partial'
                    userType = 'Guest'
                    assignedPlans = @(
                        [pscustomobject]@{ capabilityStatus = 'Deleted'; service = 'MicrosoftOffice' }
                        [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'exchange' }
                    )
                }
            )
            $mockSignInUsers = @(
                [pscustomobject]@{
                    Id = 'user-ready'
                    SignInActivity = [pscustomobject]@{ LastSignInDateTime = (Get-Date).ToUniversalTime().AddDays(-120).ToString('o') }
                }
            )

            function Get-GraphRestAll {
                param([string]$Uri, [hashtable]$Query = @{}, [switch]$BypassCache)
                if ($Uri -match '/subscribedSkus$') {
                    return @([pscustomobject]@{ SkuPartNumber = 'SPE_E5'; EnabledUnits = 10; ConsumedUnits = 2 })
                }
                if ($Uri -match '/users$') {
                    if ($Query['$select'] -match 'signInActivity') { return $mockSignInUsers }
                    return $mockBaseUsers
                }
                return @()
            }
            function Test-GraphPermissionsAvailable { param([string[]]$RequiredPermissions) return $true }

            $merged = @(Get-CachedGraphUsers -IncludeSignInActivity)
            Assert-Equal $merged.Count 2 'Both cached users are returned.'
            Assert-True ($null -ne $merged[0].SignInActivity) 'Matching user receives SignInActivity.'
            Assert-True ($null -eq $merged[1].SignInActivity) 'Unmatched user receives null SignInActivity without StrictMode error.'

            $licensing = Get-LicensingReadiness
            Assert-Equal $licensing.PrereqReadyUserCount 1 'Only the user with all four enabled workload services is counted ready.'
            Assert-Equal $licensing.PrereqReadyPercent 50 'PrereqReadyPercent is 50% for 1 of 2 users.'
        }

        It 'resolves output paths with spaces without creating directories prematurely and enforces Child parameter set' {
            $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("ai-ready-path test-" + [guid]::NewGuid().ToString('N'))
            $targetFile = Join-Path $tempDir 'Sub Folder [2026]\Contoso AI Readiness.json'
            $resolved = Resolve-OutputFilePath -Path $targetFile
            Assert-Equal $resolved ([System.IO.Path]::GetFullPath($targetFile)) 'Path with spaces and brackets resolves cleanly.'
            Assert-True (-not (Test-Path -LiteralPath $tempDir)) 'Resolve-OutputFilePath does not create directories before ShouldProcess.'

            $cmdInfo = Get-Command $scriptPath
            $paramSets = @($cmdInfo.ParameterSets | Select-Object -ExpandProperty Name)
            Assert-True ($paramSets -contains 'Standard' -and $paramSets -contains 'Child') 'Script defines Standard and Child parameter sets.'
            $defaultSet = $cmdInfo.ParameterSets | Where-Object { $_.IsDefault } | Select-Object -ExpandProperty Name -First 1
            Assert-Equal $defaultSet 'Standard' 'Standard is the default parameter set.'

            $childSet = $cmdInfo.ParameterSets | Where-Object { $_.Name -eq 'Child' }
            $adminUrlParam = $childSet.Parameters | Where-Object { $_.Name -eq 'SharePointAdminUrl' }
            Assert-True ([bool]$adminUrlParam.IsMandatory) 'SharePointAdminUrl is mandatory in Child parameter set.'
        }

        It 'does not auto-install PnP.PowerShell when only Windows PowerShell modules exist' {
            $installCommands = @($ast.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Install-Module'
            }, $true))
            Assert-Equal $installCommands.Count 0 'Script does not invoke Install-Module during read-only collection.'
        }
    }
}

Write-Output 'm365-ai-ready-export collector tests passed.'
