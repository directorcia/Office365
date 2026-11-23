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
            Assert-Equal $licensing.PrereqReadyDenominator 1 'Guests are excluded from the licence-readiness denominator.'
            Assert-Equal $licensing.PrereqReadyPercent 100 'PrereqReadyPercent uses enabled member users, so 1 of 1 members is 100%.'
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

    Context 'Permission implication, CSV reports, Copilot activity, Secure Score, gap walk, and labels' {
        It 'keeps broader Graph permissions, decoded CSV rows, active Copilot users, scored gaps, date walks, and null label counts' {
# Permission implication: Directory.Read.All and *.ReadWrite.All satisfy narrower reads.
Assert-True (Test-GraphPermissionSatisfied -GrantedPermission 'Directory.Read.All' -RequiredPermission 'Organization.Read.All') 'Directory.Read.All covers Organization.Read.All.'
Assert-True (Test-GraphPermissionSatisfied -GrantedPermission 'Directory.Read.All' -RequiredPermission 'Group.Read.All') 'Directory.Read.All covers Group.Read.All.'
Assert-True (Test-GraphPermissionSatisfied -GrantedPermission 'Directory.Read.All' -RequiredPermission 'User.Read.All') 'Directory.Read.All covers User.Read.All.'
Assert-True (Test-GraphPermissionSatisfied -GrantedPermission 'Directory.Read.All' -RequiredPermission 'GroupMember.Read.All') 'Directory.Read.All covers GroupMember.Read.All.'
Assert-True (Test-GraphPermissionSatisfied -GrantedPermission 'Directory.Read.All' -RequiredPermission 'Application.Read.All') 'Directory.Read.All covers Application.Read.All.'
Assert-True (Test-GraphPermissionSatisfied -GrantedPermission 'Directory.ReadWrite.All' -RequiredPermission 'Directory.Read.All') 'Directory.ReadWrite.All covers Directory.Read.All.'
Assert-True (Test-GraphPermissionSatisfied -GrantedPermission 'Directory.ReadWrite.All' -RequiredPermission 'Organization.Read.All') 'Directory.ReadWrite.All covers Organization.Read.All.'
Assert-True (Test-GraphPermissionSatisfied -GrantedPermission 'Group.ReadWrite.All' -RequiredPermission 'Group.Read.All') 'Group.ReadWrite.All covers Group.Read.All.'
Assert-True (Test-GraphPermissionSatisfied -GrantedPermission 'https://graph.microsoft.com/Reports.ReadWrite.All' -RequiredPermission 'Reports.Read.All') 'A full Graph URL ReadWrite permission covers Read.'
Assert-True (-not (Test-GraphPermissionSatisfied -GrantedPermission 'Group.ReadWrite.All' -RequiredPermission 'GroupMember.Read.All')) 'Group.ReadWrite.All does not cover GroupMember.Read.All.'
Assert-True (-not (Test-GraphPermissionSatisfied -GrantedPermission 'Reports.Read.All' -RequiredPermission 'Group.Read.All')) 'Reports.Read.All does not cover Group.Read.All.'

function Get-MgContext { return $null }
function New-UnsignedTestJwt {
    param([hashtable]$ClaimSet)
    $json = $ClaimSet | ConvertTo-Json -Compress -Depth 6
    $encodedClaimSet = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($json)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    return ('e30.{0}.sig' -f $encodedClaimSet)
}
$script:GraphAccessToken = New-UnsignedTestJwt @{ roles = @('Directory.Read.All', 'Reports.Read.All'); scp = '' }
$missingDirectory = @(Get-GraphMissingScopes -RequiredScopes @(
    'Directory.Read.All', 'Organization.Read.All', 'Group.Read.All', 'User.Read.All', 'GroupMember.Read.All', 'Application.Read.All', 'Reports.Read.All'
))
Assert-Equal @($missingDirectory).Count 0 'A Directory.Read.All app role satisfies the narrower directory reads.'
$script:GraphAccessToken = New-UnsignedTestJwt @{ roles = @('Group.ReadWrite.All'); scp = '' }
$missingGroup = @(Get-GraphMissingScopes -RequiredScopes @('Group.Read.All', 'Organization.Read.All'))
Assert-Equal @($missingGroup).Count 1 'Group.ReadWrite.All leaves Organization.Read.All missing.'
Assert-Equal $missingGroup[0] 'organization.read.all' 'The remaining missing scope is organization.read.all.'

# CSV bytes with a UTF-8 BOM decode to rows instead of an empty report.
$csvText = "Report Refresh Date,Office 365`r`n2026-11-22,42`r`n"
$csvBytes = [byte[]]((0xEF, 0xBB, 0xBF) + [System.Text.Encoding]::UTF8.GetBytes($csvText))
$decodedRows = @(ConvertFrom-GraphReportContent -Content $csvBytes)
Assert-Equal @($decodedRows).Count 1 'BOM-prefixed CSV bytes produce one row.'
Assert-Equal (Get-RowValueNormalized -Row $decodedRows[0] -CandidateNames @('office365', 'office 365')) '42' 'The Office 365 column survives BOM removal.'
Assert-Equal (Get-RowValueNormalized -Row $decodedRows[0] -CandidateNames @('reportRefreshDate', 'report refresh date')) '2026-11-22' 'The report date header is not BOM-prefixed.'

# Copilot active users are lastActivityDate within the period, not licensed-user row counts.
$copilotDetail = @(
    [pscustomobject]@{ 'Report Refresh Date' = '2026-11-22'; 'User Principal Name' = 'active@contoso.com'; 'Last Activity Date' = '2026-11-20' }
    [pscustomobject]@{ 'Report Refresh Date' = '2026-11-22'; 'User Principal Name' = 'stale@contoso.com'; 'Last Activity Date' = '2026-01-01' }
    [pscustomobject]@{ 'Report Refresh Date' = '2026-11-22'; 'User Principal Name' = 'licensed-only@contoso.com'; 'Last Activity Date' = '' }
)
$copilotActive = Get-CopilotActiveUserCount -Rows $copilotDetail -Period 'D30' -ReportDate ([datetime]'2026-11-22T00:00:00Z')
Assert-Equal $copilotActive 1 'Only the user active inside D30 is counted.'
$summaryRow = [pscustomobject]@{ 'Any App Enabled Users' = 100; 'Any App Active Users' = 5; 'Report Period' = 30 }
Assert-Equal (Get-CopilotSummaryActiveUserCount -Rows @($summaryRow) -Period 'D30') 5 'Summary active users are not the enabled/licensed count.'
$nestedSummary = [pscustomobject]@{
    reportRefreshDate = '2026-11-22'
    adoptionByProduct = @([pscustomobject]@{ reportPeriod = 30; anyAppEnabledUsers = 100; anyAppActiveUsers = 4 })
}
Assert-Equal (Get-CopilotSummaryActiveUserCount -Rows @($nestedSummary) -Period 'D30') 4 'Nested summary JSON uses anyAppActiveUsers.'
$candidateUris = @(Get-CopilotReportCandidateUris -Period 'D30' | ForEach-Object { [string]$_.Uri })
Assert-True (($candidateUris -join "`n") -match 'getMicrosoft365CopilotUserCountSummary') 'D30 candidates include getMicrosoft365CopilotUserCountSummary.'
Assert-True (($candidateUris -join "`n") -match 'getMicrosoft365CopilotUsageUserDetail') 'D30 candidates include getMicrosoft365CopilotUsageUserDetail.'
Assert-True (-not (($candidateUris -join "`n") -match 'getMicrosoft365CopilotUsageUserCounts')) 'The non-existent UsageUserCounts report is not requested.'
Assert-True (($candidateUris -join "`n") -match "period='D30'") 'The beta reports path keeps D30.'
Assert-True (($candidateUris -join "`n") -match "period='D28'") 'The copilot reports path maps D30 to the documented D28 period.'
foreach ($periodName in @('D7', 'D90', 'D180')) {
    $periodUris = @(Get-CopilotReportCandidateUris -Period $periodName | ForEach-Object { [string]$_.Uri })
    Assert-True (($periodUris -join "`n") -match "period='$periodName'") "$periodName is sent to the documented Copilot report endpoints."
}

# Secure Score open controls are scored gaps, not the whole catalog.
$secureProfiles = @(
    [pscustomobject]@{ id = 'mfa'; title = 'Require MFA'; maxScore = 9; rank = 5; controlCategory = 'Identity'; controlStateUpdates = @() }
    [pscustomobject]@{ id = 'legacy'; title = 'Block legacy'; maxScore = 8; rank = 1; controlCategory = 'Identity'; controlStateUpdates = @([pscustomobject]@{ state = 'ThirdParty' }) }
    [pscustomobject]@{ id = 'done'; title = 'Fully achieved'; maxScore = 10; rank = 2; controlCategory = 'Data'; controlStateUpdates = @() }
    [pscustomobject]@{ id = 'bigger'; title = 'Larger gap'; maxScore = 20; rank = 9; controlCategory = 'Apps'; controlStateUpdates = @() }
    [pscustomobject]@{ id = 'catalog-only'; title = 'Not scored'; maxScore = 50; rank = 1; controlCategory = 'Device'; controlStateUpdates = @() }
)
$secureScores = @(
    [pscustomobject]@{ controlName = 'mfa'; score = 3 }
    [pscustomobject]@{ controlName = 'legacy'; score = 1 }
    [pscustomobject]@{ controlName = 'done'; score = 10 }
    [pscustomobject]@{ controlName = 'bigger'; score = 2 }
)
$secureAssessment = Get-SecureScoreOpenControlAssessment -ControlProfiles $secureProfiles -ControlScores $secureScores
Assert-Equal $secureAssessment.OpenControlCount 2 'Open controls are scored, have points available, and are not addressed.'
Assert-Equal $secureAssessment.NotApplicableCount 1 'Catalog profiles absent from controlScores are not applicable.'
Assert-Equal $secureAssessment.AddressedWithPointsRemaining 1 'Addressed controls with a remaining gap are counted separately.'
Assert-Equal $secureAssessment.FullyAchievedCount 1 'A scored control at max score is fully achieved.'
Assert-Equal $secureAssessment.TotalPointsAvailable 31 'Points available sum the scored gaps, including addressed ones (6 + 7 + 0 + 18).'
Assert-Equal $secureAssessment.Open[0].id 'bigger' 'Open controls sort by pointsAvailable descending.'
Assert-Equal $secureAssessment.Open[0].pointsAvailable 18 'The larger gap is maxScore minus the current control score.'
Assert-Equal $secureAssessment.Open[1].pointsAvailable 6 'The MFA gap uses the profile max score when the controlScore has no maxScore.'
Assert-True ($secureAssessment.OpenControlCount -ne @($secureProfiles).Count) 'Open is not the whole control catalog.'

# DateTime values stop the not-collected walk instead of recursing through Date.
$dateWalk = [pscustomobject]@{
    Status = 'reported'
    When = [datetime]'2026-11-22T00:00:00Z'
    Nested = [ordered]@{
        Note = 'not collected'
        Stamp = [datetimeoffset]::UtcNow
    }
}
$dateWatch = [System.Diagnostics.Stopwatch]::StartNew()
$datePaths = @(Get-NotCollectedFieldPaths -InputObject $dateWalk)
$directDatePaths = @(Get-NotCollectedFieldPaths -InputObject ([datetime]::UtcNow))
$enumPaths = @(Get-NotCollectedFieldPaths -InputObject ([SharingCapabilities]::ExternalUserAndGuestSharing))
$dateWatch.Stop()
Assert-Equal @($datePaths).Count 1 'Only the nested not-collected string is reported.'
Assert-Equal $datePaths[0] 'Nested.Note' 'The not-collected path keeps its field name.'
Assert-Equal @($directDatePaths).Count 0 'A DateTime value produces no paths.'
Assert-Equal @($enumPaths).Count 0 'An enum value produces no paths.'
Assert-True ($dateWatch.ElapsedMilliseconds -lt 2000) 'DateTime recursion returns immediately.'

# Labels that were not collected stay null. A collected zero stays zero.
$emptyLabels = Get-SensitivityLabelSnapshot -GraphCollected $false -ComplianceSessionConnected $false
Assert-Equal $emptyLabels.Collected $false 'Labels without Graph or Purview are not collected.'
Assert-Equal $emptyLabels.Published $null 'A missing published-label count is null.'
Assert-Equal $emptyLabels.Total $null 'A missing label total is null.'
Assert-True (-not [string]::IsNullOrWhiteSpace([string]$emptyLabels.Reason)) 'The null label count includes a reason.'

$zeroLabels = Get-SensitivityLabelSnapshot -GraphCollected $true -GraphLabels @()
Assert-Equal $zeroLabels.Published 0 'A successful empty Graph label response remains zero.'

$purviewLabels = Get-SensitivityLabelSnapshot -GraphCollected $false -ComplianceSessionConnected $true -ComplianceLabels @(
    [pscustomobject]@{ Name = 'Public' }
    [pscustomobject]@{ Name = 'Confidential' }
) -ComplianceLabelPolicies @(
    [pscustomobject]@{ Name = 'Default'; Labels = @('Public', 'Confidential', 'Public') }
) -AutoSensitivityLabelPolicies @(
    [pscustomobject]@{ Name = 'Auto finance' }
)
Assert-Equal $purviewLabels.Total 2 'Purview Get-Label supplies the label total when Graph is unavailable.'
Assert-Equal $purviewLabels.Published 2 'Published labels come from Get-LabelPolicy, with duplicates removed.'
Assert-Equal $purviewLabels.AutoLabel 1 'Get-AutoSensitivityLabelPolicy supplies the auto-label count.'

$script:Report = [ordered]@{
    DataGovernance = [ordered]@{
        Status = 'partial'
        Reason = 'Sensitivity labels were not collected.'
        SensitivityLabelsPublished = $null
        SensitivityLabelsTotal = $null
        LabelCollectionReason = 'Sensitivity labels were not collected.'
        DlpPolicies = 'not collected'
        RetentionPolicies = 'not collected'
    }
}
$labelAdvanced = Get-DataProtectionAdvanced
Assert-Equal $labelAdvanced.LabelPoliciesPublished $null 'Advanced label publication is null when labels were not collected.'
Assert-Equal $labelAdvanced.LabelPoliciesTotal $null 'Advanced label total is null when labels were not collected.'
Assert-True ($labelAdvanced.LabelPoliciesPublished -ne 0) 'A missing label count is not reported as zero.'
        }
    }

    Context 'Remaining headless collection bugs' {
        It 'does not import PnP.PowerShell on the parent connection path' {
            $importCommands = @($ast.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Import-Module'
            }, $true))
            $pnpImports = @($importCommands | Where-Object { $_.Extent.Text -match 'PnP\.PowerShell' })
            Assert-Equal @($pnpImports).Count 0 'The parent script has no Import-Module PnP.PowerShell command.'

            $connectServices = @($ast.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Connect-Services'
            }, $false)) | Select-Object -First 1
            Assert-True ($null -ne $connectServices) 'Connect-Services is defined.'
            Assert-True ($connectServices.Body.Extent.Text -notmatch 'Import-Module PnP') 'Connect-Services does not import PnP.'
            Assert-True ($connectServices.Body.Extent.Text -notmatch 'Import-SilentPnPModule') 'Connect-Services does not load the SharePoint child module helper.'
            $exchangeIndex = $connectServices.Body.Extent.Text.IndexOf('Connect-ExchangeOnline')
            $ippsIndex = $connectServices.Body.Extent.Text.IndexOf('Connect-IPPSSession')
            $teamsIndex = $connectServices.Body.Extent.Text.IndexOf('Connect-MicrosoftTeams')
            $exchangeModuleImport = $connectServices.Body.Extent.Text.IndexOf('Import-Module ExchangeOnlineManagement')
            $teamsModuleImport = $connectServices.Body.Extent.Text.IndexOf('Import-Module MicrosoftTeams')
            $lastExchangeConnect = $connectServices.Body.Extent.Text.LastIndexOf('Connect-ExchangeOnline')
            $lastIppsConnect = $connectServices.Body.Extent.Text.LastIndexOf('Connect-IPPSSession')
            Assert-True ($exchangeIndex -ge 0 -and $ippsIndex -ge 0 -and $teamsIndex -ge 0) 'Exchange, Purview, and Teams connections are present.'
            Assert-True ($exchangeIndex -lt $teamsIndex -and $ippsIndex -lt $teamsIndex) 'Exchange and Purview connect before Teams, and before any SharePoint child work.'
            Assert-True ($exchangeModuleImport -ge 0 -and $teamsModuleImport -gt $exchangeModuleImport) 'ExchangeOnlineManagement is imported before MicrosoftTeams so its MSAL assembly loads first.'
            Assert-True ($lastExchangeConnect -lt $teamsModuleImport -and $lastIppsConnect -lt $teamsModuleImport) 'Exchange and Purview connect before the MicrosoftTeams module is imported.'
            Assert-True ($connectServices.Body.Extent.Text -match "Set-ServiceConnectionResult -Name 'Exchange' -Status 'failed'") 'A failed Exchange connection is recorded.'
            Assert-True ($connectServices.Body.Extent.Text -match "Set-ServiceConnectionResult -Name 'Compliance' -Status 'failed'") 'A failed Purview connection is recorded.'

            $silentImport = @($ast.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Import-SilentPnPModule'
            }, $false)) | Select-Object -First 1
            Assert-True ($silentImport.Body.Extent.Text -match 'PnP\.PowerShell') 'The SharePoint child helper still loads PnP.PowerShell.'
        }

        It 'emits not-collected site counts and scores tenant anyone-link capability with lower confidence' {
            $script:Report = [ordered]@{
                SharePointOneDrive = [ordered]@{
                    Status = 'partial'
                    ExternalSharingLevel = 'ExternalUserAndGuestSharing'
                    SiteInventoryStatus = 'not collected'
                    SiteInventoryReason = 'Site inventory was not collected.'
                    TotalSiteCount = 'not collected'
                    Sites = @()
                    RestrictedSearchEnabled = 'not collected'
                    OneDrivePersonalSiteSharePercent = 'not collected'
                }
            }
            $exposure = Get-SharePointExposureAdvanced
            Assert-Equal $exposure.AnyoneLinkSites 'not collected' 'Missing site inventory does not emit zero anyone-link sites.'
            Assert-Equal $exposure.ExternalSharingSites 'not collected' 'Missing site inventory does not emit zero external-sharing sites.'
            Assert-Equal $exposure.AnyoneLinksPossible $true 'ExternalUserAndGuestSharing means anyone links are possible.'
            Assert-Equal $exposure.AnyoneLinkMeasurement 'tenant-settings' 'The anyone-link signal is marked as tenant settings.'
            Assert-True ($exposure.Reason -match 'not collected') 'The exposure reason says the site counts were not collected.'

            $script:Report = [ordered]@{
                ReadinessFlags = [ordered]@{
                    MfaRegisteredPercent = 'not collected'
                    CaPoliciesEnabled = 'not collected'
                    SensitivityLabelsPublished = 'not collected'
                    OneDrivePersonalSiteSharePercent = 'not collected'
                    RestrictedSearchEnabled = 'not collected'
                    ActiveUsersInReportPeriodPercent = 'not collected'
                    ExternalSharingLevel = 'ExternalUserAndGuestSharing'
                }
                Identity = [ordered]@{ Status = 'reported' }
                ConditionalAccess = [ordered]@{ Status = 'reported' }
                DataGovernance = [ordered]@{ Status = 'reported' }
                SharePointOneDrive = [ordered]@{ Status = 'reported' }
                AdoptionSignals = [ordered]@{ Status = 'reported' }
                IdentityAccessAdvanced = [ordered]@{ Status = 'reported' }
                DataProtectionAdvanced = [ordered]@{ Status = 'reported' }
                SharePointExposureAdvanced = $exposure
                AppGovernanceAdvanced = [ordered]@{ Status = 'reported' }
                EndpointAppAdvanced = [ordered]@{ Status = 'reported' }
                AdoptionAdvanced = [ordered]@{ Status = 'reported' }
            }
            $script:ScoringModel = $null
            $tenantEvidence = Get-ReadinessEvidence
            Assert-Equal $tenantEvidence.DomainScores.ContentExposure 0 'Tenant anyone-link capability scores as anyone links present.'
            Assert-Equal $tenantEvidence.DataCompleteness.AnyoneLinkConfidence 'reduced' 'Tenant-only anyone-link scoring lowers confidence.'
            Assert-Equal $tenantEvidence.DataCompleteness.CompletenessPercent 3.3 'Tenant-only anyone links contribute half of that metric weight.'
            Assert-True ($tenantEvidence.RiskSignals.AnyoneLinkSites -eq 'not collected') 'Risk signals do not turn a missing site count into zero.'

            $script:Report.SharePointExposureAdvanced = [ordered]@{
                AnyoneLinkSites = 0
                AnyoneLinksPossible = $false
                AnyoneLinkMeasurement = 'site-inventory'
            }
            $siteEvidence = Get-ReadinessEvidence
            Assert-Equal $siteEvidence.DomainScores.ContentExposure 100 'A collected zero anyone-link count still receives the full weight.'
            Assert-Equal $siteEvidence.DataCompleteness.CompletenessPercent 6.6 'A measured site count contributes the full anyone-link weight.'
            Assert-True ($tenantEvidence.DataCompleteness.CompletenessPercent -lt $siteEvidence.DataCompleteness.CompletenessPercent) 'Tenant-only evidence is a lower completeness percent than a site inventory.'
        }

        It 'reads Office 365 Active from the services report and keeps blanks null' {
            $blankDaily = [pscustomobject]@{ 'Report Date' = '2026-11-22'; 'Office 365' = ''; Teams = '9' }
            Assert-Equal (Get-ActiveUserCountFromRow -Row $blankDaily -CandidateNames @('Office 365', 'office365')) $null 'A blank Office 365 cell is null.'
            Assert-True ($null -eq (Get-ActiveUserCountFromRow -Row $blankDaily -CandidateNames @('Office 365', 'office365'))) 'A blank Office 365 cell is not zero.'
            Assert-Equal (Get-ActiveUserCountFromRow -Row ([pscustomobject]@{ 'Office 365 Active' = '0' }) -CandidateNames @('Office 365 Active')) 0 'A real zero stays zero.'

            $script:ServicesUserCountRow = [pscustomobject]@{
                'Report Refresh Date' = '2026-11-22'
                'Office 365 Active' = '120'
                'Office 365 Inactive' = '30'
                'Teams Active' = '80'
                'SharePoint Active' = '70'
                'OneDrive Active' = '60'
                'Exchange Active' = '90'
            }
            function Get-GraphReportRows {
                param([Parameter(Mandatory)][string]$Uri, [hashtable]$Query = @{}, [switch]$DiagnosticOnly)
                if ($Uri -match "getOffice365ServicesUserCounts\(period='D30'\)") { return @($script:ServicesUserCountRow) }
                if ($Uri -match 'getOffice365ActiveUserCounts') {
                    return @([pscustomobject]@{ 'Report Date' = '2026-11-22'; 'Office 365' = '' })
                }
                return @()
            }
            function Get-CopilotUsageSnapshot {
                param([string]$Period = 'D30')
                return [ordered]@{ Status = 'not collected'; ActiveUsers = 'not collected'; Reason = ''; Source = 'not collected'; ReportDate = 'not collected' }
            }

            $adoption = Get-AdoptionSignals
            Assert-Equal $adoption.ActiveUsers30Days 120 'D30 Office 365 Active comes from getOffice365ServicesUserCounts.'
            Assert-Equal $adoption.TeamsActiveUsers 80 'Teams Active comes from the services report.'
            Assert-Equal $adoption.SharePointActiveUsers 70 'SharePoint Active comes from the services report.'
            Assert-Equal $adoption.OneDriveActiveUsers 60 'OneDrive Active comes from the services report.'
            Assert-Equal $adoption.OutlookActiveUsers 90 'Exchange Active comes from the services report.'
            Assert-Equal $adoption.ReportSource 'office365ServicesUserCounts' 'The report source names the services user count report.'
            Assert-True ($adoption.ActiveUsers30Days -ne 0) 'The services count is not the blank daily cell coerced to zero.'
        }

        It 'tries a supported Copilot period first and records fallback probes as diagnostics' {
            $candidates = @(Get-CopilotReportCandidateUris -Period 'D30')
            Assert-True ($candidates[0].Uri -match "period='D28'") 'The first Copilot candidate uses D28.'
            Assert-True ($candidates[0].Uri -match '/copilot/reports/') 'The first Copilot candidate uses the copilot reports path.'
            Assert-True ($candidates[0].Uri -notmatch "period='D30'") 'The first Copilot candidate does not send the rejected D30 period.'

            $script:Report = [ordered]@{ Errors = [System.Collections.Generic.List[object]]::new() }
            $script:GraphReportDiagnostics = [System.Collections.Generic.List[object]]::new()
            $script:GraphReportRowsCache = @{}
            $script:CopilotUsageCache = @{}
            $script:CopilotReportCalls = 0
            function Test-GraphPermissionsAvailable { param([string[]]$RequiredPermissions) return $true }
            function Get-GraphReportRows {
                param([Parameter(Mandatory)][string]$Uri, [hashtable]$Query = @{}, [switch]$DiagnosticOnly)
                $script:CopilotReportCalls++
                if ($script:CopilotReportCalls -eq 1) {
                    Register-GraphReportDiagnostic -Uri $Uri -Message 'Response status code does not indicate success: 400 (Bad Request).' -StatusCode 400 -DiagnosticOnly:$DiagnosticOnly
                    return @()
                }
                return @([pscustomobject]@{
                    'Report Refresh Date' = '2026-11-22'
                    'User Principal Name' = 'active@contoso.com'
                    'Last Activity Date' = '2026-11-20'
                })
            }

            $snapshot = Get-CopilotUsageSnapshot -Period 'D30'
            Assert-Equal $snapshot.Status 'reported' 'A later Copilot candidate can still succeed.'
            Assert-Equal @($script:Report.Errors).Count 0 'The failed first probe is not a collector error.'
            Assert-True ((@($snapshot.ReportDiagnostics) -join ' ') -match '400') 'The failed probe is retained as a diagnostic.'
        }

        It 'applies the app-only certificate fallback when the mode condition is parenthesized' {
            $script:NonInteractive = $false
            $script:Headless = $false
            $script:GraphAuthenticationMode = 'AppOnlyCertificate'
            $script:GraphCertificateThumbprint = 'DEADBEEF'
            $script:GraphCertificatePath = ''
            $script:GraphCertificatePasswordEnvironmentVariable = 'GRAPH_CERT_PASSWORD'
            $script:GraphClientId = '11111111-2222-3333-4444-555555555555'
            $script:TenantId = 'contoso.onmicrosoft.com'

            $selection = Get-WorkloadCertificateSelection -AppId '' -Organization 'contoso.onmicrosoft.com' -Thumbprint '' -Path '' -SecretVariableName '' -AllowGraphFallback
            Assert-Equal $selection.Thumbprint 'DEADBEEF' 'App-only mode falls back to the Graph certificate.'
            Assert-Equal $selection.AppId '11111111-2222-3333-4444-555555555555' 'App-only mode falls back to the Graph client id.'

            $script:GraphAuthenticationMode = 'DeviceCode'
            $noFallback = Get-WorkloadCertificateSelection -AppId '' -Organization 'contoso.onmicrosoft.com' -Thumbprint '' -Path '' -SecretVariableName '' -AllowGraphFallback
            Assert-Equal $noFallback $null 'Device-code mode without non-interactive execution does not take the certificate fallback.'
        }

        It 'keeps addressed Secure Score controls with remaining points on the gap list' {
            $profiles = @(
                [pscustomobject]@{ id = 'MFARegistrationV2'; title = 'MFA registration'; maxScore = 8.93; rank = 4; controlCategory = 'Identity'; controlStateUpdates = @([pscustomobject]@{ state = 'ignored' }) }
                [pscustomobject]@{ id = 'done'; title = 'Fully achieved'; maxScore = 5; rank = 1; controlCategory = 'Data'; controlStateUpdates = @() }
            )
            $scores = @(
                [pscustomobject]@{ controlName = 'MFARegistrationV2'; score = 0 }
                [pscustomobject]@{ controlName = 'done'; score = 5 }
            )
            $assessment = Get-SecureScoreOpenControlAssessment -ControlProfiles $profiles -ControlScores $scores
            Assert-Equal $assessment.OpenControlCount 0 'An ignored control is not an open control.'
            Assert-Equal $assessment.AddressedWithPointsRemaining 1 'The ignored control is counted as addressed with points remaining.'
            $gap = @($assessment.Gaps | Where-Object { $_.id -eq 'MFARegistrationV2' }) | Select-Object -First 1
            Assert-True ($null -ne $gap) 'MFARegistrationV2 is on the gap list.'
            Assert-Equal $gap.gapClassification 'addressed-with-points-remaining' 'The gap is flagged as addressed with points remaining.'
            Assert-Equal $gap.pointsAvailable 8.93 'The remaining points stay on the gap.'
            Assert-Equal $gap.addressedState 'ignored' 'The addressed state is preserved.'
            $topGap = @($assessment.TopGaps | Where-Object { $_.id -eq 'MFARegistrationV2' }) | Select-Object -First 1
            Assert-True ($null -ne $topGap) 'The addressed gap is included in the top gap list.'
        }

        It 'uses enabled member users as the licence-readiness denominator' {
            $script:GraphUsersCache = @{}
            function Get-GraphRestAll {
                param([string]$Uri, [hashtable]$Query = @{}, [switch]$BypassCache)
                if ($Uri -match '/subscribedSkus$') {
                    return @([pscustomobject]@{ SkuPartNumber = 'SPE_E5'; EnabledUnits = 10; ConsumedUnits = 4 })
                }
                if ($Uri -match '/users$') {
                    return @(
                        [pscustomobject]@{
                            Id = 'member-ready'
                            userType = 'Member'
                            accountEnabled = $true
                            assignedPlans = @(
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'MicrosoftOffice' }
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'exchange' }
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'SharePointEnterprise' }
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'TeamspaceAPI' }
                            )
                        }
                        [pscustomobject]@{
                            Id = 'member-partial'
                            userType = 'Member'
                            accountEnabled = $true
                            assignedPlans = @(
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'exchange' }
                            )
                        }
                        [pscustomobject]@{
                            Id = 'member-disabled'
                            userType = 'Member'
                            accountEnabled = $false
                            assignedPlans = @(
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'MicrosoftOffice' }
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'exchange' }
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'SharePointEnterprise' }
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'TeamspaceAPI' }
                            )
                        }
                        [pscustomobject]@{
                            Id = 'guest-ready'
                            userType = 'Guest'
                            accountEnabled = $true
                            assignedPlans = @(
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'MicrosoftOffice' }
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'exchange' }
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'SharePointEnterprise' }
                                [pscustomobject]@{ capabilityStatus = 'Enabled'; service = 'TeamspaceAPI' }
                            )
                        }
                    )
                }
                return @()
            }
            function Test-GraphPermissionsAvailable { param([string[]]$RequiredPermissions) return $true }

            $licensing = Get-LicensingReadiness
            Assert-Equal $licensing.PrereqReadyUserCount 1 'Only the enabled member with all four services is ready.'
            Assert-Equal $licensing.PrereqReadyDenominator 2 'The denominator is the two enabled members.'
            Assert-Equal $licensing.PrereqReadyPercent 50 '1 of 2 enabled members is 50 percent.'
            Assert-True ($licensing.PrereqReadyDenominatorDefinition -match 'Guests are excluded') 'The denominator definition is exposed.'
        }

        It 'weights completeness by measured metrics instead of collected sections' {
            $script:ScoringModel = $null
            $sections = [ordered]@{
                ReadinessFlags = [ordered]@{
                    MfaRegisteredPercent = 'not collected'
                    CaPoliciesEnabled = 'not collected'
                    SensitivityLabelsPublished = 'not collected'
                    OneDrivePersonalSiteSharePercent = 'not collected'
                    RestrictedSearchEnabled = 'not collected'
                    ActiveUsersInReportPeriodPercent = 'not collected'
                }
                Identity = [ordered]@{ Status = 'reported' }
                ConditionalAccess = [ordered]@{ Status = 'reported' }
                DataGovernance = [ordered]@{ Status = 'reported' }
                SharePointOneDrive = [ordered]@{ Status = 'reported' }
                AdoptionSignals = [ordered]@{ Status = 'reported' }
                IdentityAccessAdvanced = [ordered]@{ Status = 'reported' }
                DataProtectionAdvanced = [ordered]@{ Status = 'reported' }
                SharePointExposureAdvanced = [ordered]@{
                    Status = 'reported'
                    AnyoneLinkSites = 0
                    AnyoneLinkMeasurement = 'site-inventory'
                }
                AppGovernanceAdvanced = [ordered]@{ Status = 'reported' }
                EndpointAppAdvanced = [ordered]@{ Status = 'reported' }
                AdoptionAdvanced = [ordered]@{ Status = 'reported' }
            }
            $script:Report = $sections
            $evidence = Get-ReadinessEvidence
            Assert-Equal $evidence.DataCompleteness.CompletenessScope 'metric-level' 'Completeness is calculated from metrics.'
            Assert-Equal $evidence.DataCompleteness.PossibleMetricWeight 500 'Every scoring metric weight is in the denominator.'
            Assert-Equal $evidence.DataCompleteness.MeasuredMetricWeight 33 'Only the anyone-link metric weight was measured.'
            Assert-Equal $evidence.DataCompleteness.CompletenessPercent 6.6 'Collected sections do not make a one-metric domain 100 percent complete.'
            Assert-Equal $evidence.DomainScores.ContentExposure 100 'The measured anyone-link metric still scores from its own weight.'
            Assert-Equal $evidence.OverallScore 100 'The overall score averages measured domains.'
            Assert-Equal $evidence.ConfidenceAdjustedScore 6.6 'The confidence-adjusted score uses metric completeness.'
            Assert-True ($evidence.DataCompleteness.CompletenessPercent -ne 100) 'Section-level collection is not reported as full completeness.'
            Assert-Equal $evidence.DataCompleteness.ConfidenceLevel 'low' 'A small measured weight is low confidence.'
        }

        It 'saves retention and DLP policies as lists and keeps label policy counts separate' {
            function Get-ConnectionInformation { [pscustomobject]@{ Name = 'IPPSSession' } }
            function Get-DlpCompliancePolicy {
                [CmdletBinding()]
                param()
                [pscustomobject]@{ Name = 'DLP baseline'; Mode = 'Enable'; Workload = 'Exchange, SharePoint' }
            }
            function Get-RetentionCompliancePolicy {
                [CmdletBinding()]
                param()
                [pscustomobject]@{ Name = 'Finance hold'; Enabled = $true; Mode = 'Enforce'; Workload = 'SharePoint, OneDriveForBusiness' }
            }
            function Get-GraphRestAll {
                param([string]$Uri, [hashtable]$Query = @{}, [switch]$BypassCache)
                $script:GraphLastHttpStatus = 403
                return @()
            }

            $governance = Get-DataGovernanceReadiness
            Assert-True ($governance.DlpPolicies -is [System.Array]) 'DLP policies are saved as a list.'
            Assert-Equal @($governance.DlpPolicies).Count 1 'The single DLP policy stays in a list.'
            Assert-Equal $governance.DlpPolicies[0].State 'Enable' 'DLP State is populated from Mode when State is absent.'
            Assert-Equal $governance.DlpPolicies[0].Enabled $true 'An Enable-mode DLP policy is enabled.'
            Assert-Equal $governance.DlpPolicies[0].Workloads 'Exchange, SharePoint' 'DLP workloads are kept.'
            Assert-True ($governance.RetentionPolicies -is [System.Array]) 'Retention policies are saved as a list.'
            Assert-Equal $governance.RetentionPolicies[0].Name 'Finance hold' 'The retention policy name is kept.'
            Assert-Equal $governance.RetentionPolicies[0].Enabled $true 'The retention policy enabled flag is kept.'
            Assert-Equal $governance.RetentionPolicies[0].Mode 'Enforce' 'The retention policy mode is kept.'
            Assert-Equal $governance.RetentionPolicies[0].Workloads 'SharePoint, OneDriveForBusiness' 'Retention workloads are kept.'
            $policyJson = ([ordered]@{ DlpPolicies = $governance.DlpPolicies; RetentionPolicies = $governance.RetentionPolicies } | ConvertTo-Json -Depth 6 -Compress)
            Assert-True ($policyJson -match '"DlpPolicies":\[\{') 'DLP policies serialize as one JSON array.'
            Assert-True ($policyJson -notmatch '"DlpPolicies":\[\[') 'DLP policies are not a nested JSON array.'
            Assert-True ($policyJson -match '"RetentionPolicies":\[\{') 'Retention policies serialize as one JSON array.'
            Assert-True ($policyJson -notmatch '"RetentionPolicies":\[\[') 'Retention policies are not a nested JSON array.'

            $snapshot = Get-SensitivityLabelSnapshot -GraphCollected $false -ComplianceSessionConnected $true -ComplianceLabels @(
                [pscustomobject]@{ Name = 'Public' }
                [pscustomobject]@{ Name = 'Confidential' }
            ) -ComplianceLabelPolicies @(
                [pscustomobject]@{ Name = 'Publish default'; Labels = @('Public', 'Confidential') }
                [pscustomobject]@{ Name = 'Publish public'; Labels = @('Public') }
                [pscustomobject]@{ Name = 'Empty policy'; Labels = @() }
            )
            Assert-Equal $snapshot.Total 2 'Label total counts labels.'
            Assert-Equal $snapshot.Published 2 'Published label count is unique label names.'
            Assert-Equal $snapshot.PolicyTotal 3 'Label policy total counts policies.'
            Assert-Equal $snapshot.PolicyPublished 2 'Label policies that publish labels are counted separately from labels.'

            $script:Report = [ordered]@{
                DataGovernance = [ordered]@{
                    Status = 'reported'
                    SensitivityLabelsPublished = 12
                    SensitivityLabelsTotal = 20
                    LabelPoliciesPublished = $snapshot.PolicyPublished
                    LabelPoliciesTotal = $snapshot.PolicyTotal
                    DlpPolicies = 'not collected'
                    RetentionPolicies = 'not collected'
                }
            }
            $advanced = Get-DataProtectionAdvanced
            Assert-Equal $advanced.SensitivityLabelsPublished 12 'Advanced output keeps the label publication count.'
            Assert-Equal $advanced.SensitivityLabelsTotal 20 'Advanced output keeps the label total.'
            Assert-Equal $advanced.LabelPoliciesPublished 2 'Label policy publication is not the label count.'
            Assert-Equal $advanced.LabelPoliciesTotal 3 'Label policy total is not the label total.'
        }

        It 'reads AllowCloudRecording from the Teams meeting policy' {
            $script:TeamsConnected = $true
            $script:GraphLastHttpStatus = 200
            function Test-GraphPermissionsAvailable { param([string[]]$RequiredPermissions) return $true }
            function Get-GraphRestAll { param([string]$Uri, [hashtable]$Query = @{}, [switch]$BypassCache) return @() }
            function Get-CsTeamsMeetingPolicy {
                [CmdletBinding()]
                param([string]$Identity)
                [pscustomobject]@{ AllowCloudRecording = $true; AllowTranscription = $false }
            }

            $teams = Get-TeamsReadiness
            Assert-Equal $teams.RecordingEnabled $true 'AllowCloudRecording populates RecordingEnabled.'
            Assert-Equal $teams.TranscriptionEnabled $false 'AllowTranscription still populates TranscriptionEnabled.'
            Assert-Equal $teams.MeetingPolicyState 'enabled' 'Cloud recording makes the meeting policy enabled.'
        }

        It 'keeps compliance policy lists flat for empty, one, and many policies' {
            $empty = ConvertTo-CompliancePolicyList -Policies @()
            Assert-True ($empty -is [System.Array]) 'An empty policy result is an array.'
            Assert-Equal @($empty).Count 0 'An empty policy result has no items.'
            Assert-Equal (([ordered]@{ DlpPolicies = $empty } | ConvertTo-Json -Depth 5 -Compress)) '{"DlpPolicies":[]}' 'An empty policy list serializes as [].'

            $one = ConvertTo-CompliancePolicyList -Policies @([pscustomobject]@{ Name = 'One'; Mode = 'Enable'; Workload = 'Exchange' })
            Assert-Equal $one[0].Name 'One' 'A single policy is the first element, not a nested array.'
            Assert-True ($one[0] -isnot [System.Array]) 'A single policy element is the policy object.'

            $two = ConvertTo-CompliancePolicyList -Policies @(
                [pscustomobject]@{ Name = 'One'; Mode = 'Enable'; Workload = 'Exchange' }
                [pscustomobject]@{ Name = 'Two'; Mode = 'Enforce'; Workload = 'SharePoint' }
            )
            Assert-Equal @($two).Count 2 'Two policies stay in one list.'
            $twoJson = ([ordered]@{ DlpPolicies = $two } | ConvertTo-Json -Depth 5 -Compress)
            Assert-True ($twoJson -match '"DlpPolicies":\[\{') 'Several policies serialize as one JSON array.'
            Assert-True ($twoJson -notmatch '"DlpPolicies":\[\[') 'Several policies are not nested.'
        }

        It 'scores enabled-member MFA and divides adoption by enabled members' {
            $directoryUsers = [System.Collections.Generic.List[object]]::new()
            $registration = [System.Collections.Generic.List[object]]::new()
            foreach ($index in 1..10) {
                $directoryUsers.Add([pscustomobject]@{ id = "member-$index"; userType = 'Member'; accountEnabled = $true }) | Out-Null
                $registration.Add([pscustomobject]@{ id = "member-$index"; userType = 'member'; isMfaRegistered = ($index -le 4) }) | Out-Null
            }
            $directoryUsers.Add([pscustomobject]@{ id = 'disabled-member'; userType = 'Member'; accountEnabled = $false }) | Out-Null
            $directoryUsers.Add([pscustomobject]@{ id = 'guest-user'; userType = 'Guest'; accountEnabled = $true }) | Out-Null
            $registration.Add([pscustomobject]@{ id = 'disabled-member'; userType = 'member'; isMfaRegistered = $true }) | Out-Null
            $registration.Add([pscustomobject]@{ id = 'guest-user'; userType = 'guest'; isMfaRegistered = $true }) | Out-Null
            $script:GraphUsersCache = @{ base = $directoryUsers.ToArray() }
            $enabledMemberMfa = Get-EnabledMemberMfaRegistration -RegistrationDetails $registration.ToArray()
            Assert-Equal $enabledMemberMfa.Population 10 'The MFA population is the ten enabled members.'
            Assert-Equal $enabledMemberMfa.Registered 4 'Disabled members and guests are not counted as registered enabled members.'
            Assert-Equal $enabledMemberMfa.Percent 40 '4 of 10 enabled members is 40 percent.'

            $script:Report = [ordered]@{
                Metadata = [ordered]@{ TotalUsers = 265 }
                Licensing = [ordered]@{ PrereqReadyDenominator = 10; PrereqReadyPercent = 'not collected' }
                Identity = [ordered]@{
                    MfaRegisteredCount = 7
                    MfaPopulationUserCount = 251
                    MfaRegisteredEnabledMemberPercent = 40
                    MfaRegisteredMemberPercent = 40
                }
                ConditionalAccess = [ordered]@{}
                DataGovernance = [ordered]@{}
                SharePointOneDrive = [ordered]@{}
                AdoptionSignals = [ordered]@{ ActiveUsers30Days = 3 }
            }
            $flags = Get-ReadinessFlags
            Assert-Equal $flags.MfaRegisteredPercent 2.79 'The all-user MFA percent, including guests, stays informational.'
            Assert-Equal $flags.MfaRegisteredEnabledMemberPercent 40 'The enabled-member MFA percent is available for scoring.'
            Assert-Equal $flags.ActiveUsersInReportPeriodPercent 30 '3 active users over 10 enabled members is 30 percent.'
            Assert-Equal $flags.ActiveUserPercentByWorkload 30 'The workload adoption percent uses the same enabled-member denominator.'
            Assert-Equal $flags.ActiveUserPercentDenominator 10 'The adoption denominator is the enabled member count.'
            Assert-True ($flags.ActiveUsersInReportPeriodPercent -ne 1.13) 'Adoption is not divided by all 265 users.'

            $script:Report.ReadinessFlags = [ordered]@{
                MfaRegisteredPercent = 2.79
                MfaRegisteredEnabledMemberPercent = 90
                CaPoliciesEnabled = 'not collected'
                SensitivityLabelsPublished = 'not collected'
                OneDrivePersonalSiteSharePercent = 'not collected'
                RestrictedSearchEnabled = 'not collected'
                ActiveUsersInReportPeriodPercent = 30
            }
            $script:Report.IdentityAccessAdvanced = [ordered]@{ Status = 'reported'; PhishingResistantMfaCoverage = 'not collected' }
            $script:Report.DataProtectionAdvanced = [ordered]@{ Status = 'reported' }
            $script:Report.SharePointExposureAdvanced = [ordered]@{ Status = 'reported' }
            $script:Report.AppGovernanceAdvanced = [ordered]@{ Status = 'reported' }
            $script:Report.EndpointAppAdvanced = [ordered]@{ Status = 'reported' }
            $script:Report.AdoptionAdvanced = [ordered]@{ Status = 'not collected'; TrendSignal = 'not collected' }
            $script:Report.ReadinessEvidence = [ordered]@{ OverallScore = 'not collected' }
            $script:ScoringModel = $null
            $scored = Get-ReadinessEvidence
            Assert-Equal $scored.IdentityScoreInputs.MfaRegisteredPercent 2.79 'Score inputs keep the all-user MFA percent.'
            Assert-Equal $scored.IdentityScoreInputs.MfaRegisteredEnabledMemberPercent 90 'Score inputs record the enabled-member MFA percent.'
            Assert-Equal $scored.IdentityScoreInputs.MfaScoreBasis 'enabled-members' 'The MFA score basis is enabled members.'
            Assert-Equal $scored.DomainScores.IdentityAccess 100 'An enabled-member MFA percent of 90 scores the MFA weight even though the all-user percent is 2.79.'

            $script:Report.ReadinessFlags.MfaRegisteredEnabledMemberPercent = 40
            $belowThreshold = Get-ReadinessEvidence
            Assert-Equal $belowThreshold.DomainScores.IdentityAccess 0 'An enabled-member MFA percent of 40 is below the 80 percent threshold.'
            Assert-Equal $belowThreshold.IdentityScoreInputs.MfaRegisteredPercent 2.79 'The all-user figure remains informational when the score uses 40 percent.'
            $recommendations = Get-Recommendations
            Assert-True ((@($recommendations.TopRisks) -join "`n") -match 'enabled members') 'The MFA recommendation follows the enabled-member percentage.'
        }

        It 'records a failed Exchange or Purview connection in errors and connection status' {
            $script:Report = [ordered]@{
                Errors = [System.Collections.Generic.List[object]]::new()
                Exchange = [ordered]@{ Status = 'partial'; Reason = 'Exchange Online connection failed: Could not load Microsoft.Identity.Client.dll'; TotalMailboxes = 'not collected' }
                DataGovernance = [ordered]@{ Status = 'partial'; Reason = 'Purview connection failed: Could not load Microsoft.Identity.Client.dll'; DlpPolicies = 'not collected' }
            }
            $script:GraphAccessToken = $null
            $script:MgContext = $null
            $script:TeamsConnected = $false
            $script:SharePointChildCollectionSucceeded = $false
            $script:GraphMissingScopes = @()
            $assemblyError = 'Could not load file or assembly Microsoft.Identity.Client.dll'
            Set-ServiceConnectionResult -Name 'Exchange' -Status 'failed' -Message $assemblyError
            Set-ServiceConnectionResult -Name 'Compliance' -Status 'failed' -Message $assemblyError
            Assert-Equal @($script:Report.Errors).Count 2 'Exchange and Purview failures are collector errors.'
            Assert-Equal $script:Report.Errors[0].Section 'Exchange' 'The Exchange failure is named Exchange.'
            Assert-Equal $script:Report.Errors[1].Section 'Purview' 'The Purview failure is named Purview.'
            Assert-Equal $script:ExchangeConnected $false 'Exchange stays disconnected after a failure.'
            Assert-Equal $script:ComplianceConnected $false 'Purview stays disconnected after a failure.'

            function Get-ConnectionInformation {
                [pscustomobject]@{ Name = 'ExchangeOnline' }
                [pscustomobject]@{ Name = 'IPPSSession' }
            }
            $gaps = Get-PrerequisitesAndGaps
            Assert-Equal $gaps.CurrentPrerequisites.SessionAvailability.ExchangeConnected $false 'A recorded Exchange failure is not cleared by a later session probe.'
            Assert-Equal $gaps.CurrentPrerequisites.SessionAvailability.ComplianceConnected $false 'A recorded Purview failure is not cleared by a later session probe.'
            Assert-Equal $gaps.CurrentPrerequisites.SessionAvailability.ExchangeConnectionStatus 'failed' 'Exchange connection status is failed.'
            Assert-Equal $gaps.CurrentPrerequisites.SessionAvailability.ComplianceConnectionStatus 'failed' 'Purview connection status is failed.'
            Assert-True ($gaps.CurrentPrerequisites.SessionAvailability.ExchangeConnectionError -match 'Microsoft\.Identity\.Client') 'The Exchange connection error is kept.'
            Assert-True ($gaps.CurrentPrerequisites.SessionAvailability.ComplianceConnectionError -match 'Microsoft\.Identity\.Client') 'The Purview connection error is kept.'
        }

        It 'keeps the Purview connection failure on a not-collected data governance section' {
            $script:ComplianceConnectionStatus = 'failed'
            $script:ComplianceConnectionError = 'Could not load file or assembly Microsoft.Identity.Client.dll'
            function Get-ConnectionInformation {
                [CmdletBinding()]
                param()
            }
            function Get-GraphRestAll {
                param([string]$Uri, [hashtable]$Query = @{}, [switch]$BypassCache)
                $script:GraphLastHttpStatus = 403
                return @()
            }

            $governance = Get-DataGovernanceReadiness
            Assert-Equal $governance.Status 'not collected' 'Missing labels and a failed Purview connect are not collected.'
            Assert-True ($governance.Reason -match 'HTTP 403') 'The Graph label error stays in the reason.'
            Assert-True ($governance.Reason -match 'Purview connection failed: Could not load file or assembly Microsoft\.Identity\.Client\.dll') 'The Purview connection failure is included with the Graph label error.'

            $script:GraphLastHttpStatus = 200
            function Get-GraphRestAll {
                param([string]$Uri, [hashtable]$Query = @{}, [switch]$BypassCache)
                $script:GraphLastHttpStatus = 200
                return @([pscustomobject]@{ name = 'Public'; isPublished = $true })
            }
            $partial = Get-DataGovernanceReadiness
            Assert-Equal $partial.Status 'partial' 'Collected labels with a failed Purview connect stay partial.'
            Assert-True ($partial.Reason -match 'Purview connection failed: Could not load file or assembly Microsoft\.Identity\.Client\.dll') 'The partial path still records the Purview connection failure.'
        }

        It 'exits 3 for finished runs with collector errors only when requested' {
            $failOnCollectorError = @($ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'FailOnCollectorError' })
            Assert-Equal @($failOnCollectorError).Count 1 'FailOnCollectorError is a script parameter.'
            $helpText = Get-Content -LiteralPath $scriptPath -Raw
            Assert-True ($helpText -match 'Exit codes:') 'The help documents the exit codes.'
            Assert-True ($helpText -match 'exits 3') 'The help documents exit code 3.'

            Assert-Equal (Get-ExportExitCode -FailedCollectorCount 2 -ProcessedSectionCount 8 -FailOnCollectorError $false) 0 'Without the switch, a finished run with collector errors still exits 0.'
            Assert-Equal (Get-ExportExitCode -FailedCollectorCount 2 -ProcessedSectionCount 0 -FailOnCollectorError $false) 2 'No processed section still exits 2.'
            Assert-Equal (Get-ExportExitCode -FailedCollectorCount 2 -ProcessedSectionCount 8 -FailOnCollectorError $true) 3 'The switch exits 3 when the run finished with collector errors.'
            Assert-Equal (Get-ExportExitCode -FailedCollectorCount 2 -ProcessedSectionCount 0 -FailOnCollectorError $true) 2 'The switch keeps exit 2 when no section was processed.'
            Assert-Equal (Get-ExportExitCode -FailedCollectorCount 0 -ProcessedSectionCount 8 -FailOnCollectorError $true) 0 'The switch does not fail a run that recorded no collector errors.'
        }

        It 'retries an empty label policy read and does not score a second empty result as zero' {
            $script:Report = [ordered]@{ Errors = [System.Collections.Generic.List[object]]::new() }
            $script:GraphReportDiagnostics = [System.Collections.Generic.List[object]]::new()
            $script:LabelPolicyRetryDelaySeconds = 0
            $script:LabelPolicyCalls = 0
            $script:ComplianceConnectionStatus = 'connected'
            $script:ComplianceConnectionError = ''
            $script:GraphLastHttpStatus = 403
            function Get-ConnectionInformation {
                [CmdletBinding()] param()
                [pscustomobject]@{ Name = 'IPPSSession' }
            }
            function Get-GraphRestAll {
                param([string]$Uri, [hashtable]$Query = @{}, [switch]$BypassCache)
                $script:GraphLastHttpStatus = 403
                return @()
            }
            function Get-Label {
                [CmdletBinding()] param()
                [pscustomobject]@{ Name = 'Public' }
                [pscustomobject]@{ Name = 'Confidential' }
            }
            function Get-LabelPolicy {
                [CmdletBinding()] param()
                $script:LabelPolicyCalls++
            }
            function Get-AutoSensitivityLabelPolicy { [CmdletBinding()] param() }
            function Get-DlpCompliancePolicy {
                [CmdletBinding()] param()
                [pscustomobject]@{ Name = 'DLP baseline'; Mode = 'Enable'; Workload = 'Exchange' }
            }
            function Get-RetentionCompliancePolicy {
                [CmdletBinding()] param()
                [pscustomobject]@{ Name = 'Finance hold'; Enabled = $true; Mode = 'Enforce'; Workload = 'SharePoint' }
            }

            $governance = Get-DataGovernanceReadiness
            Assert-Equal $script:LabelPolicyCalls 2 'Get-LabelPolicy is retried once when labels exist and the first read is empty.'
            Assert-Equal @($script:Report.Errors).Count 0 'An empty policy read is a diagnostic, not a collector error.'
            Assert-Equal $governance.SensitivityLabelsTotal 2 'Labels returned by Get-Label are kept.'
            Assert-Equal $governance.SensitivityLabelsPublished $null 'Published label count is not collected when policies stay empty.'
            Assert-Equal $governance.LabelPoliciesTotal $null 'Label policy total is not collected instead of zero.'
            Assert-Equal $governance.LabelPoliciesPublished $null 'Label policies published is not collected instead of zero.'
            Assert-True ($governance.LabelCollectionReason -match 'after one retry') 'The diagnostic note says the policy read was retried.'
            Assert-True ($governance.Status -eq 'partial') 'A withheld policy count makes the section partial.'
            $policyDiagnostics = @(Get-GraphReportDiagnosticsForUri -UriPattern 'Get-LabelPolicy')
            Assert-True ((@($policyDiagnostics | ForEach-Object { $_.Message }) -join ' ') -match 'after one retry') 'The empty retry is recorded in diagnostics.'
            Assert-Equal (Get-ExportExitCode -FailedCollectorCount @($script:Report.Errors).Count -ProcessedSectionCount 8 -FailOnCollectorError $false) 0 'The default exit code stays 0 when the empty read is only a diagnostic.'

            $script:Report = [ordered]@{
                Metadata = [ordered]@{ TotalUsers = 10 }
                Licensing = [ordered]@{ PrereqReadyDenominator = 10; PrereqReadyPercent = 'not collected' }
                Identity = [ordered]@{}
                ConditionalAccess = [ordered]@{}
                DataGovernance = $governance
                SharePointOneDrive = [ordered]@{}
                AdoptionSignals = [ordered]@{ ActiveUsers30Days = 'not collected' }
            }
            $flags = Get-ReadinessFlags
            Assert-Equal $flags.SensitivityLabelsPublished 'not collected' 'Scoring does not treat the missing policy count as zero published labels.'
            $script:Report.ReadinessFlags = $flags
            $script:Report.IdentityAccessAdvanced = [ordered]@{ Status = 'reported'; PhishingResistantMfaCoverage = 'not collected' }
            $script:Report.DataProtectionAdvanced = Get-DataProtectionAdvanced
            $script:Report.SharePointExposureAdvanced = [ordered]@{ Status = 'reported' }
            $script:Report.AppGovernanceAdvanced = [ordered]@{ Status = 'reported' }
            $script:Report.EndpointAppAdvanced = [ordered]@{ Status = 'reported' }
            $script:Report.AdoptionAdvanced = [ordered]@{ Status = 'not collected'; TrendSignal = 'not collected' }
            $script:ScoringModel = $null
            $evidence = Get-ReadinessEvidence
            Assert-Equal $evidence.DomainScores.DataGovernance 100 'DLP and retention still score when label publication was not collected.'
            Assert-Equal $evidence.DataCompleteness.MeasuredMetricWeight 66 'The label-publication weight is left out of completeness.'
            Assert-True ($evidence.DomainScores.DataGovernance -ne 66) 'A missing policy list is not scored as labels unpublished.'

            $script:Report = [ordered]@{ Errors = [System.Collections.Generic.List[object]]::new() }
            $script:GraphReportDiagnostics = [System.Collections.Generic.List[object]]::new()
            $script:LabelPolicyCalls = 0
            function Get-LabelPolicy {
                [CmdletBinding()] param()
                $script:LabelPolicyCalls++
                throw 'Label policy service timed out.'
            }
            $failed = Get-DataGovernanceReadiness
            Assert-Equal $script:LabelPolicyCalls 1 'A failing Get-LabelPolicy is not retried as an empty result.'
            Assert-Equal @($script:Report.Errors).Count 1 'The label policy failure is a collector error.'
            Assert-Equal $script:Report.Errors[0].Section 'DataGovernance' 'The error is recorded on Data governance.'
            Assert-True ($script:Report.Errors[0].Message -match 'Label policy service timed out') 'The cmdlet error is kept.'
            Assert-Equal $failed.LabelPoliciesTotal $null 'A failed policy read is not collected.'
            Assert-Equal $failed.Status 'partial' 'A failed Get-LabelPolicy with labels present is partial.'
            Assert-True ($failed.Reason -match 'Get-LabelPolicy failed; label policy counts are not collected\.') 'The failed policy read explains why the section is partial.'
            Assert-Equal (Get-ExportExitCode -FailedCollectorCount @($script:Report.Errors).Count -ProcessedSectionCount 8 -FailOnCollectorError $true) 3 'FailOnCollectorError exits 3 when the label policy error is counted.'
            Assert-Equal (Get-ExportExitCode -FailedCollectorCount @($script:Report.Errors).Count -ProcessedSectionCount 8 -FailOnCollectorError $false) 0 'The default exit code stays 0 for the same error.'

            $script:Report = [ordered]@{ Errors = [System.Collections.Generic.List[object]]::new() }
            $script:GraphReportDiagnostics = [System.Collections.Generic.List[object]]::new()
            $script:LabelPolicyCalls = 0
            function Get-LabelPolicy {
                [CmdletBinding()] param()
                $script:LabelPolicyCalls++
                if ($script:LabelPolicyCalls -eq 1) { return }
                throw 'Label policy retry failed.'
            }
            $retryFailed = Get-DataGovernanceReadiness
            Assert-Equal $script:LabelPolicyCalls 2 'An empty first policy read still retries once.'
            Assert-Equal $retryFailed.Status 'partial' 'A failed policy retry with labels present is partial.'
            Assert-True ($retryFailed.Reason -match 'Get-LabelPolicy failed; label policy counts are not collected\.') 'The failed retry explains why the section is partial.'
            Assert-Equal @($script:Report.Errors).Count 1 'The failed retry is one collector error.'
        }

        It 'records DLP and retention cmdlet failures as collector errors' {
            $script:Report = [ordered]@{ Errors = [System.Collections.Generic.List[object]]::new() }
            $script:GraphReportDiagnostics = [System.Collections.Generic.List[object]]::new()
            $script:GraphLastHttpStatus = 200
            function Get-ConnectionInformation {
                [CmdletBinding()] param()
                [pscustomobject]@{ Name = 'IPPSSession' }
            }
            function Get-GraphRestAll {
                param([string]$Uri, [hashtable]$Query = @{}, [switch]$BypassCache)
                $script:GraphLastHttpStatus = 200
                return @([pscustomobject]@{ name = 'Public'; isPublished = $true })
            }
            function Get-DlpCompliancePolicy {
                [CmdletBinding()] param()
                throw 'DLP query failed.'
            }
            function Get-RetentionCompliancePolicy {
                [CmdletBinding()] param()
                throw 'Retention query failed.'
            }

            $governance = Get-DataGovernanceReadiness
            Assert-Equal $governance.DlpPolicies 'not collected' 'A failed DLP read stays not collected.'
            Assert-Equal $governance.RetentionPolicies 'not collected' 'A failed retention read stays not collected.'
            Assert-Equal @($script:Report.Errors).Count 2 'DLP and retention failures are collector errors.'
            $dlpError = @($script:Report.Errors | Where-Object { $_.Detail -eq 'Get-DlpCompliancePolicy' }) | Select-Object -First 1
            $retentionError = @($script:Report.Errors | Where-Object { $_.Detail -eq 'Get-RetentionCompliancePolicy' }) | Select-Object -First 1
            Assert-Equal $dlpError.Section 'DataGovernance' 'The DLP error is recorded on Data governance.'
            Assert-Equal $dlpError.Message 'DLP query failed.' 'The DLP exception message is kept.'
            Assert-Equal $retentionError.Section 'DataGovernance' 'The retention error is recorded on Data governance.'
            Assert-Equal $retentionError.Message 'Retention query failed.' 'The retention exception message is kept.'
            Assert-Equal (Get-ExportExitCode -FailedCollectorCount @($script:Report.Errors).Count -ProcessedSectionCount 8 -FailOnCollectorError $true) 3 'FailOnCollectorError exits 3 when DLP and retention errors are counted.'
            Assert-Equal (Get-ExportExitCode -FailedCollectorCount @($script:Report.Errors).Count -ProcessedSectionCount 8 -FailOnCollectorError $false) 0 'The default exit code stays 0 for the same DLP and retention errors.'
        }

        It 'marks Data governance partial when only one compliance policy read fails' {
            $script:Report = [ordered]@{ Errors = [System.Collections.Generic.List[object]]::new() }
            $script:GraphReportDiagnostics = [System.Collections.Generic.List[object]]::new()
            function Get-ConnectionInformation {
                [CmdletBinding()] param()
                [pscustomobject]@{ Name = 'IPPSSession' }
            }
            function Get-GraphRestAll {
                param([string]$Uri, [hashtable]$Query = @{}, [switch]$BypassCache)
                $script:GraphLastHttpStatus = 200
                return @([pscustomobject]@{ name = 'Public'; isPublished = $true })
            }
            function Get-DlpCompliancePolicy {
                [CmdletBinding()] param()
                throw 'DLP query failed.'
            }
            function Get-RetentionCompliancePolicy {
                [CmdletBinding()] param()
                [pscustomobject]@{ Name = 'Finance hold'; Enabled = $true; Mode = 'Enforce'; Workload = 'SharePoint' }
            }

            $dlpFailed = Get-DataGovernanceReadiness
            Assert-Equal $dlpFailed.Status 'partial' 'A failed DLP read with retention collected is partial.'
            Assert-True ($dlpFailed.Reason -match 'DLP policy read failed: DLP query failed\.') 'The reason names the failed DLP read.'
            Assert-Equal $dlpFailed.DlpPolicies 'not collected' 'The failed DLP list stays not collected.'
            Assert-True ($dlpFailed.RetentionPolicies -is [System.Array]) 'The successful retention list is kept.'
            Assert-Equal @($script:Report.Errors).Count 1 'The DLP failure is still a collector error.'
            Assert-Equal $script:Report.Errors[0].Message 'DLP query failed.' 'The D1 DLP error message is kept.'

            $script:Report = [ordered]@{ Errors = [System.Collections.Generic.List[object]]::new() }
            function Get-DlpCompliancePolicy {
                [CmdletBinding()] param()
                [pscustomobject]@{ Name = 'DLP baseline'; Mode = 'Enable'; Workload = 'Exchange' }
            }
            function Get-RetentionCompliancePolicy {
                [CmdletBinding()] param()
                throw 'Retention query failed.'
            }
            $retentionFailed = Get-DataGovernanceReadiness
            Assert-Equal $retentionFailed.Status 'partial' 'A failed retention read with DLP collected is partial.'
            Assert-True ($retentionFailed.Reason -match 'Retention policy read failed: Retention query failed\.') 'The reason names the failed retention read.'
            Assert-Equal @($script:Report.Errors).Count 1 'The retention failure is still a collector error.'
            Assert-Equal (Get-ExportExitCode -FailedCollectorCount @($script:Report.Errors).Count -ProcessedSectionCount 8 -FailOnCollectorError $true) 3 'FailOnCollectorError still exits 3 for the retention error.'
            Assert-Equal (Get-ExportExitCode -FailedCollectorCount @($script:Report.Errors).Count -ProcessedSectionCount 8 -FailOnCollectorError $false) 0 'The default exit code stays 0.'
        }

        It 'gives full anyone-link credit when tenant settings block them and sites were not collected' {
            $script:Report = [ordered]@{
                SharePointOneDrive = [ordered]@{
                    Status = 'partial'
                    ExternalSharingLevel = 'ExternalUserSharingOnly'
                    SiteInventoryStatus = 'not collected'
                    SiteInventoryReason = 'Site inventory was not collected.'
                    TotalSiteCount = 'not collected'
                    Sites = @()
                }
            }
            $blocked = Get-SharePointExposureAdvanced
            Assert-Equal $blocked.AnyoneLinkSites 'not collected' 'Blocked tenant settings do not invent a zero site count.'
            Assert-Equal $blocked.AnyoneLinksPossible $false 'ExternalUserSharingOnly means anyone links are not possible.'
            Assert-Equal $blocked.AnyoneLinkMeasurement 'tenant-settings' 'The blocked signal is still tenant settings.'

            $script:Report = [ordered]@{
                ReadinessFlags = [ordered]@{
                    MfaRegisteredPercent = 'not collected'
                    CaPoliciesEnabled = 'not collected'
                    SensitivityLabelsPublished = 'not collected'
                    OneDrivePersonalSiteSharePercent = 'not collected'
                    RestrictedSearchEnabled = 'not collected'
                    ActiveUsersInReportPeriodPercent = 'not collected'
                    ExternalSharingLevel = 'ExternalUserSharingOnly'
                }
                Identity = [ordered]@{ Status = 'reported' }
                ConditionalAccess = [ordered]@{ Status = 'reported' }
                DataGovernance = [ordered]@{ Status = 'reported' }
                SharePointOneDrive = [ordered]@{ Status = 'reported' }
                AdoptionSignals = [ordered]@{ Status = 'reported' }
                IdentityAccessAdvanced = [ordered]@{ Status = 'reported' }
                DataProtectionAdvanced = [ordered]@{ Status = 'reported' }
                SharePointExposureAdvanced = $blocked
                AppGovernanceAdvanced = [ordered]@{ Status = 'reported' }
                EndpointAppAdvanced = [ordered]@{ Status = 'reported' }
                AdoptionAdvanced = [ordered]@{ Status = 'reported' }
            }
            $script:ScoringModel = $null
            $blockedEvidence = Get-ReadinessEvidence
            Assert-Equal $blockedEvidence.DomainScores.ContentExposure 100 'Tenant settings that block anyone links earn the full factor.'
            Assert-Equal $blockedEvidence.DataCompleteness.AnyoneLinkConfidence 'reduced' 'Tenant-only evidence keeps reduced confidence.'
            Assert-Equal $blockedEvidence.DataCompleteness.AnyoneLinkMeasurement 'tenant-settings' 'Completeness records the tenant-settings measurement.'
            Assert-Equal $blockedEvidence.DataCompleteness.CompletenessPercent 3.3 'The anyone-link weight is scaled by the tenant-only factor.'
            Assert-Equal $blockedEvidence.RiskSignals.AnyoneLinkSites 'not collected' 'Risk signals do not turn the missing site count into zero.'

            $script:Report.SharePointExposureAdvanced = [ordered]@{
                AnyoneLinkSites = 'not collected'
                AnyoneLinksPossible = $true
                AnyoneLinkMeasurement = 'tenant-settings'
            }
            $script:Report.ReadinessFlags.ExternalSharingLevel = 'ExternalUserAndGuestSharing'
            $possibleEvidence = Get-ReadinessEvidence
            Assert-Equal $possibleEvidence.DomainScores.ContentExposure 0 'Tenant settings that allow anyone links still score as present.'
        }

        It 'scores phishing-resistant MFA for enabled members and reports PIM without failing the run' {
            $directoryUsers = [System.Collections.Generic.List[object]]::new()
            foreach ($index in 1..10) {
                $directoryUsers.Add([pscustomobject]@{ id = "member-$index"; userType = 'Member'; accountEnabled = $true }) | Out-Null
            }
            $directoryUsers.Add([pscustomobject]@{ id = 'guest-user'; userType = 'Guest'; accountEnabled = $true }) | Out-Null
            $directoryUsers.Add([pscustomobject]@{ id = 'disabled-member'; userType = 'Member'; accountEnabled = $false }) | Out-Null
            $script:GraphUsersCache = @{ base = $directoryUsers.ToArray() }
            $registration = @(
                [pscustomobject]@{ id = 'member-1'; userType = 'member'; methodsRegistered = @('fido2') }
                [pscustomobject]@{ id = 'member-2'; userType = 'member'; methodsRegistered = @('x509Certificate') }
                [pscustomobject]@{ id = 'member-3'; userType = 'member'; methodsRegistered = @('microsoftAuthenticatorPush') }
                [pscustomobject]@{ id = 'guest-user'; userType = 'guest'; methodsRegistered = @('windowsHelloForBusiness') }
                [pscustomobject]@{ id = 'disabled-member'; userType = 'member'; methodsRegistered = @('passKeyDeviceBoundAuthenticator') }
            )

            $coverage = Get-PhishingResistantMfaCoverageResult -RegistrationDetails $registration
            Assert-Equal $coverage.UserCount 2 'Only enabled members with a phishing-resistant method are counted.'
            Assert-Equal $coverage.Denominator 10 'The denominator is the enabled member count.'
            Assert-Equal $coverage.Coverage 20 '2 of 10 enabled members is 20 percent.'

            $script:Report = [ordered]@{ Errors = [System.Collections.Generic.List[object]]::new() }
            $script:GraphReportDiagnostics = [System.Collections.Generic.List[object]]::new()
            $denied = Get-PhishingResistantMfaCoverageResult -FailureReason 'HTTP 403. UserAuthenticationMethod.Read.All or AuditLog.Read.All is required.'
            Assert-Equal $denied.Coverage 'not collected' 'A 403 leaves phishing-resistant coverage not collected.'
            Assert-True ($denied.Reason -match 'HTTP 403') 'The permission failure is the reason.'
            Assert-Equal @($script:Report.Errors).Count 0 'The diagnostic does not become a collector error.'
            Assert-Equal (Get-ExportExitCode -FailedCollectorCount @($script:Report.Errors).Count -ProcessedSectionCount 8 -FailOnCollectorError $true) 0 'FailOnCollectorError does not fail the run for this diagnostic.'

            function Get-GraphRestAll {
                param([string]$Uri, [hashtable]$Query = @{}, [switch]$BypassCache)
                if ($Uri -match 'roleEligibilitySchedules$') {
                    $script:GraphLastHttpStatus = 200
                    $script:GraphLastErrorMessage = $null
                    return @(
                        [pscustomobject]@{ roleDefinitionId = '62e90394-69f5-4237-9190-012177145e10'; principalId = 'admin' }
                        [pscustomobject]@{ roleDefinitionId = '11111111-1111-1111-1111-111111111111'; principalId = 'other' }
                    )
                }
                if ($Uri -match 'roleAssignmentSchedules') {
                    $script:GraphLastHttpStatus = 403
                    $script:GraphLastErrorMessage = 'Authorization_RequestDenied'
                    return @()
                }
                if ($Uri -match 'roleAssignments') {
                    $script:GraphLastHttpStatus = 200
                    $script:GraphLastErrorMessage = $null
                    return @([pscustomobject]@{ roleDefinitionId = '194ae4cb-b126-40b2-bd5b-6091b380977d' })
                }
                $script:GraphLastHttpStatus = 200
                return @()
            }
            $pim = Get-PimDirectorySignal
            Assert-Equal $pim.Signal 'in use' 'An eligible privileged assignment means PIM is in use.'
            Assert-Equal $pim.EligibleAssignmentCount 2 'Both eligibility schedules are counted.'
            Assert-Equal $pim.PrivilegedRoleEligibleCount 1 'Only the Global Administrator schedule is a privileged role.'
            Assert-Equal $pim.PermanentPrivilegedAssignmentCount 1 'The Security Administrator role assignment is permanent and privileged.'
            Assert-Equal $pim.ActiveAssignmentScheduleCount 'not collected' 'A denied assignment-schedule read stays not collected.'
            Assert-Equal @($script:Report.Errors).Count 0 'A denied optional PIM read does not fail the run.'

            function Get-GraphRestAll {
                param([string]$Uri, [hashtable]$Query = @{}, [switch]$BypassCache)
                $script:GraphLastHttpStatus = 400
                $script:GraphLastErrorMessage = 'AadPremiumLicenseRequired'
                $script:GraphLastErrorCode = 'AadPremiumLicenseRequired'
                return @()
            }
            $unlicensed = Get-PimDirectorySignal
            Assert-Equal $unlicensed.Signal 'not collected' 'A tenant without Entra ID P2 does not report a false PIM signal.'
            Assert-True ($unlicensed.Reason -match 'AadPremiumLicenseRequired') 'The license error is the reason.'
            Assert-Equal @($script:Report.Errors).Count 0 'The PIM license error stays a diagnostic.'

            $script:Report = [ordered]@{
                ReadinessFlags = [ordered]@{
                    MfaRegisteredPercent = 'not collected'
                    MfaRegisteredEnabledMemberPercent = 'not collected'
                    CaPoliciesEnabled = 'not collected'
                    SensitivityLabelsPublished = 'not collected'
                    OneDrivePersonalSiteSharePercent = 'not collected'
                    RestrictedSearchEnabled = 'not collected'
                    ActiveUsersInReportPeriodPercent = 'not collected'
                }
                Identity = [ordered]@{ Status = 'reported' }
                ConditionalAccess = [ordered]@{ Status = 'reported' }
                DataGovernance = [ordered]@{ Status = 'reported' }
                SharePointOneDrive = [ordered]@{ Status = 'reported' }
                AdoptionSignals = [ordered]@{ Status = 'reported' }
                IdentityAccessAdvanced = [ordered]@{
                    Status = 'reported'
                    PhishingResistantMfaCoverage = $coverage.Coverage
                    PimSignal = $pim.Signal
                }
                DataProtectionAdvanced = [ordered]@{ Status = 'reported' }
                SharePointExposureAdvanced = [ordered]@{ Status = 'reported' }
                AppGovernanceAdvanced = [ordered]@{ Status = 'reported' }
                EndpointAppAdvanced = [ordered]@{ Status = 'reported' }
                AdoptionAdvanced = [ordered]@{ Status = 'reported' }
            }
            $script:ScoringModel = $null
            $evidence = Get-ReadinessEvidence
            Assert-Equal $evidence.IdentityScoreInputs.AdvancedMfaCoveragePercent 20 'The advanced MFA score input is the phishing-resistant percent.'
            Assert-Equal $evidence.DomainScores.IdentityAccess 100 '20 percent meets the advanced MFA threshold and earns that weight.'
            Assert-Equal $evidence.DataCompleteness.MeasuredMetricWeight 30 'The existing advanced MFA weight is measured.'
            Assert-Equal $evidence.DataCompleteness.PossibleMetricWeight 500 'PIM does not add a new scoring weight.'
        }
    }
}

Write-Output 'm365-ai-ready-export collector tests passed.'
