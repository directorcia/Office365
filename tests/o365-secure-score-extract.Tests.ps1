#requires -Version 5.1
<#
.SYNOPSIS
    Offline Pester tests for o365-secure-score-extract.ps1.
.EXAMPLE
    Invoke-Pester -Script .\tests\o365-secure-score-extract.Tests.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath 'Analysis\Secure Score\o365-secure-score-extract.ps1'

function Assert-ThrowsWithMessage {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Action,

        [Parameter(Mandatory)]
        [string]$ExpectedMessage
    )

    $caught = $null
    try {
        & $Action
    }
    catch {
        $caught = $_
    }

    if ($null -eq $caught -or $caught.Exception.Message -notlike "*$ExpectedMessage*") {
        throw "Expected an exception containing '$ExpectedMessage'; received '$caught'."
    }
}

Describe 'o365-secure-score-extract' {
    BeforeEach {
        $script:tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ('ss-extract-tests-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $script:tempDir -Force | Out-Null

        Mock Import-Module {
            [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication' }
        }
        Mock Connect-MgGraph {}
        Mock Disconnect-MgGraph {}
        Mock Get-MgContext {
            [pscustomobject]@{
                Account  = 'secadmin@contoso.com'
                TenantId = '11111111-2222-3333-4444-555555555555'
                Scopes   = @(
                    'SecurityEvents.Read.All',
                    'Policy.Read.All',
                    'Reports.Read.All',
                    'Directory.Read.All'
                )
            }
        }
        Mock Invoke-MgGraphRequest {
            switch -Wildcard ($Uri) {
                '*/v1.0/organization*' {
                    return @{
                        value = @(
                            @{
                                id              = '11111111-2222-3333-4444-555555555555'
                                displayName     = 'Contoso'
                                verifiedDomains = @(
                                    @{ name = 'contoso.onmicrosoft.com'; isDefault = $true }
                                    @{ name = 'contoso.com'; isDefault = $false }
                                )
                            }
                        )
                    }
                }
                '*/beta/security/secureScores*' {
                    return @{
                        value = @(
                            @{
                                createdDateTime = '2026-03-02T00:00:00Z'
                                currentScore    = 82.5
                                maxScore        = 100
                                controlScores   = @(
                                    @{ controlName = 'AdminMFAV2'; score = 4 }
                                    @{ controlName = 'BlockLegacyAuth'; score = 8 }
                                )
                            }
                            @{ createdDateTime = '2026-03-01T00:00:00Z'; currentScore = 78.0; maxScore = 100; controlScores = @() }
                        )
                    }
                }
                '*/beta/security/secureScoreControlProfiles*' {
                    return @{
                        value = @(
                            @{
                                id                  = 'AdminMFAV2'
                                title               = 'Require MFA for administrative roles'
                                rank                = 1
                                tier                = 'Core'
                                maxScore            = 10
                                controlStateUpdates = @(@{ state = 'default' })
                            }
                            @{
                                id                  = 'BlockLegacyAuth'
                                title               = 'Block legacy authentication'
                                rank                = 2
                                tier                = 'Core'
                                maxScore            = 8
                                controlStateUpdates = @(@{ state = 'completed' })
                            }
                            @{
                                id                  = 'InfoControl'
                                title               = 'Informational item'
                                rank                = 3
                                tier                = 'informational'
                                maxScore            = 0
                                controlStateUpdates = @()
                            }
                        )
                    }
                }
                '*/v1.0/identity/conditionalAccess/policies*' {
                    return @{
                        value = @(
                            @{
                                id              = 'ca-1'
                                displayName     = 'Require MFA for Admins'
                                state           = 'enabled'
                                createdDateTime = '2026-01-10T00:00:00Z'
                                conditions      = @{
                                    users = @{
                                        includeRoles = @('62e90394-69f5-4237-9190-012177145e10')
                                    }
                                }
                                grantControls   = @{ builtInControls = @('mfa') }
                            }
                        )
                    }
                }
                '*/v1.0/policies/identitySecurityDefaultsEnforcementPolicy' {
                    return @{
                        id        = '00000000-0000-0000-0000-000000000005'
                        isEnabled = $false
                    }
                }
                '*/beta/reports/authenticationMethods/userRegistrationActivity*' {
                    return @{
                        value = @(
                            @{ userPrincipalName = 'alice@contoso.com'; isMfaRegistered = $true }
                        )
                    }
                }
                default {
                    throw "Unexpected Graph URI in test mock: $Uri"
                }
            }
        }
    }

    AfterEach {
        if ($script:tempDir -and (Test-Path -LiteralPath $script:tempDir)) {
            Remove-Item -LiteralPath $script:tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        foreach ($envName in @('SS_EXTRACT_SCORE_FIXTURE', 'SS_EXTRACT_PROFILE_FIXTURE', 'SS_EXTRACT_CONTEXT_CALLS')) {
            if (Test-Path -LiteralPath "Env:$envName") {
                Remove-Item -LiteralPath "Env:$envName"
            }
        }
    }

    It 'extracts full and compact JSON files (supporting spaces in paths) and emits a summary object' {
        $outPath = Join-Path -Path $script:tempDir -ChildPath 'sub folder\contoso export.json'
        $summary = & $scriptPath -TenantDomain 'contoso.com' -DataFile $outPath -Compact

        $summary | Should Not BeNullOrEmpty
        $summary.Tenant | Should Be 'contoso.com'
        $summary.LatestScore | Should Be 82.5
        $summary.OpenControlCount | Should Be 1
        $summary.ScoredControlCount | Should Be 2
        $summary.FullyAchievedCount | Should Be 1
        $summary.AddressedWithPointsRemaining | Should Be 0
        $summary.NotApplicableCount | Should Be 1
        $summary.TotalPointsAvailable | Should Be 6
        $summary.ConditionalAccessCount | Should Be 1
        $summary.SecurityDefaultsEnabled | Should Be $false
        $summary.MfaRecordCount | Should Be 1
        $summary.StagesSucceeded | Should Be 5
        $summary.StagesDegraded | Should Be 0
        $LASTEXITCODE | Should Be 0

        (Test-Path -LiteralPath $outPath) | Should Be $true
        (Test-Path -LiteralPath $summary.CompactDataFile) | Should Be $true

        $partialPath = Join-Path -Path $script:tempDir -ChildPath 'sub folder\contoso export_partial.json'
        (Test-Path -LiteralPath $partialPath) | Should Be $false

        $parsedFull = Get-Content -LiteralPath $outPath -Raw | ConvertFrom-Json
        $parsedFull.SecureScoreControls.Open.Count | Should Be 1
        $parsedFull.SecureScoreControls.Open[0].id | Should Be 'AdminMFAV2'

        $parsedCompact = Get-Content -LiteralPath $summary.CompactDataFile -Raw | ConvertFrom-Json
        $parsedCompact.SecureScore.HistorySummary.Count | Should Be 2
        $parsedCompact.SecureScoreControls.OpenCount | Should Be 1
        $parsedCompact.SecureScoreControls.OpenControlCount | Should Be 1
        $parsedCompact.SecureScoreControls.ScoredControlCount | Should Be 2
        $parsedCompact.MfaRegistrationSummaryDescription | Should Match 'not a point-in-time'

        # Reused active session because contoso.com is in verifiedDomains and all 4 scopes match
        Assert-MockCalled Connect-MgGraph -Times 0 -Scope It
    }

    It 'reconnects when active session is missing a required scope' {
        $state = @{ Reconnected = $false }
        Mock Get-MgContext {
            [pscustomobject]@{
                Account  = 'secadmin@contoso.com'
                TenantId = '11111111-2222-3333-4444-555555555555'
                Scopes   = @('SecurityEvents.Read.All')
            }
        }
        Mock Connect-MgGraph { $null = $state }

        $outPath = Join-Path -Path $script:tempDir -ChildPath 'scope-reconnect.json'
        $null = & $scriptPath -TenantDomain 'contoso.onmicrosoft.com' -DataFile $outPath -WarningAction SilentlyContinue

        Assert-MockCalled Disconnect-MgGraph -Times 1 -Scope It
        Assert-MockCalled Connect-MgGraph -Times 1 -Scope It
    }

    It 'honors WhatIf without connecting to Graph or creating files' {
        $outPath = Join-Path -Path $script:tempDir -ChildPath 'whatif-test.json'
        $result = @(& $scriptPath -TenantDomain 'contoso.onmicrosoft.com' -DataFile $outPath -Compact -WhatIf)

        $result.Count | Should Be 0
        (Test-Path -LiteralPath $outPath) | Should Be $false
        Assert-MockCalled Invoke-MgGraphRequest -Times 0 -Scope It
        Assert-MockCalled Connect-MgGraph -Times 0 -Scope It
    }

    It 'supports pipeline input across multiple tenants without overwriting DataFile' {
        $baseOutPath = Join-Path -Path $script:tempDir -ChildPath 'batch.json'
        $results = @('contoso.onmicrosoft.com', 'contoso.com' | & $scriptPath -DataFile $baseOutPath)

        $results.Count | Should Be 2
        $results[0].FullDataFile | Should Not Be $results[1].FullDataFile
        (Test-Path -LiteralPath $results[0].FullDataFile) | Should Be $true
        (Test-Path -LiteralPath $results[1].FullDataFile) | Should Be $true
    }

    It 'retries transient HTTP 429 responses and follows @odata.nextLink pagination' {
        $attemptCounter = [System.Collections.Generic.List[int]]::new()
        Mock Start-Sleep {}
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*userRegistrationActivity*page2*') {
                return @{
                    value = @(@{ userPrincipalName = 'bob@contoso.com'; isMfaRegistered = $false })
                }
            }
            if ($Uri -like '*userRegistrationActivity*') {
                $attemptCounter.Add(1)
                if ($attemptCounter.Count -eq 1) {
                    throw 'Response status code does not indicate success: 429 (TooManyRequests).'
                }
                return @{
                    value             = @(@{ userPrincipalName = 'alice@contoso.com'; isMfaRegistered = $true })
                    '@odata.nextLink' = 'https://graph.microsoft.com/beta/reports/authenticationMethods/userRegistrationActivity?page2=1'
                }
            }
            if ($Uri -like '*/v1.0/organization*') {
                return @{
                    value = @(@{ id = '11111111-2222-3333-4444-555555555555'; verifiedDomains = @(@{ name = 'contoso.com'; isDefault = $true }) })
                }
            }
            return @{ value = @() }
        }

        $outPath = Join-Path -Path $script:tempDir -ChildPath 'retry-paged.json'
        $summary = & $scriptPath -TenantDomain 'contoso.com' -DataFile $outPath -InitialBackoffMs 100 -WarningAction SilentlyContinue

        $summary.MfaRecordCount | Should Be 2
        Assert-MockCalled Start-Sleep -Times 1 -Scope It
    }

    It 'preserves the _partial recovery file and sets LASTEXITCODE to 1 when a core stage fails' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*/v1.0/organization*') {
                return @{
                    value = @(@{ id = '11111111-2222-3333-4444-555555555555'; verifiedDomains = @(@{ name = 'contoso.com'; isDefault = $true }) })
                }
            }
            if ($Uri -like '*/beta/security/secureScores*') {
                return @{ value = @(@{ createdDateTime = '2026-03-01T00:00:00Z'; currentScore = 50; maxScore = 100 }) }
            }
            if ($Uri -like '*/beta/security/secureScoreControlProfiles*') {
                throw 'Simulated control profile failure'
            }
            return @{ value = @() }
        }

        $outPath = Join-Path -Path $script:tempDir -ChildPath 'fail-recovery.json'
        $partialPath = Join-Path -Path $script:tempDir -ChildPath 'fail-recovery_partial.json'

        Assert-ThrowsWithMessage -Action {
            & $scriptPath -TenantDomain 'contoso.com' -DataFile $outPath -WarningAction SilentlyContinue
        } -ExpectedMessage 'Simulated control profile failure'

        $LASTEXITCODE | Should Be 1
        (Test-Path -LiteralPath $outPath) | Should Be $false
        (Test-Path -LiteralPath $partialPath) | Should Be $true
    }

    It 'counts scored, open, fully achieved, addressed, and not-applicable controls from the latest secureScore' {
        $fixtureDir = Join-Path -Path $PSScriptRoot -ChildPath 'fixtures'
        $fixtureDir = Join-Path -Path $fixtureDir -ChildPath 'secure-score-open-controls'
        # Pester 4 mocks run inside the script under test, so pass fixture JSON via the environment.
        $env:SS_EXTRACT_SCORE_FIXTURE = Get-Content -LiteralPath (Join-Path -Path $fixtureDir -ChildPath 'secureScores.json') -Raw
        $env:SS_EXTRACT_PROFILE_FIXTURE = Get-Content -LiteralPath (Join-Path -Path $fixtureDir -ChildPath 'controlProfiles.json') -Raw

        Mock Invoke-MgGraphRequest {
            switch -Wildcard ($Uri) {
                '*/v1.0/organization*' {
                    return @{
                        value = @(
                            @{
                                id              = '11111111-2222-3333-4444-555555555555'
                                verifiedDomains = @(@{ name = 'contoso.com'; isDefault = $true })
                            }
                        )
                    }
                }
                '*/beta/security/secureScoreControlProfiles*' { return ($env:SS_EXTRACT_PROFILE_FIXTURE | ConvertFrom-Json) }
                '*/beta/security/secureScores*' { return ($env:SS_EXTRACT_SCORE_FIXTURE | ConvertFrom-Json) }
                '*/v1.0/identity/conditionalAccess/policies*' { return @{ value = @() } }
                '*/v1.0/policies/identitySecurityDefaultsEnforcementPolicy' { return @{ isEnabled = $false } }
                '*/beta/reports/authenticationMethods/userRegistrationActivity*' { return @{ value = @() } }
                default { throw "Unexpected Graph URI in fixture test: $Uri" }
            }
        }

        $outPath = Join-Path -Path $script:tempDir -ChildPath 'fixture-counts.json'
        $captured = @(
            & $scriptPath -TenantDomain 'contoso.com' -DataFile $outPath -Compact -InformationAction Continue 6>&1
        )
        $summary = $captured | Where-Object { $null -ne $_.PSObject.Properties['ScoredControlCount'] } | Select-Object -First 1
        $infoText = @(
            $captured |
                Where-Object { $_ -is [System.Management.Automation.InformationRecord] } |
                ForEach-Object { [string]$_.MessageData }
        ) -join "`n"

        $summary.ScoredControlCount | Should Be 3
        $summary.OpenControlCount | Should Be 1
        $summary.FullyAchievedCount | Should Be 1
        $summary.AddressedWithPointsRemaining | Should Be 1
        $summary.NotApplicableCount | Should Be 2
        $summary.TotalPointsAvailable | Should Be 13
        $summary.LatestScore | Should Be 15

        $parsedFull = Get-Content -LiteralPath $outPath -Raw | ConvertFrom-Json
        $parsedFull.SecureScoreControls.ScoredControlCount | Should Be 3
        $parsedFull.SecureScoreControls.OpenControlCount | Should Be 1
        $parsedFull.SecureScoreControls.FullyAchievedCount | Should Be 1
        $parsedFull.SecureScoreControls.AddressedWithPointsRemaining | Should Be 1
        $parsedFull.SecureScoreControls.NotApplicableCount | Should Be 2
        $parsedFull.SecureScoreControls.TotalPointsAvailable | Should Be 13
        $parsedFull.SecureScoreControls.Open[0].id | Should Be 'ControlPartial'
        $parsedFull.SecureScoreControls.Open[0].pointsAvailable | Should Be 7
        $parsedFull.SecureScoreControls.Open[0].currentScore | Should Be 3
        $parsedFull.SecureScoreControls.Open[0].implementationCost | Should Be 'low'
        @($parsedFull.SecureScoreControls.All).Count | Should Be 5
        $parsedFull.MfaRegistrationSummaryDescription | Should Match 'registration activity events'
        $parsedFull.PSObject.Properties.Name -contains 'MfaRegistrationSummary' | Should Be $true

        $ignored = @($parsedFull.SecureScoreControls.All | Where-Object { $_.id -eq 'ControlIgnored' })
        $ignored.Count | Should Be 1
        $ignored[0].pointsAvailable | Should Be 6

        $parsedCompact = Get-Content -LiteralPath $summary.CompactDataFile -Raw | ConvertFrom-Json
        $parsedCompact.SecureScoreControls.TotalCount | Should Be 5
        $parsedCompact.SecureScoreControls.OpenCount | Should Be 1
        $parsedCompact.SecureScoreControls.ScoredControlCount | Should Be 3
        $parsedCompact.SecureScoreControls.NotApplicableCount | Should Be 2
        $parsedCompact.SecureScoreControls.TopOpen[0].id | Should Be 'ControlPartial'
        $parsedCompact.SecureScoreControls.TopOpen[0].implementationCost | Should Be 'Low'
        $parsedCompact.SecureScoreControls.TopOpen[0].userImpact | Should Be 'Moderate'
        $parsedCompact.SecureScoreControls.TopOpen[0].pointsAvailable | Should Be 7
        $parsedCompact.MfaRegistrationSummaryDescription | Should Match 'not a point-in-time'

        $infoText | Should Match 'MFA registration activity \(events\)'
        $infoText | Should Not Match 'MFA registration summary'
    }

    It 'orders top open controls by points available, then rank with zero and null last' {
        Mock Invoke-MgGraphRequest {
            switch -Wildcard ($Uri) {
                '*/v1.0/organization*' {
                    return @{
                        value = @(@{ id = '11111111-2222-3333-4444-555555555555'; verifiedDomains = @(@{ name = 'contoso.com'; isDefault = $true }) })
                    }
                }
                '*/beta/security/secureScores*' {
                    return @{
                        value = @(
                            @{
                                createdDateTime = '2026-11-22T00:00:00Z'
                                currentScore    = 11
                                maxScore        = 40
                                controlScores   = @(
                                    @{ controlName = 'LowPoints'; score = 8 }
                                    @{ controlName = 'HighPointsZeroRank'; score = 1 }
                                    @{ controlName = 'HighPointsRank3'; score = 1 }
                                    @{ controlName = 'HighPointsNullRank'; score = 1 }
                                )
                            }
                        )
                    }
                }
                '*/beta/security/secureScoreControlProfiles*' {
                    return @{
                        value = @(
                            @{ id = 'LowPoints'; title = 'Small gap'; rank = 1; tier = 'Core'; maxScore = 10; controlStateUpdates = @(@{ state = 'Default' }) }
                            @{ id = 'HighPointsZeroRank'; title = 'Unranked device gap'; rank = 0; tier = 'Core'; maxScore = 10; controlStateUpdates = @(@{ state = 'Default' }) }
                            @{ id = 'HighPointsRank3'; title = 'Ranked gap'; rank = 3; tier = 'Core'; maxScore = 10; controlStateUpdates = @(@{ state = 'Default' }) }
                            @{ id = 'HighPointsNullRank'; title = 'Missing rank'; rank = $null; tier = 'Core'; maxScore = 10; controlStateUpdates = @(@{ state = 'Default' }) }
                        )
                    }
                }
                '*/v1.0/identity/conditionalAccess/policies*' { return @{ value = @() } }
                '*/v1.0/policies/identitySecurityDefaultsEnforcementPolicy' { return @{ isEnabled = $true } }
                '*/beta/reports/authenticationMethods/userRegistrationActivity*' { return @{ value = @() } }
                default { throw "Unexpected Graph URI in ordering test: $Uri" }
            }
        }

        $outPath = Join-Path -Path $script:tempDir -ChildPath 'ordering.json'
        $summary = & $scriptPath -TenantDomain 'contoso.com' -DataFile $outPath -Compact

        $summary.OpenControlCount | Should Be 4
        $summary.TotalPointsAvailable | Should Be 29

        $parsed = Get-Content -LiteralPath $outPath -Raw | ConvertFrom-Json
        $ids = @($parsed.SecureScoreControls.TopOpen | ForEach-Object { $_.id })
        $ids[0] | Should Be 'HighPointsRank3'
        $ids[$ids.Count - 1] | Should Be 'LowPoints'
        (@($ids[1], $ids[2]) -contains 'HighPointsZeroRank') | Should Be $true
        (@($ids[1], $ids[2]) -contains 'HighPointsNullRank') | Should Be $true

        $compact = Get-Content -LiteralPath $summary.CompactDataFile -Raw | ConvertFrom-Json
        $compact.SecureScoreControls.TopOpen[0].id | Should Be 'HighPointsRank3'
        $compact.SecureScoreControls.TopOpen[0].pointsAvailable | Should Be 9
    }

    It 'labels a reused app-only session when Account is empty' {
        Mock Get-MgContext {
            [pscustomobject]@{
                Account  = '   '
                AppName  = 'Contoso Extractor'
                ClientId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
                TenantId = '11111111-2222-3333-4444-555555555555'
                AuthType = 'AppOnly'
                Scopes   = @()
            }
        }

        $outPath = Join-Path -Path $script:tempDir -ChildPath 'apponly-reuse.json'
        $captured = @(& $scriptPath -TenantDomain 'contoso.com' -DataFile $outPath -InformationAction Continue 6>&1)
        $summary = $captured | Where-Object { $null -ne $_.PSObject.Properties['CollectedBy'] } | Select-Object -First 1
        $infoText = @(
            $captured |
                Where-Object { $_ -is [System.Management.Automation.InformationRecord] } |
                ForEach-Object { [string]$_.MessageData }
        ) -join "`n"
        $expectedLabel = 'Contoso Extractor (aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee) [AppOnly]'

        $summary.CollectedBy | Should Be $expectedLabel
        $infoText | Should Match 'Reusing active Microsoft Graph session \(Contoso Extractor \(aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee\) \[AppOnly\], API health check'
        $infoText | Should Not Match '\(, API health check'
        Assert-MockCalled Connect-MgGraph -Times 0 -Scope It

        $parsed = Get-Content -LiteralPath $outPath -Raw | ConvertFrom-Json
        $parsed.CollectedBy | Should Be $expectedLabel
    }

    It 'labels a fresh app-only sign-in when Account is empty' {
        $env:SS_EXTRACT_CONTEXT_CALLS = '0'
        Mock Get-MgContext {
            $callCount = [int]$env:SS_EXTRACT_CONTEXT_CALLS + 1
            $env:SS_EXTRACT_CONTEXT_CALLS = [string]$callCount
            if ($callCount -eq 1) {
                return $null
            }
            [pscustomobject]@{
                Account  = $null
                AppName  = 'Contoso Extractor'
                ClientId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
                TenantId = '11111111-2222-3333-4444-555555555555'
                AuthType = 'AppOnly'
                Scopes   = @()
            }
        }

        $outPath = Join-Path -Path $script:tempDir -ChildPath 'apponly-connect.json'
        $captured = @(& $scriptPath -TenantDomain 'contoso.com' -DataFile $outPath -InformationAction Continue 6>&1)
        $summary = $captured | Where-Object { $null -ne $_.PSObject.Properties['CollectedBy'] } | Select-Object -First 1
        $infoText = @(
            $captured |
                Where-Object { $_ -is [System.Management.Automation.InformationRecord] } |
                ForEach-Object { [string]$_.MessageData }
        ) -join "`n"

        $summary.CollectedBy | Should Be 'Contoso Extractor (aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee) [AppOnly]'
        $infoText | Should Match "Connected as 'Contoso Extractor \(aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee\) \[AppOnly\]'"
        Assert-MockCalled Connect-MgGraph -Times 1 -Scope It
    }

    It 'maps compact Medium and medium cost and impact to Moderate and title-cases other values' {
        Mock Invoke-MgGraphRequest {
            switch -Wildcard ($Uri) {
                '*/v1.0/organization*' {
                    return @{
                        value = @(@{ id = '11111111-2222-3333-4444-555555555555'; verifiedDomains = @(@{ name = 'contoso.com'; isDefault = $true }) })
                    }
                }
                '*/beta/security/secureScores*' {
                    return @{
                        value = @(
                            @{
                                createdDateTime = '2026-11-22T00:00:00Z'
                                currentScore    = 0
                                maxScore        = 20
                                controlScores   = @(
                                    @{ controlName = 'mdo_targeteduserprotectionaction'; score = 0 }
                                    @{ controlName = 'CustomEffortControl'; score = 0 }
                                )
                            }
                        )
                    }
                }
                '*/beta/security/secureScoreControlProfiles*' {
                    return @{
                        value = @(
                            @{
                                id                  = 'mdo_targeteduserprotectionaction'
                                title               = 'Targeted user protection action'
                                rank                = 1
                                tier                = 'Core'
                                maxScore            = 10
                                implementationCost  = 'Medium'
                                userImpact          = 'medium'
                                controlStateUpdates = @(@{ state = 'Default' })
                            }
                            @{
                                id                  = 'CustomEffortControl'
                                title               = 'Unrecognised effort labels'
                                rank                = 2
                                tier                = 'Core'
                                maxScore            = 10
                                implementationCost  = ' CUSTOM '
                                userImpact          = 'x'
                                controlStateUpdates = @(@{ state = 'Default' })
                            }
                        )
                    }
                }
                '*/v1.0/identity/conditionalAccess/policies*' { return @{ value = @() } }
                '*/v1.0/policies/identitySecurityDefaultsEnforcementPolicy' { return @{ isEnabled = $false } }
                '*/beta/reports/authenticationMethods/userRegistrationActivity*' { return @{ value = @() } }
                default { throw "Unexpected Graph URI in cost normalisation test: $Uri" }
            }
        }

        $outPath = Join-Path -Path $script:tempDir -ChildPath 'cost-normalise.json'
        $summary = & $scriptPath -TenantDomain 'contoso.com' -DataFile $outPath -Compact

        $compact = Get-Content -LiteralPath $summary.CompactDataFile -Raw | ConvertFrom-Json
        $medium = @($compact.SecureScoreControls.TopOpen | Where-Object { $_.id -eq 'mdo_targeteduserprotectionaction' })
        $medium.Count | Should Be 1
        $medium[0].implementationCost | Should Be 'Moderate'
        $medium[0].userImpact | Should Be 'Moderate'

        $custom = @($compact.SecureScoreControls.TopOpen | Where-Object { $_.id -eq 'CustomEffortControl' })
        $custom.Count | Should Be 1
        $custom[0].implementationCost | Should Be 'Custom'
        $custom[0].userImpact | Should Be 'X'

        $full = Get-Content -LiteralPath $outPath -Raw | ConvertFrom-Json
        $fullMedium = @($full.SecureScoreControls.Open | Where-Object { $_.id -eq 'mdo_targeteduserprotectionaction' })
        $fullMedium[0].implementationCost | Should Be 'Medium'
        $fullMedium[0].userImpact | Should Be 'medium'
    }

    It 'rejects malformed TenantDomain values via parameter validation' {
        Assert-ThrowsWithMessage -Action {
            & $scriptPath -TenantDomain 'not a valid domain!'
        } -ExpectedMessage 'Cannot validate argument on parameter'
    }
}
