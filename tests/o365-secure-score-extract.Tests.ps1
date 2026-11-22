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
                            @{ createdDateTime = '2026-03-02T00:00:00Z'; currentScore = 82.5; maxScore = 100 }
                            @{ createdDateTime = '2026-03-01T00:00:00Z'; currentScore = 78.0; maxScore = 100 }
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
    }

    It 'extracts full and compact JSON files (supporting spaces in paths) and emits a summary object' {
        $outPath = Join-Path -Path $script:tempDir -ChildPath 'sub folder\contoso export.json'
        $summary = & $scriptPath -TenantDomain 'contoso.com' -DataFile $outPath -Compact

        $summary | Should Not BeNullOrEmpty
        $summary.Tenant | Should Be 'contoso.com'
        $summary.LatestScore | Should Be 82.5
        $summary.OpenControlCount | Should Be 1
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

    It 'rejects malformed TenantDomain values via parameter validation' {
        Assert-ThrowsWithMessage -Action {
            & $scriptPath -TenantDomain 'not a valid domain!'
        } -ExpectedMessage 'Cannot validate argument on parameter'
    }
}
