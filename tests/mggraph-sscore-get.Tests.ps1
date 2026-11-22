#requires -Version 5.1
<#
.SYNOPSIS
    Offline Pester tests for the Microsoft Graph Secure Score report.
.EXAMPLE
    Invoke-Pester -Script .\tests\mggraph-sscore-get.Tests.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'mggraph-sscore-get.ps1'

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

Describe 'mggraph-sscore-get' {
    BeforeEach {
        Mock Import-Module {
            [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication' }
        }
        Mock Connect-MgGraph {}
        Mock Get-MgContext {
            [pscustomobject]@{ Account = 'analyst@contoso.com'; TenantId = 'tenant-id' }
        }
        Mock Invoke-MgGraphRequest {
            @{
                value = @(
                    @{
                        createdDateTime = '2026-03-01T00:00:00Z'
                        currentScore    = 75
                        maxScore        = 100
                    }
                    @{
                        createdDateTime = '2026-03-02T00:00:00Z'
                        currentScore    = 40
                        maxScore        = 80
                    }
                )
            }
        }
    }

    It 'emits one typed result per Graph record with a calculated percentage' {
        $records = @(& $scriptPath)

        $records.Count | Should Be 2
        $records[0].CreatedDateTime.ToString('yyyy-MM-dd') | Should Be '2026-03-01'
        $records[0].CurrentScore | Should Be 75
        $records[0].MaxScore | Should Be 100
        $records[0].ScorePercent | Should Be 75
        $records[1].ScorePercent | Should Be 50
        $LASTEXITCODE | Should Be 0
    }

    It 'uses the default Microsoft Graph endpoint' {
        $records = @(& $scriptPath)

        $records.Count | Should Be 2
        Assert-MockCalled Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            $Uri -eq 'https://graph.microsoft.com/beta/security/securescores'
        }
    }

    It 'supports an alternate Microsoft Graph endpoint' {
        $records = @(& $scriptPath -GraphEndpoint 'https://graph.microsoft.us/beta/security/securescores')

        $records.Count | Should Be 2
        Assert-MockCalled Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            $Uri -eq 'https://graph.microsoft.us/beta/security/securescores'
        }
    }

    It 'reuses an active Graph session when required scopes are already present' {
        Mock Get-MgContext {
            [pscustomobject]@{
                Account  = 'analyst@contoso.com'
                TenantId = 'tenant-id'
                Scopes   = @('SecurityEvents.Read.All', 'User.Read')
            }
        }

        $records = @(& $scriptPath)

        $records.Count | Should Be 2
        Assert-MockCalled Connect-MgGraph -Times 0 -Scope It
    }

    It 'reconnects when ForceReconnect is specified despite an active session' {
        Mock Get-MgContext {
            [pscustomobject]@{
                Account  = 'analyst@contoso.com'
                TenantId = 'tenant-id'
                Scopes   = @('SecurityEvents.Read.All')
            }
        }

        $records = @(& $scriptPath -ForceReconnect)

        $records.Count | Should Be 2
        Assert-MockCalled Connect-MgGraph -Times 1 -Scope It
    }

    It 'follows @odata.nextLink pagination across multiple pages' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -eq 'https://graph.microsoft.com/beta/security/securescores') {
                return @{
                    value             = @(
                        @{ createdDateTime = '2026-03-01T00:00:00Z'; currentScore = 80; maxScore = 100 }
                    )
                    '@odata.nextLink' = 'https://graph.microsoft.com/beta/security/securescores?$skip=1'
                }
            }
            return @{
                value = @(
                    @{ createdDateTime = '2026-03-02T00:00:00Z'; currentScore = 60; maxScore = 100 }
                )
            }
        }

        $records = @(& $scriptPath)

        $records.Count | Should Be 2
        $records[0].ScorePercent | Should Be 80
        $records[1].ScorePercent | Should Be 60
        Assert-MockCalled Invoke-MgGraphRequest -Times 2 -Scope It
    }

    It 'limits returned records and appends $top when Top is specified' {
        Mock Invoke-MgGraphRequest {
            @{
                value             = @(
                    @{ createdDateTime = '2026-03-01T00:00:00Z'; currentScore = 90; maxScore = 100 }
                    @{ createdDateTime = '2026-03-02T00:00:00Z'; currentScore = 70; maxScore = 100 }
                )
                '@odata.nextLink' = 'https://graph.microsoft.com/beta/security/securescores?$skip=2'
            }
        }

        $records = @(& $scriptPath -Top 1)

        $records.Count | Should Be 1
        $records[0].ScorePercent | Should Be 90
        Assert-MockCalled Invoke-MgGraphRequest -Times 1 -Scope It -ParameterFilter {
            $Uri -eq 'https://graph.microsoft.com/beta/security/securescores?$top=1'
        }
    }

    It 'announces each potentially slow operation before it begins' {
        Mock Import-Module { Write-Information -MessageData 'import started' -InformationAction Continue }
        Mock Connect-MgGraph { Write-Information -MessageData 'connection started' -InformationAction Continue }
        Mock Invoke-MgGraphRequest {
            Write-Information -MessageData 'request started' -InformationAction Continue
            @{ value = @() }
        }

        $messages = @(& $scriptPath 6>&1 | ForEach-Object { $_.MessageData })

        $messages.Count | Should Be 7
        $messages[0] | Should Match '^Loading the Microsoft\.Graph\.Authentication module'
        $messages[1] | Should Be 'import started'
        $messages[2] | Should Match '^Connecting to Microsoft Graph'
        $messages[3] | Should Be 'connection started'
        $messages[4] | Should Match '^Retrieving Secure Score history'
        $messages[5] | Should Be 'request started'
        $messages[6] | Should Match '^Secure Score report completed.*Elapsed:'
    }

    It 'returns no records when Graph returns an empty collection' {
        Mock Invoke-MgGraphRequest { @{ value = @() } }

        $records = @(& $scriptPath)

        $records.Count | Should Be 0
    }

    It 'uses a null percentage when the maximum score is zero' {
        Mock Invoke-MgGraphRequest {
            @{
                value = @(
                    @{
                        createdDateTime = '2026-03-01T00:00:00Z'
                        currentScore    = 0
                        maxScore        = 0
                    }
                )
            }
        }

        $records = @(& $scriptPath)

        $records.Count | Should Be 1
        $records[0].ScorePercent | Should Be $null
    }

    It 'skips malformed records (including non-numeric or boolean scores) and continues emitting valid records' {
        Mock Invoke-MgGraphRequest {
            @{
                value = @(
                    @{
                        createdDateTime = 'not-a-date'
                        currentScore    = 10
                        maxScore        = 100
                    }
                    @{
                        createdDateTime = '2026-03-01T00:00:00Z'
                        currentScore    = $true
                        maxScore        = 100
                    }
                    @{
                        createdDateTime = '2026-03-02T00:00:00Z'
                        currentScore    = 40
                        maxScore        = 80
                    }
                )
            }
        }

        $records = @(& $scriptPath -WarningVariable warnings -WarningAction Continue)

        $records.Count | Should Be 1
        $records[0].ScorePercent | Should Be 50
        $warnings.Count | Should Be 2
        $warnings[0] | Should Match 'Skipping Secure Score record 1'
        $warnings[1] | Should Match 'Skipping Secure Score record 2'
    }

    It 'creates missing transcript parent directories and stops the transcript when Debug is used' {
        Mock Start-Transcript {}
        Mock Stop-Transcript {}
        $tempRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ('sscore-test-' + [guid]::NewGuid().ToString('N'))
        $transcriptFile = Join-Path -Path $tempRoot -ChildPath 'nested folder\secure score.txt'

        try {
            $records = @(& $scriptPath -Debug -TranscriptPath $transcriptFile)

            $records.Count | Should Be 2
            (Test-Path -LiteralPath (Split-Path -Path $transcriptFile -Parent)) | Should Be $true
            Assert-MockCalled Start-Transcript -Times 1 -ParameterFilter {
                $Path -eq $transcriptFile -and $Append
            }
            Assert-MockCalled Stop-Transcript -Times 1
        }
        finally {
            if (Test-Path -LiteralPath $tempRoot) {
                Remove-Item -LiteralPath $tempRoot -Recurse -Force
            }
        }
    }

    It 'fails with an installation hint and sets LASTEXITCODE to 1 when the Graph module is missing' {
        Mock Import-Module { throw 'module not found' }

        Assert-ThrowsWithMessage -Action { & $scriptPath } -ExpectedMessage 'Install the module if it is missing'
        $LASTEXITCODE | Should Be 1
    }

    It 'adds request context when the Graph request fails' {
        Mock Invoke-MgGraphRequest { throw 'backend unavailable' }

        Assert-ThrowsWithMessage -Action { & $scriptPath } -ExpectedMessage 'Secure Score request failed'
        $LASTEXITCODE | Should Be 1
    }
}
