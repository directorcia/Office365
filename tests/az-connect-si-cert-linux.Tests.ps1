#requires -Version 7.4
<#
.SYNOPSIS
    Offline tenant validation regression checks; no Az modules or Pester required.
.EXAMPLE
    pwsh -NoProfile -File ./tests/az-connect-si-cert-linux.Tests.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'az-connect-si-cert-linux.ps1'
$tokens = $null
$parseErrors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }

function Assert-Throws {
    param([scriptblock]$Action, [string]$Pattern)
    $caught = $null
    try { & $Action }
    catch { $caught = $_ }
    if ($null -eq $caught -or $caught.Exception.Message -notlike $Pattern) {
        throw "Expected error matching '$Pattern'; got '$caught'."
    }
}

function Read-Host { throw 'Unexpected password prompt.' }
function Get-Module { throw 'Unexpected module lookup.' }
function Connect-AzAccount { throw 'Unexpected Azure sign-in.' }

$profilePath = Join-Path ([IO.Path]::GetTempPath()) ("sentinel-tenant-test-" + [guid]::NewGuid() + '.json')
$tenantId = '99999999-8888-7777-6666-555555555555'
$connection = @{
    AppId = '11111111-2222-3333-4444-555555555555'
    SubscriptionId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
    CertificateFilePath = 'unused.pfx'
    CertificateMapPath = $profilePath
    Headless = $true
}
$tenantError = 'Tenant must be a non-empty Entra tenant ID GUID*tenant domains are not supported*'
$passwordError = 'Headless mode requires -CertificatePassword*'
try {
    foreach ($invalidTenant in @('ciaopslabs.onmicrosoft.com', 'not-a-guid', '00000000-0000-0000-0000-000000000000')) {
        Assert-Throws { & $scriptPath @connection -Tenant $invalidTenant } $tenantError
        Assert-Throws { & $scriptPath @connection -Tenant $invalidTenant -Headless:$false } $tenantError
    }
    Assert-Throws { & $scriptPath @connection -Tenant $tenantId } $passwordError
    Assert-Throws { & $scriptPath @connection -TenantId $tenantId } $passwordError

    foreach ($profileTenant in @('ciaopslabs.onmicrosoft.com', $tenantId)) {
        @{
            profiles = @(@{
                name = 'sentinel'
                tenant = $profileTenant
                appId = $connection.AppId
                subscriptionId = $connection.SubscriptionId
                certificateFilePath = $connection.CertificateFilePath
            })
        } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $profilePath
        $expectedError = if ($profileTenant -eq $tenantId) { $passwordError } else { $tenantError }
        Assert-Throws {
            & $scriptPath -ProfileName sentinel -CertificateMapPath $profilePath -Headless
        } $expectedError
        Assert-Throws {
            & $scriptPath -ProfileName sentinel -CertificateMapPath $profilePath -Tenant $tenantId -Headless
        } $passwordError
    }
    Write-Host 'All Sentinel tenant validation regression checks passed.'
}
finally {
    if (Test-Path -LiteralPath $profilePath) { Remove-Item -LiteralPath $profilePath -Force }
}
