<#
.SYNOPSIS
    Offline regression checks for Windows Exchange Online certificate profile handling.
.EXAMPLE
    Invoke-Pester -Script ./tests/o365-connect-exo-cert.Tests.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'o365-connect-exo-cert.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
foreach ($functionName in @('Resolve-ExoCertificateProfile', 'Set-ExoProfileMapEntry', 'New-ExoLocalCertificate')) {
    $definition = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName
    }, $true)
    . ([scriptblock]::Create($definition.Extent.Text))
}

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

function Assert-Throws {
    param([scriptblock]$Action, [string]$Pattern)
    $caught = $null
    try { & $Action }
    catch { $caught = $_ }
    Assert-True ($null -ne $caught -and $caught.Exception.Message -like $Pattern) "Expected error matching '$Pattern'; got '$caught'."
}

Describe 'o365-connect-exo-cert profile and export behavior' {
    BeforeEach {
        $script:testRoot = Join-Path ([IO.Path]::GetTempPath()) ('exo-cert-tests-' + [guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($script:testRoot)
    }

    AfterEach {
        Microsoft.PowerShell.Management\Remove-Item -LiteralPath $script:testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'preserves separate app profiles in one tenant and updates by app ID' {
        $colors = @{ ProcessMessage = 'Green' }
        $mapPath = Join-Path $script:testRoot 'profiles [test].json'
        $first = [pscustomobject]@{
            name = 'Primary'
            tenant = 'tenant.example'
            organization = 'tenant.example'
            appId = 'app-1'
            certificateThumbprint = 'thumb-1'
        }
        $second = [pscustomobject]@{
            name = 'Secondary'
            tenant = 'tenant.example'
            organization = 'tenant.example'
            appId = 'app-2'
            certificateThumbprint = 'thumb-2'
        }

        Set-ExoProfileMapEntry -MapPath $mapPath -ProfileEntry $first -Colors $colors
        Set-ExoProfileMapEntry -MapPath $mapPath -ProfileEntry $second -Colors $colors
        $savedProfiles = (Get-Content -LiteralPath $mapPath -Raw | ConvertFrom-Json).profiles
        Assert-True ($savedProfiles.Count -eq 2) 'Different app profiles in one tenant must both remain in the map.'

        $updatedFirst = [pscustomobject]@{
            name = 'Primary'
            tenant = 'tenant.example'
            organization = 'tenant.example'
            appId = 'app-1'
            certificateThumbprint = 'thumb-updated'
        }
        Set-ExoProfileMapEntry -MapPath $mapPath -ProfileEntry $updatedFirst -Colors $colors
        $savedProfiles = (Get-Content -LiteralPath $mapPath -Raw | ConvertFrom-Json).profiles
        Assert-True ($savedProfiles.Count -eq 2) 'Updating an app profile must not add a duplicate.'
        Assert-True (($savedProfiles | Where-Object appId -eq 'app-1').certificateThumbprint -eq 'thumb-updated') 'An existing app profile must be replaced by app ID.'
    }

    It 'reads and updates a one-entry bare-array profile map' {
        $colors = @{ ProcessMessage = 'Green' }
        $first = [pscustomobject]@{
            name = 'Primary'
            tenant = 'tenant.example'
            organization = 'tenant.example'
            appId = 'app-1'
            certificateThumbprint = 'thumb-1'
        }
        $second = [pscustomobject]@{
            name = 'Secondary'
            tenant = 'tenant.example'
            organization = 'tenant.example'
            appId = 'app-2'
            certificateThumbprint = 'thumb-2'
        }
        $arrayMapPath = Join-Path $script:testRoot 'array-map.json'
        ConvertTo-Json -InputObject @($first) -Depth 4 | Set-Content -LiteralPath $arrayMapPath

        $resolved = Resolve-ExoCertificateProfile -Path $arrayMapPath -ProfileFilter 'Primary' -NoPrompt -Colors $colors
        Assert-True ($resolved.appId -eq 'app-1') 'A one-entry bare JSON array must resolve as a profile.'

        Set-ExoProfileMapEntry -MapPath $arrayMapPath -ProfileEntry $second -Colors $colors
        $savedProfiles = (Get-Content -LiteralPath $arrayMapPath -Raw | ConvertFrom-Json).profiles
        Assert-True ($savedProfiles.Count -eq 2) 'Updating a one-entry bare array map must preserve its existing profile.'
    }

    It 'preserves existing profile records without an app ID' {
        $colors = @{ ProcessMessage = 'Green' }
        $legacyProfile = [pscustomobject]@{ name = 'Legacy'; tenant = 'tenant.example'; organization = 'tenant.example' }
        $newProfile = [pscustomobject]@{ name = 'New'; tenant = 'tenant.example'; organization = 'tenant.example'; appId = 'app-2' }
        $mapPath = Join-Path $script:testRoot 'legacy-map.json'
        ConvertTo-Json -InputObject @($legacyProfile) -Depth 4 | Set-Content -LiteralPath $mapPath

        Set-ExoProfileMapEntry -MapPath $mapPath -ProfileEntry $newProfile -Colors $colors
        $savedProfiles = (Get-Content -LiteralPath $mapPath -Raw | ConvertFrom-Json).profiles
        Assert-True ($savedProfiles.Count -eq 2) 'A legacy profile without appId must be retained when adding another entry.'
    }

    It 'does not overwrite a corrupt profile map' {
        $colors = @{ ProcessMessage = 'Green' }
        $corruptMapPath = Join-Path $script:testRoot 'corrupt.json'
        $corruptContent = '{ invalid json'
        $profile = [pscustomobject]@{ name = 'Primary'; tenant = 'tenant.example'; organization = 'tenant.example'; appId = 'app-1' }
        Set-Content -LiteralPath $corruptMapPath -Value $corruptContent -NoNewline

        Assert-Throws {
            Set-ExoProfileMapEntry -MapPath $corruptMapPath -ProfileEntry $profile -Colors $colors
        } '*Unable to read existing certificate profile map*'
        Assert-True ((Get-Content -LiteralPath $corruptMapPath -Raw) -eq $corruptContent) 'A corrupt profile map must remain unchanged.'
    }

    It 'honors WhatIf before loading modules or making changes' {
        $powerShellPath = (Get-Process -Id $PID).MainModule.FileName
        $output = & $powerShellPath -NoProfile -ExecutionPolicy Bypass -File $scriptPath -UseCertificateAuth -WhatIf 2>&1 | Out-String
        Assert-True ($LASTEXITCODE -eq 0) 'A WhatIf preview must exit successfully.'
        Assert-True ($output -like '*What if:*Connect to Exchange Online using certificate authentication*') 'WhatIf must report the selected operation.'
        Assert-True ($output -notlike '*Loading Exchange Online PowerShell module*') 'WhatIf must stop before loading or installing modules.'
    }

    It 'removes partial certificate files and the store certificate when PFX export fails' {
        $script:removedCertificate = $null
        function New-SelfSignedCertificate {
            [CmdletBinding()]
            param([string]$Subject, [string]$CertStoreLocation, [string]$KeyAlgorithm, [int]$KeyLength,
                [string]$HashAlgorithm, [string]$KeyExportPolicy, [datetime]$NotAfter, [string]$KeySpec)
            [pscustomobject]@{ Thumbprint = 'ABCDEF0123456789ABCDEF0123456789ABCDEF01'; Subject = $Subject; NotAfter = $NotAfter; FriendlyName = '' }
        }
        function Export-Certificate {
            [CmdletBinding()]
            param([object]$Cert, [string]$FilePath, [string]$Type, [switch]$Force)
            [IO.File]::WriteAllText($FilePath, 'partial cer')
        }
        function Export-PfxCertificate {
            [CmdletBinding()]
            param([object]$Cert, [string]$FilePath, [securestring]$Password, [switch]$Force)
            [IO.File]::WriteAllText($FilePath, 'partial pfx')
            throw 'Simulated PFX export failure.'
        }
        function Remove-Item {
            [CmdletBinding()]
            param([string]$LiteralPath, [switch]$Force)
            if ($LiteralPath -like 'Cert:\*') {
                $script:removedCertificate = $LiteralPath
                return
            }
            Microsoft.PowerShell.Management\Remove-Item -LiteralPath $LiteralPath -Force:$Force -ErrorAction SilentlyContinue
        }

        $outputPath = Join-Path $script:testRoot '[certificate output]'
        $password = ConvertTo-SecureString 'test-password' -AsPlainText -Force
        Assert-Throws {
            New-ExoLocalCertificate -SubjectName 'test' -YearsValid 1 -OutputPath $outputPath -ExportPfx -PfxPassword $password -NoPrompt
        } '*partial export files were removed*'
        Assert-True ($script:removedCertificate -eq 'Cert:\CurrentUser\My\ABCDEF0123456789ABCDEF0123456789ABCDEF01') 'The generated certificate must be removed after export failure.'
        Assert-True (@(Get-ChildItem -LiteralPath $outputPath -Force).Count -eq 0) 'Partial CER/PFX files must be removed after export failure.'
    }
}
