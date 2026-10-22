#requires -Version 7.4
<#
.SYNOPSIS
    Offline Linux regression tests; no tenant, modules or Pester installation needed.
.EXAMPLE
    pwsh -NoProfile -File ./tests/o365-connect-tms-cert-linux.Tests.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
if (-not $IsLinux) { throw 'Run these tests on Linux to validate Unix permissions and ephemeral PFX loading.' }
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'o365-connect-tms-cert-linux.ps1'
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

# Mock only service/module boundaries; cryptography, storage and script flow are real.
function Get-Module {
    param([string]$Name, [switch]$ListAvailable)
    if ($Name -notin @('MicrosoftTeams', 'Microsoft.Graph.Authentication')) { throw "Unexpected module: $Name" }
    [pscustomobject]@{ Name = $Name; Version = [version]'99.0.0' }
}

function Connect-MicrosoftTeams {
    [CmdletBinding()]
    param([string]$TenantId, [string]$ApplicationId,
        [Security.Cryptography.X509Certificates.X509Certificate2]$Certificate)
    Assert-True $Certificate.HasPrivateKey 'Teams must receive a certificate with its private key.'
    $testState.Connection = @{ Tenant = $TenantId; AppId = $ApplicationId; Certificate = $Certificate }
}

function Connect-MgGraph {
    [CmdletBinding()]
    param([string]$TenantId, [switch]$UseDeviceAuthentication, [string]$ContextScope,
        [switch]$NoWelcome, [string[]]$Scopes, [string]$ClientId)
    Assert-True ($UseDeviceAuthentication -and $ContextScope -eq 'Process') 'Provisioning uses process-scoped device auth.'
    Assert-True ($Scopes -contains 'RoleManagement.ReadWrite.Directory') 'Provisioning requests directory role assignment scope.'
}

function Disconnect-MgGraph {
    [CmdletBinding()]
    param()
    $testState.GraphDisconnected = $true
}

function Invoke-MgGraphRequest {
    [CmdletBinding()]
    param([string]$Method, [string]$Uri, [string]$OutputType, [string]$Body, [string]$ContentType)
    $path = [uri]::UnescapeDataString($Uri.Replace('https://graph.microsoft.com/v1.0/', ''))
    $payload = if ($Body) { $Body | ConvertFrom-Json -AsHashtable } else { $null }
    $testState.Requests.Add(@{ Method = $Method; Path = $path; Body = $payload })
    if ($Method -eq 'GET') {
        switch -Wildcard ($path) {
            "servicePrincipals[?]*00000003-0000-0000-c000-000000000000*" {
                return @{ value = @(@{ id = 'graph-sp'; appRoles = $testState.GraphRoles }) }
            }
            'roleManagement/directory/roleDefinitions?*' { return @{ value = @(@{ id = 'teams-role' }) } }
            'applications[?]*' {
                if ($testState.CreateNewApp) { return @{ value = @() } }
                return @{ value = @(@{ id = 'app-object'; appId = 'app-client' }) }
            }
            'applications/app-object[?]*' { return $testState.ExistingApp }
            'servicePrincipals[?]*' {
                if ($testState.CreateNewApp) { return @{ value = @() } }
                return @{ value = @(@{ id = 'app-sp' }) }
            }
            'servicePrincipals/app-sp/appRoleAssignments' { return @{ value = $testState.ExistingConsents } }
            'roleManagement/directory/roleAssignments?*' { return @{ value = $testState.ExistingDirectoryRoles } }
            'pagination' { return @{ value = @('first'); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/page2' } }
            'page2' { return @{ value = @('second') } }
            default { throw "Unexpected Graph GET: $path" }
        }
    }
    if ($Method -notin @('PATCH', 'POST')) { throw "Unexpected Graph method: $Method" }
    if ($Method -eq 'POST' -and $path -eq 'applications') { return @{ id = 'app-object'; appId = 'new-client' } }
    if ($Method -eq 'POST' -and $path -eq 'servicePrincipals') { return @{ id = 'app-sp' } }
    return @{}
}

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ("teams-cert-tests-" + [guid]::NewGuid())
$oldData = $env:XDG_DATA_HOME
$oldConfig = $env:XDG_CONFIG_HOME
$oldPassword = $env:CIAOPS_TEAMS_TEST_PASSWORD
$testState = @{ Requests = [Collections.Generic.List[hashtable]]::new() }
$testState.GraphRoles = @(
    foreach ($name in @('Organization.Read.All', 'User.Read.All', 'Group.ReadWrite.All',
        'AppCatalog.ReadWrite.All', 'TeamSettings.ReadWrite.All', 'Channel.Delete.All',
        'ChannelSettings.ReadWrite.All', 'ChannelMember.ReadWrite.All')) {
        @{ id = $name; value = $name; isEnabled = $true; allowedMemberTypes = @('Application') }
    }
)
$testState.ExistingApp = @{
    keyCredentials = @(@{ keyId = 'old-key'; key = 'existing-public-key'; customKeyIdentifier = 'old-hash' })
    requiredResourceAccess = @(@{ resourceAppId = 'other-resource'; resourceAccess = @(@{ id = 'other-role'; type = 'Role' }) })
}
$testState.ExistingConsents = @()
$testState.ExistingDirectoryRoles = @()
try {
    $env:XDG_DATA_HOME = Join-Path $testRoot 'data'
    $env:XDG_CONFIG_HOME = Join-Path $testRoot 'config'
    $env:CIAOPS_TEAMS_TEST_PASSWORD = [guid]::NewGuid().ToString()
    $password = ConvertTo-SecureString $env:CIAOPS_TEAMS_TEST_PASSWORD -AsPlainText -Force
    Assert-Throws { & $scriptPath } '*exactly one*'
    Assert-Throws { & $scriptPath -GenerateLocalCertificate -UseCertificateAuth } '*exactly one*'
    Assert-Throws { & $scriptPath -GenerateLocalCertificate -ProvisionEntraApp -Headless } '*interactive administrator*'
    Assert-Throws { & $scriptPath -GenerateLocalCertificate -Headless } '*Headless mode requires*'
    Assert-Throws { Get-TeamsPfxPassword -NoPrompt -EnvironmentVariable CIAOPS_TEAMS_MISSING_TEST_PASSWORD } '*missing or empty*'
    Assert-Throws { Get-TeamsPfxPassword -Password ([securestring]::new()) -NoPrompt } '*empty PFX password*'

    & $scriptPath -GenerateLocalCertificate -ProvisionEntraApp -Tenant contoso.onmicrosoft.com -CertificatePasswordEnvironmentVariable CIAOPS_TEAMS_TEST_PASSWORD
    Assert-True $testState.GraphDisconnected 'Setup Graph context is disconnected.'
    $profilePath = Join-Path $env:XDG_CONFIG_HOME 'ciaops/tms/o365-tms-cert-auth.json'
    $profiles = @(Read-TeamsProfiles $profilePath)
    Assert-True ($profiles.Count -eq 1) 'Provisioning saves exactly one profile.'
    $profile = $profiles[0]
    $pfxPath = $profile.certificateFilePath
    $cerPath = [IO.Path]::ChangeExtension($pfxPath, '.cer')
    foreach ($path in @($pfxPath, $cerPath, $profilePath)) {
        Assert-True ([int][IO.File]::GetUnixFileMode($path) -eq 384) "File permissions are 600: $path"
    }
    foreach ($path in @((Split-Path $pfxPath), (Split-Path $profilePath))) {
        Assert-True ([int][IO.File]::GetUnixFileMode($path) -eq 448) "Directory permissions are 700: $path"
    }
    Assert-True (-not (Get-Content $profilePath -Raw).Contains($env:CIAOPS_TEAMS_TEST_PASSWORD)) 'Profile contains no password.'
    $publicCert = [Security.Cryptography.X509Certificates.X509Certificate2]::new($cerPath)
    try {
        Assert-True (-not $publicCert.HasPrivateKey) 'CER has no private key.'
        $publicKey = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($publicCert)
        try { Assert-True ($publicKey.ExportParameters($false).Modulus.Length * 8 -eq 2048) 'Certificate uses RSA-2048.' }
        finally { $publicKey.Dispose() }
        Assert-True ($publicCert.SignatureAlgorithm.FriendlyName -eq 'sha256RSA') 'Certificate uses SHA256.'
    }
    finally { $publicCert.Dispose() }
    $patch = @($testState.Requests | Where-Object Method -eq PATCH)[0].Body
    Assert-True ($patch.keyCredentials.Count -eq 2) 'Existing certificate is preserved.'
    Assert-True ($patch.requiredResourceAccess.Count -eq 2) 'Other resource permissions are preserved.'
    $consents = @($testState.Requests | Where-Object { $_.Method -eq 'POST' -and $_.Path -like '*/appRoleAssignments' })
    Assert-True ($consents.Count -eq 8) 'Exactly the eight documented Graph application permissions are consented.'
    Assert-True (@($consents | Where-Object { $_.Body.resourceId -ne 'graph-sp' }).Count -eq 0) 'No Teams Tenant Admin API permission is granted.'
    $directoryRole = @($testState.Requests | Where-Object Path -eq 'roleManagement/directory/roleAssignments')[0].Body
    Assert-True ($directoryRole.roleDefinitionId -eq 'teams-role' -and $directoryRole.directoryScopeId -eq '/') 'Teams Administrator is assigned at tenant scope.'

    & $scriptPath -UseCertificateAuth -ProfileName O365-TMS-AppAuth -Headless -CertificatePassword $password
    Assert-True ($testState.Connection.Tenant -eq 'contoso.onmicrosoft.com' -and $testState.Connection.AppId -eq 'app-client') 'Saved profile selects the correct tenant/app.'
    Assert-True ($testState.Connection.Certificate.Thumbprint -eq $profile.certificateThumbprint) 'Teams receives the selected PFX.'
    $testState.Connection.Certificate.Dispose()
    & $scriptPath -UseCertificateAuth -Organization contoso.onmicrosoft.com -AppId explicit-app -CertificateFilePath $pfxPath -Headless -CertificatePasswordEnvironmentVariable CIAOPS_TEAMS_TEST_PASSWORD
    Assert-True ($testState.Connection.AppId -eq 'explicit-app') 'Explicit connection bypasses profiles and supports Organization alias.'
    $testState.Connection.Certificate.Dispose()
    Assert-Throws { & $scriptPath -UseCertificateAuth -ProfileName O365-TMS-AppAuth -CertificateThumbprint DEADBEEF -CertificatePassword $password -Headless } '*thumbprint does not match*'
    Assert-Throws { & $scriptPath -UseCertificateAuth -Tenant contoso.onmicrosoft.com -AppId explicit-app -CertificateFilePath $cerPath -CertificatePassword $password -Headless } '*no private key*'
    Assert-Throws { Import-TeamsLinuxPfx -Path $pfxPath -Password (ConvertTo-SecureString 'wrong-password' -AsPlainText -Force) } '*'
    [IO.File]::SetUnixFileMode($pfxPath, [IO.UnixFileMode]::UserRead -bor [IO.UnixFileMode]::GroupRead)
    Assert-Throws { Import-TeamsLinuxPfx -Path $pfxPath -Password $password } '*owner-only*'
    [IO.File]::SetUnixFileMode($pfxPath, [IO.UnixFileMode]::UserRead -bor [IO.UnixFileMode]::UserWrite)
    $linkPath = Join-Path $testRoot 'linked.pfx'
    [IO.File]::CreateSymbolicLink($linkPath, $pfxPath) | Out-Null
    Assert-Throws { Import-TeamsLinuxPfx -Path $linkPath -Password $password } '*symbolic link*'
    $directoryLink = Join-Path $testRoot 'linked-directory'
    [IO.Directory]::CreateSymbolicLink($directoryLink, (Split-Path $pfxPath)) | Out-Null
    Assert-Throws { Initialize-TeamsPrivateDirectory (Join-Path $directoryLink 'nested') } '*Symbolic links*'

    Save-TeamsLinuxProfile $profilePath $profile
    Assert-True (@(Read-TeamsProfiles $profilePath).Count -eq 1) 'Saving a profile again does not duplicate it.'
    $relativeMap = Join-Path (Split-Path $pfxPath) 'relative-profile.json'
    $relativeProfile = $profile.Clone()
    $relativeProfile.certificateFilePath = Split-Path $pfxPath -Leaf
    Save-TeamsLinuxProfile $relativeMap $relativeProfile
    & $scriptPath -UseCertificateAuth -CertificateMapPath $relativeMap -Headless -CertificatePassword $password
    Assert-True ($testState.Connection.Certificate.Thumbprint -eq $profile.certificateThumbprint) 'Relative PFX paths resolve against the profile map.'
    $testState.Connection.Certificate.Dispose()
    $secondProfile = $profile.Clone()
    $secondProfile.name = 'Second'
    Save-TeamsLinuxProfile $profilePath $secondProfile
    Assert-Throws { & $scriptPath -UseCertificateAuth -Tenant contoso.onmicrosoft.com -Headless -CertificatePassword $password } '*found 2*'
    $pages = @(Get-TeamsLinuxGraphCollection 'pagination')
    Assert-True ($pages.Count -eq 2 -and $pages[1] -eq 'second') 'Graph collection pagination is followed.'

    $testState.ExistingConsents = @($consents | ForEach-Object { $_.Body })
    $testState.ExistingDirectoryRoles = @($directoryRole)
    $testState.Requests.Clear()
    New-TeamsLinuxEntraApp -DisplayName O365-TMS-AppAuth -CerPath $cerPath | Out-Null
    Assert-True (@($testState.Requests | Where-Object Method -eq POST).Count -eq 0) 'Existing consent and directory assignments are not duplicated.'
    $testState.ExistingApp.keyCredentials[0].Remove('key')
    Assert-Throws { New-TeamsLinuxEntraApp -DisplayName O365-TMS-AppAuth -CerPath $cerPath } '*Refusing to remove existing certificates*'
    $testState.CreateNewApp = $true
    $testState.ExistingApp = @{ keyCredentials = @(); requiredResourceAccess = @() }
    $testState.ExistingConsents = @()
    $testState.ExistingDirectoryRoles = @()
    $testState.Requests.Clear()
    $newApp = New-TeamsLinuxEntraApp -DisplayName NewTeamsApp -CerPath $cerPath
    Assert-True ($newApp.AppId -eq 'new-client') 'A new application returns its client ID.'
    Assert-True (@($testState.Requests | Where-Object { $_.Method -eq 'POST' -and $_.Path -eq 'servicePrincipals' }).Count -eq 1) 'A new application gets a service principal.'
    $newPermissions = @($testState.Requests | Where-Object Method -eq PATCH)[0].Body.requiredResourceAccess
    Assert-True ($newPermissions.Count -eq 1 -and $newPermissions[0].resourceAccess.Count -eq 8) 'A new app gets exactly the documented Graph permissions.'
    Write-Host 'All offline Teams Linux certificate regression tests passed.'
}
finally {
    $env:XDG_DATA_HOME = $oldData
    $env:XDG_CONFIG_HOME = $oldConfig
    $env:CIAOPS_TEAMS_TEST_PASSWORD = $oldPassword
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
