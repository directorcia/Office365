#requires -Version 7.4
<#
.SYNOPSIS
    Offline Linux regression tests for o365-connect-pnp-cert-linux.ps1.
    No tenant, Graph module, or Pester installation required.
.EXAMPLE
    pwsh -NoProfile -File ./tests/o365-connect-pnp-cert-linux.Tests.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
if (-not $IsLinux) { throw 'Run these tests on Linux to validate Unix permissions and ephemeral PFX loading.' }
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'o365-connect-pnp-cert-linux.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }

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

function Get-ParameterAst {
    param([string]$Name)
    return @($ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq $Name })[0]
}

$scriptText = Get-Content -LiteralPath $scriptPath -Raw
Assert-True ($scriptText -match 'Version:\s*1\.1') 'Script version marker is 1.1.'
Assert-True ($scriptText -notmatch '\$env:TMPDIR') 'Temp paths are not built from TMPDIR.'
Assert-True ($scriptText -match '\[System\.IO\.Path\]::GetTempPath\(\)') 'Temp directory resolution uses GetTempPath.'
$timeoutParameter = Get-ParameterAst 'DeviceCodeTimeoutSeconds'
Assert-True ($timeoutParameter.DefaultValue.Extent.Text -eq '900') 'Device-code timeout defaults to 900 seconds.'
$permissionParameter = Get-ParameterAst 'SharePointPermission'
Assert-True ($permissionParameter.DefaultValue.Extent.Text -eq "'Sites.FullControl.All'") 'SharePoint permission defaults to Sites.FullControl.All.'
$certificateParameter = Get-ParameterAst 'CertificateFilePath'
$certificateAliases = @($certificateParameter.Attributes | Where-Object { $_.TypeName.Name -eq 'Alias' } | ForEach-Object { $_.PositionalArguments.Value })
Assert-True ($certificateAliases -contains 'CertificatePath') 'CertificateFilePath accepts -CertificatePath.'

foreach ($definition in $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
}, $false)) {
    . ([scriptblock]::Create($definition.Extent.Text))
}

function New-TestJwt {
    param([string]$TenantId)
    $header = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('{"alg":"none"}')).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $payloadJson = @{ tid = $TenantId; aud = 'https://graph.microsoft.com' } | ConvertTo-Json -Compress
    $payload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($payloadJson)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    return "$header.$payload.sig"
}

function New-OAuthError {
    param([string]$Code)
    $json = @{ error = $Code; error_description = $Code } | ConvertTo-Json -Compress
    $exception = [System.Exception]::new("Response status code does not indicate success: 400 ($Code).")
    $record = [System.Management.Automation.ErrorRecord]::new($exception, 'OAuthError', 'InvalidOperation', $null)
    $record.ErrorDetails = [System.Management.Automation.ErrorDetails]::new($json)
    return $record
}

$testState = @{
    Requests = [Collections.Generic.List[hashtable]]::new()
    Rest = [Collections.Generic.List[hashtable]]::new()
    CreateNewApp = $true
    ExistingConsents = @()
    TokenResults = @('ok')
    TokenIndex = 0
    DeviceResults = @('ok')
    DeviceIndex = 0
    AlwaysPending = $false
    Interval = 1
    AccessTokenToReturn = New-TestJwt 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
    TenantIds = @{
        'ciaops365.com' = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        'ciaops365e1.onmicrosoft.com' = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        'other.onmicrosoft.com' = '11111111-2222-3333-4444-555555555555'
    }
}
$testState.SpRoles = @(
    foreach ($name in @('Sites.FullControl.All', 'Sites.Read.All', 'Sites.Selected')) {
        @{ id = "sp-$name"; value = $name; isEnabled = $true; allowedMemberTypes = @('Application') }
    }
)
$testState.GraphRoles = @(
    @{ id = 'graph-app-read'; value = 'Application.Read.All'; isEnabled = $true; allowedMemberTypes = @('Application') }
)
$testState.ExistingApp = @{ keyCredentials = @(); requiredResourceAccess = @() }

function Get-Module {
    param([string]$Name, [switch]$ListAvailable)
    if ($Name -and $Name -notin @('PnP.PowerShell', 'Microsoft.Graph.Authentication')) { throw "Unexpected module: $Name" }
    [pscustomobject]@{ Name = $Name; Version = [version]'99.0.0' }
}
function Import-Module { [CmdletBinding()] param([string]$Name, [version]$MinimumVersion) }
function Install-Module { throw 'Install-Module should not run in offline tests.' }
function Find-Module { throw 'Find-Module should not run in offline tests.' }
function Read-Host { throw 'Unexpected prompt.' }
function Connect-MgGraph {
    [CmdletBinding()]
    param(
        [securestring]$AccessToken,
        [switch]$NoWelcome,
        [double]$ClientTimeout,
        [string]$TenantId,
        [switch]$UseDeviceAuthentication,
        [string]$ContextScope,
        [string[]]$Scopes,
        [string]$ClientId
    )
    if ($UseDeviceAuthentication -or $ContextScope) {
        throw 'Provisioning must pass -AccessToken and must not use the SDK device-code timeout.'
    }
    Assert-True ($null -ne $AccessToken -and $AccessToken.Length -gt 0) 'Access token is supplied to Connect-MgGraph.'
    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($AccessToken)
    try { $testState.PlainAccessToken = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer) }
    $testState.GraphConnected = $true
}
function Disconnect-MgGraph { [CmdletBinding()] param() $testState.GraphDisconnected = $true }
function Connect-PnPOnline {
    [CmdletBinding()]
    param([string]$Url, [string]$Tenant, [string]$ClientId, [string]$CertificatePath, [securestring]$CertificatePassword, [switch]$ValidateConnection)
    $testState.Connection = @{ Url = $Url; Tenant = $Tenant; AppId = $ClientId; CertificatePath = $CertificatePath }
}
function Invoke-MgGraphRequest {
    [CmdletBinding()]
    param([string]$Method = 'GET', [string]$Uri, [string]$OutputType, [string]$Body, [string]$ContentType)
    $path = [uri]::UnescapeDataString($Uri.Replace('https://graph.microsoft.com/v1.0/', ''))
    $payload = if ($Body) { $Body | ConvertFrom-Json -AsHashtable } else { $null }
    $testState.Requests.Add(@{ Method = $Method; Path = $path; Body = $payload })
    if ($Method -eq 'GET' -and $path -like 'applications/app-object*') { return $testState.ExistingApp }
    if ($Method -eq 'GET' -and $path -like 'servicePrincipals/app-sp/appRoleAssignments') {
        return @{ value = @($testState.ExistingConsents) }
    }
    if ($Method -eq 'GET' -and $path -like 'servicePrincipals[?]*00000003-0000-0ff1-ce00-000000000000*') {
        return @{ value = @(@{ id = 'spo-sp'; appRoles = $testState.SpRoles }) }
    }
    if ($Method -eq 'GET' -and $path -like 'servicePrincipals[?]*00000003-0000-0000-c000-000000000000*') {
        return @{ value = @(@{ id = 'graph-sp'; appRoles = $testState.GraphRoles }) }
    }
    if ($Method -eq 'GET' -and $path -like 'applications[?]*') {
        if ($testState.CreateNewApp) { return @{ value = @() } }
        return @{ value = @(@{ id = 'app-object'; appId = 'app-client' }) }
    }
    if ($Method -eq 'GET' -and $path -like 'servicePrincipals[?]*') {
        if ($testState.CreateNewApp) { return @{ value = @() } }
        return @{ value = @(@{ id = 'app-sp' }) }
    }
    if ($Method -eq 'POST' -and $path -eq 'applications') { return @{ id = 'app-object'; appId = 'app-client' } }
    if ($Method -eq 'POST' -and $path -eq 'servicePrincipals') { return @{ id = 'app-sp' } }
    if ($Method -eq 'POST' -and $path -like 'servicePrincipals/*/appRoleAssignments') { return @{} }
    if ($Method -eq 'PATCH') { return @{} }
    throw "Unexpected Graph $Method $path"
}
function Invoke-RestMethod {
    [CmdletBinding()]
    param($Method, $Uri, $Body, $ContentType, $TimeoutSec)
    $testState.Rest.Add(@{ Method = [string]$Method; Uri = [string]$Uri; TimeoutSec = $TimeoutSec; Body = $Body })
    if ($Uri -like '*/.well-known/openid-configuration') {
        $tenantKey = [uri]::UnescapeDataString(([uri]$Uri).Segments[1].TrimEnd('/')).ToLowerInvariant()
        if (-not $testState.TenantIds.ContainsKey($tenantKey)) { throw [System.InvalidOperationException]::new("Unknown tenant $tenantKey") }
        return [pscustomobject]@{ issuer = "https://login.microsoftonline.com/$($testState.TenantIds[$tenantKey])/v2.0" }
    }
    if ($Uri -like '*/devicecode') {
        $result = 'ok'
        if ($testState.DeviceIndex -lt @($testState.DeviceResults).Count) {
            $result = $testState.DeviceResults[$testState.DeviceIndex]
            $testState.DeviceIndex++
        }
        if ($result -eq 'timeout') { throw [System.TimeoutException]::new('Unable to read data from the transport connection.') }
        return [pscustomobject]@{
            user_code = 'ABCD-EFGH'; device_code = 'device-secret'
            verification_uri = 'https://login.microsoft.com/device'; expires_in = 900
            interval = $testState.Interval; message = 'Enter code ABCD-EFGH'
        }
    }
    if ($Uri -like '*/token') {
        if ($testState.AlwaysPending) { throw (New-OAuthError 'authorization_pending') }
        $result = 'ok'
        if ($testState.TokenIndex -lt @($testState.TokenResults).Count) {
            $result = $testState.TokenResults[$testState.TokenIndex]
            $testState.TokenIndex++
        }
        if ($result -eq 'timeout') { throw [System.TimeoutException]::new('The operation has timed out.') }
        if ($result -eq 'pending') { throw (New-OAuthError 'authorization_pending') }
        if ($result -in @('authorization_declined', 'expired_token', 'invalid_client')) { throw (New-OAuthError $result) }
        return [pscustomobject]@{ access_token = $testState.AccessTokenToReturn }
    }
    throw "Unexpected REST $Method $Uri"
}

function Reset-TokenPlan {
    param([string[]]$Tokens = @('ok'), [string[]]$Devices = @('ok'), [int]$Interval = 1, [switch]$AlwaysPending)
    $testState.TokenResults = $Tokens
    $testState.TokenIndex = 0
    $testState.DeviceResults = $Devices
    $testState.DeviceIndex = 0
    $testState.Interval = $Interval
    $testState.AlwaysPending = [bool]$AlwaysPending
    $testState.Rest.Clear()
}

$oldTmpDir = $env:TMPDIR
$env:TMPDIR = $null
try {
    $tempDirectory = Get-PnpTempDirectory
    Assert-True (-not [string]::IsNullOrWhiteSpace($tempDirectory)) 'GetTempPath still resolves when TMPDIR is null.'
    Assert-True ((Get-PnpProvisionPermissionWarning 'Sites.FullControl.All') -eq 'Provisioning grants full control of all SharePoint sites and Graph Application.Read.All.') 'Full-control warning text is unchanged.'

    Reset-TokenPlan -Devices @('timeout', 'ok') -Tokens @('timeout', 'pending', 'ok')
    $token = Get-PnpDeviceCodeAccessToken -TenantId 'contoso.onmicrosoft.com' -TimeoutSeconds 30
    Assert-True ($token -eq $testState.AccessTokenToReturn) 'Device-code flow returns the access token after transient failures.'
    $deviceCalls = @($testState.Rest | Where-Object { $_.Uri -like '*/devicecode' })
    $tokenCalls = @($testState.Rest | Where-Object { $_.Uri -like '*/token' })
    Assert-True ($deviceCalls.Count -eq 2 -and $tokenCalls.Count -eq 3) 'Hung device-code and token reads are retried.'
    Assert-True (@($testState.Rest | Where-Object { $_.TimeoutSec -lt 1 -or $_.TimeoutSec -gt 60 }).Count -eq 0) 'Each poll uses a bounded HTTP timeout.'
    Assert-True ($deviceCalls[0].Body.client_id -eq '14d82eec-204b-4c2f-b7e8-296a70dab67e') 'Omitted client id uses the Graph PowerShell public client.'
    Assert-True ($deviceCalls[0].Body.scope -match 'Application\.ReadWrite\.All') 'Device-code scope includes application read/write.'

    Reset-TokenPlan -Tokens @('authorization_declined')
    Assert-Throws { Get-PnpDeviceCodeAccessToken -TenantId 'contoso.onmicrosoft.com' -TimeoutSeconds 30 } '*declined*'
    Reset-TokenPlan -Tokens @('expired_token')
    Assert-Throws { Get-PnpDeviceCodeAccessToken -TenantId 'contoso.onmicrosoft.com' -TimeoutSeconds 30 } '*expired*'
    Reset-TokenPlan -Tokens @('invalid_client')
    Assert-Throws { Get-PnpDeviceCodeAccessToken -TenantId 'contoso.onmicrosoft.com' -TimeoutSeconds 30 } '*invalid_client*'
    Reset-TokenPlan -Interval 5 -AlwaysPending
    Assert-Throws { Get-PnpDeviceCodeAccessToken -TenantId 'contoso.onmicrosoft.com' -TimeoutSeconds 1 } '*timed out after 1 seconds*'

    $tenantGuid = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
    $profiles = @(
        @{ name = 'Ciaops'; tenant = 'ciaops365e1.onmicrosoft.com'; organization = 'ciaops365e1.onmicrosoft.com'; defaultDomain = 'ciaops365.com'; siteUrl = 'https://ciaops365.sharepoint.com/sites/one'; appId = 'app-1' }
        @{ name = 'Second'; tenant = 'CIAOPS365E1.ONMICROSOFT.COM'; siteUrl = 'https://ciaops365.sharepoint.com/sites/two'; appId = 'app-2' }
        @{ name = 'Other'; tenant = 'other.onmicrosoft.com'; tenantId = '11111111-2222-3333-4444-555555555555'; siteUrl = 'https://other.sharepoint.com'; appId = 'app-3' }
    )
    $restBefore = $testState.Rest.Count
    $onlyDomain = Select-PnpLinuxProfile -Profiles @(@{ name = 'Only'; tenant = 'ciaops365e1.onmicrosoft.com'; defaultDomain = 'ciaops365.com'; siteUrl = 'https://ciaops365.sharepoint.com'; appId = 'only-app' }) -Tenant 'ciaops365.com' -MapPath 'map.json'
    Assert-True ($onlyDomain.appId -eq 'only-app' -and $testState.Rest.Count -eq $restBefore) 'Stored default domain matches without a directory lookup.'
    Assert-Throws { Select-PnpLinuxProfile -Profiles $profiles -Tenant 'ciaops365.com' -MapPath 'map.json' } '*Specify -ProfileName*'
    $sameDirectory = Select-PnpLinuxProfile -Profiles $profiles -Tenant 'ciaops365.com' -ProfileName 'Second' -MapPath 'map.json'
    Assert-True ($sameDirectory.appId -eq 'app-2') 'A vanity domain also matches another profile for that directory via its onmicrosoft.com domain.'
    $byOrganization = Select-PnpLinuxProfile -Profiles @(@{ name = 'OrgOnly'; tenant = 'unrelated.example'; organization = 'ciaops365e1.onmicrosoft.com'; siteUrl = 'https://ciaops365.sharepoint.com'; appId = 'org-app' }) -Tenant 'CIAOPS365E1.onmicrosoft.com' -MapPath 'map.json'
    Assert-True ($byOrganization.appId -eq 'org-app') 'Organization matches case-insensitively.'
    $byGuid = Select-PnpLinuxProfile -Profiles $profiles -Tenant $tenantGuid.ToUpperInvariant() -ProfileName 'second' -MapPath 'map.json'
    Assert-True ($byGuid.appId -eq 'app-2') 'Tenant GUID matches the stored onmicrosoft.com domain, preferring -ProfileName.'
    $byName = Select-PnpLinuxProfile -Profiles $profiles -Tenant 'ciaops' -MapPath 'map.json'
    Assert-True ($byName.appId -eq 'app-1') 'Profile name matches case-insensitively.'
    $bySite = Select-PnpLinuxProfile -Profiles $profiles -Tenant 'ciaops365e1.onmicrosoft.com' -SiteUrl 'https://ciaops365.sharepoint.com/sites/two' -MapPath 'map.json'
    Assert-True ($bySite.appId -eq 'app-2') 'SiteUrl still narrows an otherwise ambiguous tenant match.'
    Assert-Throws { Select-PnpLinuxProfile -Profiles $profiles -Tenant 'ciaops365e1.onmicrosoft.com' -MapPath 'map.json' } '*Specify -ProfileName*'
    Assert-Throws { Select-PnpLinuxProfile -Profiles $profiles -Tenant $tenantGuid -ProfileName 'missing' -MapPath 'map.json' } '*No profile named ''missing''*'
    Assert-Throws { Select-PnpLinuxProfile -Profiles $profiles -Tenant 'fabrikam.com' -MapPath 'map.json' } '*No matching profile*'
    $other = Select-PnpLinuxProfile -Profiles $profiles -Tenant '11111111-2222-3333-4444-555555555555' -MapPath 'map.json'
    Assert-True ($other.appId -eq 'app-3') 'A stored tenantId matches a GUID without confusing it with another tenant.'

    $testRoot = Join-Path (Get-PnpTempDirectory) ("pnp-cert-tests-" + [guid]::NewGuid())
    $oldData = $env:XDG_DATA_HOME
    $oldConfig = $env:XDG_CONFIG_HOME
    $env:XDG_DATA_HOME = Join-Path $testRoot 'data'
    $env:XDG_CONFIG_HOME = Join-Path $testRoot 'config'
    $password = ConvertTo-SecureString 'pw-test' -AsPlainText -Force
    try {
        Assert-Throws { & $scriptPath } '*exactly one*'
        Assert-Throws { & $scriptPath -GenerateLocalCertificate -UseCertificateAuth } '*exactly one*'
        Assert-Throws { & $scriptPath -GenerateLocalCertificate -ProvisionEntraApp -Headless -Tenant 'contoso.onmicrosoft.com' -SiteUrl 'https://contoso.sharepoint.com' -GeneratedPfxPassword $password } '*interactive administrator*'
        Assert-Throws { & $scriptPath -ProvisionEntraApp -Tenant 'contoso.onmicrosoft.com' -SiteUrl 'https://contoso.sharepoint.com' } '*existing PFX*'
        Assert-Throws { & $scriptPath -SharePointPermission Sites.Manage.All -GenerateLocalCertificate } '*SharePointPermission*'
        Assert-Throws { & $scriptPath -DeviceCodeTimeoutSeconds 0 -GenerateLocalCertificate } '*DeviceCodeTimeoutSeconds*'

        Reset-TokenPlan
        $testState.Requests.Clear()
        $testState.GraphDisconnected = $false
        & $scriptPath -GenerateLocalCertificate -ProvisionEntraApp -Tenant 'contoso.onmicrosoft.com' -SiteUrl 'https://contoso.sharepoint.com' -GeneratedPfxPassword $password -SetupClientId 'client-under-test' -DeviceCodeTimeoutSeconds 30
        Assert-True ($testState.GraphConnected -and $testState.GraphDisconnected) 'Provisioning connects with the acquired token and disconnects.'
        Assert-True ($testState.PlainAccessToken -eq $testState.AccessTokenToReturn) 'Connect-MgGraph receives the device-code access token.'
        Assert-True (@($testState.Rest | Where-Object { $_.Uri -like '*/devicecode' -and $_.Body.client_id -eq 'client-under-test' }).Count -eq 1) '-SetupClientId is the device-code client.'
        $profilePath = Join-Path $env:XDG_CONFIG_HOME 'ciaops/pnp/o365-pnp-cert-auth.json'
        $saved = @(Read-PnpProfiles $profilePath)
        Assert-True ($saved.Count -eq 1) 'Provisioning saves one profile.'
        Assert-True ($saved[0].tenantId -eq 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee') 'Profile stores the directory tenant id from the token.'
        $pfxFiles = @(Get-ChildItem -LiteralPath (Join-Path $env:XDG_DATA_HOME 'ciaops/pnp') -Filter '*.pfx' -File)
        Assert-True ($pfxFiles.Count -eq 1) 'Generation creates one PFX.'
        Assert-True ([int][IO.File]::GetUnixFileMode($pfxFiles[0].FullName) -eq 384) 'PFX permissions are 600.'
        $created = @($testState.Requests | Where-Object { $_.Method -eq 'POST' -and $_.Path -eq 'applications' })[0].Body
        $createdSharePoint = @($created.requiredResourceAccess | Where-Object { $_.resourceAppId -eq '00000003-0000-0ff1-ce00-000000000000' })
        Assert-True (@($createdSharePoint[0].resourceAccess)[0].id -eq 'sp-Sites.FullControl.All') 'Default provisioning grants Sites.FullControl.All.'
        $createdGraph = @($created.requiredResourceAccess | Where-Object { $_.resourceAppId -eq '00000003-0000-0000-c000-000000000000' })
        Assert-True (@($createdGraph[0].resourceAccess)[0].id -eq 'graph-app-read') 'Provisioning still grants Graph Application.Read.All.'

        $cerPath = [IO.Path]::ChangeExtension($pfxFiles[0].FullName, '.cer')
        foreach ($permission in @('Sites.Read.All', 'Sites.Selected')) {
            $testState.Requests.Clear()
            New-PnpLinuxEntraApp -DisplayName "O365-$permission" -CerPath $cerPath -SharePointPermission $permission | Out-Null
            $posted = @($testState.Requests | Where-Object { $_.Method -eq 'POST' -and $_.Path -eq 'applications' })[0].Body
            $sharePointAccess = @(@($posted.requiredResourceAccess | Where-Object { $_.resourceAppId -eq '00000003-0000-0ff1-ce00-000000000000' })[0].resourceAccess)
            Assert-True ($sharePointAccess[0].id -eq "sp-$permission") "Provisioning can grant $permission."
        }

        Assert-Throws { & $scriptPath -GenerateLocalCertificate -ProvisionEntraApp -Tenant 'contoso.onmicrosoft.com' -SiteUrl 'https://contoso.sharepoint.com' -CertificatePath $pfxFiles[0].FullName -GeneratedPfxPassword $password } '*cannot both*'
        $testState.Requests.Clear()
        Reset-TokenPlan
        & $scriptPath -ProvisionEntraApp -Tenant 'contoso.onmicrosoft.com' -SiteUrl 'https://contoso.sharepoint.com' -CertificatePath $pfxFiles[0].FullName -CertificatePassword $password -SharePointPermission 'Sites.Read.All' -DeviceCodeTimeoutSeconds 30
        Assert-True (@(Get-ChildItem -LiteralPath (Join-Path $env:XDG_DATA_HOME 'ciaops/pnp') -Filter '*.pfx' -File).Count -eq 1) 'Reusing a PFX does not mint another certificate.'
        $reused = @($testState.Requests | Where-Object { $_.Method -eq 'POST' -and $_.Path -eq 'applications' })[0].Body
        Assert-True (@(@($reused.requiredResourceAccess | Where-Object { $_.resourceAppId -eq '00000003-0000-0ff1-ce00-000000000000' })[0].resourceAccess)[0].id -eq 'sp-Sites.Read.All') 'Reused-certificate provisioning uses the selected permission.'

        Reset-TokenPlan -AlwaysPending
        Assert-Throws { & $scriptPath -ProvisionEntraApp -Tenant 'contoso.onmicrosoft.com' -SiteUrl 'https://contoso.sharepoint.com' -CertificateThumbprint $pfxFiles[0].BaseName.ToLowerInvariant() -CertificatePassword $password -DeviceCodeTimeoutSeconds 1 } '*timed out after 1 seconds*'
        Assert-True (@(Get-ChildItem -LiteralPath (Join-Path $env:XDG_DATA_HOME 'ciaops/pnp') -Filter '*.pfx' -File).Count -eq 1) 'Thumbprint lookup finds the existing PFX.'

        $aliasMap = Join-Path $testRoot 'alias.json'
        Save-PnpLinuxProfile $aliasMap @{
            name = 'Ciaops'; tenant = 'ciaops365e1.onmicrosoft.com'; siteUrl = 'https://ciaops365.sharepoint.com/sites/team'
            appId = 'ciaops-app'; certificateThumbprint = $saved[0].certificateThumbprint; certificateFilePath = $pfxFiles[0].FullName
        }
        Save-PnpLinuxProfile $aliasMap @{
            name = 'Other'; tenant = 'other.onmicrosoft.com'; siteUrl = 'https://other.sharepoint.com'
            appId = 'other-app'; certificateThumbprint = $saved[0].certificateThumbprint; certificateFilePath = $pfxFiles[0].FullName
        }
        & $scriptPath -UseCertificateAuth -CertificateMapPath $aliasMap -Tenant $tenantGuid.ToUpperInvariant() -Headless -CertificatePassword $password
        Assert-True ($testState.Connection.AppId -eq 'ciaops-app' -and $testState.Connection.Tenant -eq $tenantGuid.ToUpperInvariant()) 'Connecting with the tenant GUID selects the onmicrosoft.com profile.'
        & $scriptPath -UseCertificateAuth -CertificateMapPath $aliasMap -Tenant 'CIAOPS365.com' -ProfileName 'ciaops' -Headless -CertificatePassword $password
        Assert-True ($testState.Connection.Url -eq 'https://ciaops365.sharepoint.com/sites/team') 'A vanity domain selects the same tenant profile.'

        $ambiguousMap = Join-Path $testRoot 'ambiguous.json'
        Save-PnpLinuxProfile $ambiguousMap @{
            name = 'One'; tenant = 'contoso.onmicrosoft.com'; siteUrl = 'https://contoso.sharepoint.com/sites/one'
            appId = 'one-app'; certificateThumbprint = $saved[0].certificateThumbprint; certificateFilePath = $pfxFiles[0].FullName
        }
        Save-PnpLinuxProfile $ambiguousMap @{
            name = 'Two'; tenant = 'contoso.onmicrosoft.com'; siteUrl = 'https://contoso.sharepoint.com/sites/two'
            appId = 'two-app'; certificateThumbprint = $saved[0].certificateThumbprint; certificateFilePath = $pfxFiles[0].FullName
        }
        Assert-Throws { & $scriptPath -UseCertificateAuth -CertificateMapPath $ambiguousMap -Tenant 'Contoso.OnMicrosoft.com' -Headless -CertificatePassword $password } '*Specify -ProfileName*'
        & $scriptPath -UseCertificateAuth -CertificateMapPath $ambiguousMap -Tenant 'contoso.onmicrosoft.com' -ProfileName 'two' -Headless -CertificatePassword $password
        Assert-True ($testState.Connection.AppId -eq 'two-app') '-ProfileName picks one of several case-insensitive tenant matches.'

        & $scriptPath -UseCertificateAuth -Tenant 'explicit.example' -SiteUrl 'https://explicit.sharepoint.com/sites/a' -AppId 'explicit-app' -CertificateFilePath $pfxFiles[0].FullName -Headless -CertificatePassword $password
        Assert-True ($testState.Connection.AppId -eq 'explicit-app') 'Explicit connection parameters bypass profile lookup.'

        $help = Get-Help $scriptPath -Full
        Assert-True ($help.Parameters.parameter.Name -contains 'DeviceCodeTimeoutSeconds') 'Help documents -DeviceCodeTimeoutSeconds.'
        Assert-True ($help.Parameters.parameter.Name -contains 'SharePointPermission') 'Help documents -SharePointPermission.'
        Assert-True (($help.Parameters.parameter | Where-Object Name -eq 'CertificateFilePath').Description.Text -join ' ' -match 'CertificatePath') 'Help documents -CertificatePath.'
    }
    finally {
        $env:XDG_DATA_HOME = $oldData
        $env:XDG_CONFIG_HOME = $oldConfig
        if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
    }
    Write-Host 'All offline PnP Linux certificate regression tests passed.'
}
finally { $env:TMPDIR = $oldTmpDir }
