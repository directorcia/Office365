<#
.SYNOPSIS
    Generate file-based certificates and connect to SharePoint Online on Linux.
.DESCRIPTION
    Linux counterpart of o365-connect-pnp-cert.ps1, based on
    o365-connect-exo-cert-linux.ps1. Select exactly one mode:
    -GenerateLocalCertificate creates an RSA-2048/SHA256 certificate, a public
    DER .cer, and an encrypted private-key .pfx using .NET (no OpenSSL needed).
    -UseCertificateAuth connects with PnP.PowerShell using an encrypted PFX.

    Certificates default to $XDG_DATA_HOME/ciaops/pnp (or ~/.local/share/ciaops/pnp).
    Profiles default to $XDG_CONFIG_HOME/ciaops/pnp/o365-pnp-cert-auth.json
    (or ~/.config/ciaops/pnp/o365-pnp-cert-auth.json). Dedicated storage directories
    are restricted to mode 700 and files to mode 600, independent of the umask.
    Custom output/map directories must also be dedicated to this script: their
    permissions are restricted to 700. Symbolic links are not accepted there.

    -ProvisionEntraApp creates/reuses an Entra app, uploads the public certificate,
    grants SharePoint Sites.FullControl.All and Graph Application.Read.All, and
    saves a connection profile. These match the Windows script's permissions;
    they grant full control of all SharePoint sites and read access to directory
    applications. Provisioning requires an administrator able to consent to these
    application permissions and uses Microsoft Graph device-code sign-in.

    Requires Linux, PowerShell 7.4.6+, and PnP.PowerShell 3.0.0+ for connection.
    Missing modules are installed for CurrentUser without sudo. Generation alone
    needs no modules; provisioning only needs Microsoft.Graph.Authentication.
    Run in the same PowerShell session as subsequent PnP cmdlets. No child process
    is launched and the connection is not disconnected on successful completion.
    Use a fresh pwsh session if previously loaded modules cause assembly conflicts.
.PARAMETER GenerateLocalCertificate
    Generate a certificate. Always exports an encrypted PFX on Linux.
.PARAMETER UseCertificateAuth
    Connect using Tenant/SiteUrl/AppId/CertificateFilePath or a saved profile.
.PARAMETER GeneratedCertSubject
    Certificate common name and default app name. Defaults to O365-PNP-AppAuth.
.PARAMETER GeneratedCertYearsValid
    Certificate lifetime in years (1-100). Defaults to 2.
.PARAMETER GeneratedCertOutputPath
    Dedicated certificate directory. Defaults to the XDG data location above.
.PARAMETER GeneratedPfxPassword
    Nonempty SecureString for the generated PFX; prompted for if omitted.
.PARAMETER CertificateFilePath
    Existing encrypted PFX. Overrides the profile path. Windows thumbprint-only
    profiles require this parameter or an added certificateFilePath property.
.PARAMETER CertificatePassword
    PFX password for connection; prompted for if omitted. For unattended use,
    pass a SecureString from a secret vault with -noprompt. Passwords are never
    written to profiles. Linux ConvertFrom-SecureString is not DPAPI encryption.
.PARAMETER CertificatePasswordEnvironmentVariable
    Name of an environment variable containing the PFX password, injected by a
    scheduler or secret manager. Used for generation/connection when the relevant
    SecureString parameter is omitted. Never printed or saved. Environment values
    may be accessible to processes running as the same user; prefer a secret vault.
.PARAMETER CertificateThumbprint
    Optional expected thumbprint; also read from profiles and checked against PFX.
.PARAMETER CertificateMapPath
    JSON profile map; accepts a bare array or an object containing profiles.
.PARAMETER Tenant
    Tenant ID or domain for provisioning and connection, or saved profile filter.
.PARAMETER SiteUrl
    HTTPS SharePoint site URL. Required for provisioning; also a profile filter.
.PARAMETER AppId
    Application client ID for explicit connections.
.PARAMETER ProfileName
    Saved profile name to select. Defaults to AppDisplayName during provisioning.
.PARAMETER ProvisionEntraApp
    Generate mode only: interactive one-time app setup and admin consent.
.PARAMETER AppDisplayName
    App name; defaults to GeneratedCertSubject. Duplicate names cause an error.
.PARAMETER SetupClientId
    Optional tenant-approved public client ID for Microsoft Graph device sign-in.
    If omitted, uses the Microsoft Graph PowerShell client.
.PARAMETER noprompt
    Disable prompts. Alias: -Headless. Requires a password parameter or named
    password environment variable and preinstalled modules. Skips update checks.
    Cannot be combined with ProvisionEntraApp, which requires admin sign-in.
.PARAMETER noupdate
    Skip the PnP module Gallery version check. Updates otherwise need confirmation.
.PARAMETER enableLog
    Save a timestamped transcript in the private certificate data directory.
    Provisioning logs may contain short-lived device codes; protect these logs.
.EXAMPLE
    ./o365-connect-pnp-cert-linux.ps1 -GenerateLocalCertificate -ProvisionEntraApp -Tenant contoso.onmicrosoft.com -SiteUrl https://contoso.sharepoint.com
    Create encrypted certificate files, provision the app and save a profile.
.EXAMPLE
    ./o365-connect-pnp-cert-linux.ps1 -UseCertificateAuth -Tenant contoso.onmicrosoft.com
    Prompt for the saved profile's PFX password and connect in this session.
.EXAMPLE
    ./o365-connect-pnp-cert-linux.ps1 -UseCertificateAuth -Tenant contoso.onmicrosoft.com -SiteUrl https://contoso.sharepoint.com -AppId 11111111-2222-3333-4444-555555555555 -CertificateFilePath /home/admin/auth.pfx -CertificatePassword $vaultPassword -noprompt -noupdate
    Connect unattended using an existing app and exported PFX (chmod 600 first).
.EXAMPLE
    ./o365-connect-pnp-cert-linux.ps1 -UseCertificateAuth -ProfileName Contoso -Headless -CertificatePasswordEnvironmentVariable PNP_PFX_PASSWORD
    Connect using a saved profile and a password injected into the process.
.NOTES
    CIAOPS - Script provided as is. Use at own risk. No guarantees or warranty provided.
.LINK
    https://github.com/directorcia/Office365/wiki/SharePoint-Online-certificate-authentication-on-Linux
    https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html
#>
#requires -Version 7.4.6
[CmdletBinding()]
param(
    [Alias('Headless')][switch]$noprompt,
    [switch]$noupdate,
    [switch]$enableLog,
    [switch]$GenerateLocalCertificate,
    [switch]$UseCertificateAuth,
    [string]$GeneratedCertSubject = 'O365-PNP-AppAuth',
    [ValidateRange(1, 100)][int]$GeneratedCertYearsValid = 2,
    [string]$GeneratedCertOutputPath,
    [securestring]$GeneratedPfxPassword,
    [string]$CertificateFilePath,
    [securestring]$CertificatePassword,
    [string]$CertificatePasswordEnvironmentVariable,
    [string]$CertificateThumbprint,
    [string]$CertificateMapPath,
    [string]$Tenant,
    [string]$SiteUrl,
    [string]$AppId,
    [string]$ProfileName,
    [switch]$ProvisionEntraApp,
    [string]$AppDisplayName,
    [string]$SetupClientId
)

function Initialize-PnpPrivateDirectory {
    param([Parameter(Mandatory)][string]$Path)
    $fullPath = [IO.Path]::GetFullPath($Path)
    $ancestor = $fullPath
    while ($ancestor) {
        if ([IO.File]::Exists($ancestor) -or [IO.Directory]::Exists($ancestor)) {
            if ([IO.File]::GetAttributes($ancestor) -band [IO.FileAttributes]::ReparsePoint) {
                throw "Symbolic links are not supported in storage paths: $ancestor"
            }
        }
        $ancestor = [IO.Path]::GetDirectoryName($ancestor)
    }
    [IO.Directory]::CreateDirectory($fullPath, [IO.UnixFileMode]::UserRead -bor
        [IO.UnixFileMode]::UserWrite -bor [IO.UnixFileMode]::UserExecute) | Out-Null
    [IO.File]::SetUnixFileMode($fullPath, [IO.UnixFileMode]::UserRead -bor
        [IO.UnixFileMode]::UserWrite -bor [IO.UnixFileMode]::UserExecute)
    return $fullPath
}

function Write-PnpPrivateFile {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][AllowEmptyCollection()][byte[]]$Bytes)
    $options = [IO.FileStreamOptions]::new()
    $options.Mode = [IO.FileMode]::CreateNew
    $options.Access = [IO.FileAccess]::Write
    $options.Share = [IO.FileShare]::None
    $options.UnixCreateMode = [IO.UnixFileMode]::UserRead -bor [IO.UnixFileMode]::UserWrite
    $stream = [IO.FileStream]::new($Path, $options)
    try { $stream.Write($Bytes, 0, $Bytes.Length) }
    finally { $stream.Dispose() }
}

function Read-PnpProfiles {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $map = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -ErrorAction Stop
    if ($map -is [System.Collections.IDictionary] -and $map.Contains('profiles')) { return $map.profiles }
    if ($map -is [array]) { return $map }
    throw "Invalid profile map: $Path. Expected an array or an object containing profiles."
}

function Save-PnpLinuxProfile {
    param([string]$Path, [hashtable]$Entry)
    $directory = Initialize-PnpPrivateDirectory -Path (Split-Path -Parent $Path)
    $lockOptions = [IO.FileStreamOptions]::new()
    $lockOptions.Mode = [IO.FileMode]::OpenOrCreate
    $lockOptions.Access = [IO.FileAccess]::ReadWrite
    $lockOptions.Share = [IO.FileShare]::None
    $lockOptions.UnixCreateMode = [IO.UnixFileMode]::UserRead -bor [IO.UnixFileMode]::UserWrite
    $lockStream = [IO.FileStream]::new((Join-Path $directory '.profiles.lock'), $lockOptions)
    $temporaryPath = Join-Path $directory ([guid]::NewGuid().ToString() + '.tmp')
    try {
        $profiles = @(Read-PnpProfiles -Path $Path | Where-Object {
            -not ($_.appId -eq $Entry.appId -and $_.name -eq $Entry.name)
        }) + @($Entry)
        $json = @{ profiles = $profiles } | ConvertTo-Json -Depth 10
        Write-PnpPrivateFile -Path $temporaryPath -Bytes ([Text.Encoding]::UTF8.GetBytes($json))
        [IO.File]::Move($temporaryPath, $Path, $true)
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force }
        $lockStream.Dispose()
    }
}

function Get-PnpPfxPassword {
    param([securestring]$Password, [switch]$NoPrompt, [string]$EnvironmentVariable)
    if ($null -eq $Password -and $EnvironmentVariable) {
        $environmentPassword = [Environment]::GetEnvironmentVariable($EnvironmentVariable, 'Process')
        if ([string]::IsNullOrEmpty($environmentPassword)) {
            throw "PFX password environment variable '$EnvironmentVariable' is missing or empty."
        }
        try { $Password = ConvertTo-SecureString -String $environmentPassword -AsPlainText -Force }
        finally { $environmentPassword = $null }
    }
    if ($null -eq $Password) {
        if ($NoPrompt) { throw 'Headless mode requires a SecureString PFX password or -CertificatePasswordEnvironmentVariable.' }
        $Password = Read-Host -Prompt 'PFX password' -AsSecureString
    }
    if ($Password.Length -eq 0) { throw 'An empty PFX password is not permitted.' }
    return $Password
}

function New-PnpLinuxCertificate {
    param([string]$SubjectName, [int]$YearsValid, [string]$OutputPath, [securestring]$Password)
    $directory = Initialize-PnpPrivateDirectory -Path $OutputPath
    $rsa = [Security.Cryptography.RSA]::Create(2048)
    $certificate = $null
    try {
        $subject = [Security.Cryptography.X509Certificates.X500DistinguishedName]::new('CN="' + $SubjectName.Replace('"', '""') + '"')
        $request = [Security.Cryptography.X509Certificates.CertificateRequest]::new(
            $subject, $rsa, [Security.Cryptography.HashAlgorithmName]::SHA256,
            [Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $request.CertificateExtensions.Add([Security.Cryptography.X509Certificates.X509KeyUsageExtension]::new(
            [Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature, $true))
        $certificate = $request.CreateSelfSigned([DateTimeOffset]::UtcNow.AddMinutes(-5), [DateTimeOffset]::UtcNow.AddYears($YearsValid))
        $basePath = Join-Path $directory $certificate.Thumbprint
        $pfxPath = "$basePath.pfx"
        $cerPath = "$basePath.cer"
        Write-PnpPrivateFile -Path $pfxPath -Bytes ($certificate.Export([Security.Cryptography.X509Certificates.X509ContentType]::Pfx, $Password))
        Write-PnpPrivateFile -Path $cerPath -Bytes ($certificate.Export([Security.Cryptography.X509Certificates.X509ContentType]::Cert))
        return @{ Thumbprint = $certificate.Thumbprint; PfxPath = $pfxPath; CerPath = $cerPath; NotAfter = $certificate.NotAfter }
    }
    finally {
        if ($null -ne $certificate) { $certificate.Dispose() }
        $rsa.Dispose()
    }
}

function Import-PnpLinuxPfx {
    param([string]$Path, [securestring]$Password)
    $fullPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    if ([IO.File]::GetAttributes($fullPath) -band [IO.FileAttributes]::ReparsePoint) {
        throw "PFX must not be a symbolic link: $fullPath"
    }
    $mode = [IO.File]::GetUnixFileMode($fullPath)
    $allowedMode = [IO.UnixFileMode]::UserRead -bor [IO.UnixFileMode]::UserWrite
    if (($mode -band (-bnot $allowedMode)) -ne 0) {
        throw "PFX permissions must be owner-only. Run: chmod 600 -- '$($fullPath.Replace("'", "'\''"))'"
    }
    return [Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $fullPath, $Password, [Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet)
}

function Import-PnpLinuxModule {
    param([string]$Name, [version]$MinimumVersion, [switch]$CheckUpdate)
    $loaded = Get-Module -Name $Name | Select-Object -First 1
    if ($null -ne $loaded) {
        if ($loaded.Version -lt $MinimumVersion) { throw "$Name $MinimumVersion or later is required. Start a fresh pwsh session with the newer module installed." }
        return
    }
    $available = Get-Module -ListAvailable -Name $Name | Where-Object { $_.Version -ge $MinimumVersion } |
        Sort-Object Version -Descending | Select-Object -First 1
    if ($null -eq $available) {
        if ($noprompt) { throw "Install the required module first: Install-Module $Name -Scope CurrentUser -MinimumVersion $MinimumVersion" }
        Install-Module -Name $Name -MinimumVersion $MinimumVersion -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
    }
    elseif ($CheckUpdate -and -not $noprompt) {
        try {
            $latest = Find-Module -Name $Name -Repository PSGallery -ErrorAction Stop
            if ([version]$latest.Version -gt $available.Version -and (Read-Host "Install $Name $($latest.Version)? [y/N]") -eq 'y') {
                Install-Module -Name $Name -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop
            }
        }
        catch { Write-Warning "Module update check failed: $($_.Exception.Message)" }
    }
    Import-Module -Name $Name -MinimumVersion $MinimumVersion -ErrorAction Stop
}

function Invoke-PnpLinuxGraph {
    param([string]$Method = 'GET', [string]$Path, [hashtable]$Body)
    $parameters = @{ Method = $Method; Uri = "https://graph.microsoft.com/v1.0/$Path"; OutputType = 'Hashtable'; ErrorAction = 'Stop' }
    if ($null -ne $Body) {
        $parameters.Body = $Body | ConvertTo-Json -Depth 20 -Compress
        $parameters.ContentType = 'application/json'
    }
    Invoke-MgGraphRequest @parameters
}

function New-PnpLinuxEntraApp {
    param([string]$DisplayName, [string]$CerPath)
    $targets = @(
        @{ AppId = '00000003-0000-0ff1-ce00-000000000000'; RoleName = 'Sites.FullControl.All' },
        @{ AppId = '00000003-0000-0000-c000-000000000000'; RoleName = 'Application.Read.All' }
    )
    foreach ($target in $targets) {
        $resourceFilter = [uri]::EscapeDataString("appId eq '$($target.AppId)'")
        $resources = Invoke-PnpLinuxGraph -Path "servicePrincipals?`$filter=$resourceFilter"
        if (@($resources.value).Count -ne 1) { throw "Unable to resolve resource $($target.AppId) in the provisioning tenant." }
        $target.Resource = @($resources.value)[0]
        $roles = @($target.Resource.appRoles | Where-Object {
            $_.value -eq $target.RoleName -and $_.isEnabled -and $_.allowedMemberTypes -contains 'Application'
        })
        if ($roles.Count -ne 1) { throw "Unable to resolve $($target.RoleName) in the provisioning tenant." }
        $target.RoleId = $roles[0].id
    }
    $appFilter = [uri]::EscapeDataString("displayName eq '$($DisplayName.Replace("'", "''"))'")
    $apps = Invoke-PnpLinuxGraph -Path "applications?`$filter=$appFilter"
    if (@($apps.value).Count -gt 1) { throw "Multiple applications named '$DisplayName'. Use a unique -AppDisplayName." }
    $application = if (@($apps.value).Count -eq 1) { @($apps.value)[0] } else { $null }
    if ($null -eq $application) {
        $application = Invoke-PnpLinuxGraph -Method POST -Path 'applications' -Body @{
            displayName = $DisplayName; signInAudience = 'AzureADMyOrg'
            requiredResourceAccess = @($targets | ForEach-Object {
                @{ resourceAppId = $_.AppId; resourceAccess = @(@{ id = $_.RoleId; type = 'Role' }) }
            })
        }
    }
    $publicCertificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($CerPath)
    try {
        $existing = Invoke-PnpLinuxGraph -Path "applications/$($application.id)?`$select=keyCredentials,requiredResourceAccess"
        $keyIdentifier = [Convert]::ToBase64String($publicCertificate.GetCertHash())
        $keys = @($existing.keyCredentials | Where-Object { $_.customKeyIdentifier -ne $keyIdentifier })
        foreach ($existingKey in $keys) {
            if ([string]::IsNullOrWhiteSpace($existingKey.key)) {
                throw "Microsoft Graph did not return public key material for existing credential '$($existingKey.keyId)'. Refusing to replace keyCredentials or remove existing certificates. Retrieve the application with `$select=keyCredentials before retrying."
            }
        }
        $keys += @(@{
            type = 'AsymmetricX509Cert'; usage = 'Verify'; displayName = "PNP-$($publicCertificate.Thumbprint)"
            key = [Convert]::ToBase64String($publicCertificate.RawData); customKeyIdentifier = $keyIdentifier
            startDateTime = $publicCertificate.NotBefore.ToUniversalTime().ToString('o')
            endDateTime = $publicCertificate.NotAfter.ToUniversalTime().ToString('o')
        })
        $permissions = @($existing.requiredResourceAccess)
        foreach ($target in $targets) {
            $permission = $permissions | Where-Object { $_.resourceAppId -eq $target.AppId } | Select-Object -First 1
            if ($null -eq $permission) {
                $permissions += @(@{ resourceAppId = $target.AppId; resourceAccess = @(@{ id = $target.RoleId; type = 'Role' }) })
            }
            elseif (-not ($permission.resourceAccess | Where-Object { $_.id -eq $target.RoleId -and $_.type -eq 'Role' })) {
                $permission.resourceAccess = @($permission.resourceAccess) + @(@{ id = $target.RoleId; type = 'Role' })
            }
        }
        Invoke-PnpLinuxGraph -Method PATCH -Path "applications/$($application.id)" -Body @{
            keyCredentials = $keys; requiredResourceAccess = $permissions
        } | Out-Null
    }
    finally { $publicCertificate.Dispose() }
    $principalFilter = [uri]::EscapeDataString("appId eq '$($application.appId)'")
    $principals = Invoke-PnpLinuxGraph -Path "servicePrincipals?`$filter=$principalFilter"
    $principal = if (@($principals.value).Count -eq 1) { @($principals.value)[0] } else { $null }
    if ($null -eq $principal) {
        $principal = Invoke-PnpLinuxGraph -Method POST -Path 'servicePrincipals' -Body @{ appId = $application.appId }
    }
    $assignments = Invoke-PnpLinuxGraph -Path "servicePrincipals/$($principal.id)/appRoleAssignments"
    foreach ($target in $targets) {
        if (-not ($assignments.value | Where-Object { $_.resourceId -eq $target.Resource.id -and $_.appRoleId -eq $target.RoleId })) {
            Invoke-PnpLinuxGraph -Method POST -Path "servicePrincipals/$($principal.id)/appRoleAssignments" -Body @{
                principalId = $principal.id; resourceId = $target.Resource.id; appRoleId = $target.RoleId
            } | Out-Null
        }
    }
    return @{ AppId = $application.appId }
}

if (-not $IsLinux) { throw 'This script requires Linux and PowerShell 7.4.6 or later. Use o365-connect-pnp-cert.ps1 on Windows.' }
if ($GenerateLocalCertificate -eq $UseCertificateAuth) { throw 'Specify exactly one of -GenerateLocalCertificate or -UseCertificateAuth.' }
if ($ProvisionEntraApp -and (-not $GenerateLocalCertificate -or $noprompt)) {
    throw '-ProvisionEntraApp requires -GenerateLocalCertificate and interactive administrator sign-in (no -noprompt).'
}
if ($ProvisionEntraApp -and (-not $Tenant -or -not $SiteUrl)) { throw '-ProvisionEntraApp requires -Tenant and -SiteUrl.' }
if ($SiteUrl -and ($SiteUrl -notmatch '^https://' -or -not [uri]::IsWellFormedUriString($SiteUrl, [UriKind]::Absolute))) {
    throw '-SiteUrl must be an absolute HTTPS SharePoint site URL.'
}
$dataRoot = if ($env:XDG_DATA_HOME -and [IO.Path]::IsPathRooted($env:XDG_DATA_HOME)) { $env:XDG_DATA_HOME } else { Join-Path $HOME '.local/share' }
$configRoot = if ($env:XDG_CONFIG_HOME -and [IO.Path]::IsPathRooted($env:XDG_CONFIG_HOME)) { $env:XDG_CONFIG_HOME } else { Join-Path $HOME '.config' }
$dataDirectory = Join-Path $dataRoot 'ciaops/pnp'
if (-not $GeneratedCertOutputPath) { $GeneratedCertOutputPath = $dataDirectory }
if (-not $CertificateMapPath) { $CertificateMapPath = Join-Path $configRoot 'ciaops/pnp/o365-pnp-cert-auth.json' }
$GeneratedCertOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($GeneratedCertOutputPath)
$CertificateMapPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CertificateMapPath)
$transcribing = $false
$localCertificate = $null
try {
    if ($enableLog) {
        $logDirectory = Initialize-PnpPrivateDirectory -Path $dataDirectory
        $logPath = Join-Path $logDirectory ("o365-connect-pnp-{0}.txt" -f [guid]::NewGuid())
        Write-PnpPrivateFile -Path $logPath -Bytes ([byte[]]::new(0))
        Start-Transcript -LiteralPath $logPath -Append -ErrorAction Stop | Out-Null
        $transcribing = $true
    }
    if ($GenerateLocalCertificate) {
        $password = Get-PnpPfxPassword -Password $GeneratedPfxPassword -NoPrompt:$noprompt -EnvironmentVariable $CertificatePasswordEnvironmentVariable
        $generated = New-PnpLinuxCertificate -SubjectName $GeneratedCertSubject -YearsValid $GeneratedCertYearsValid -OutputPath $GeneratedCertOutputPath -Password $password
        Write-Host "Public certificate: $($generated.CerPath)"
        Write-Host "Encrypted private key: $($generated.PfxPath)"
        Write-Host "Thumbprint: $($generated.Thumbprint); expires: $($generated.NotAfter.ToString('yyyy-MM-dd'))"
        if (-not $ProvisionEntraApp) { return }
        Write-Warning 'Provisioning grants full control of all SharePoint sites and Graph Application.Read.All.'
        Import-PnpLinuxModule -Name Microsoft.Graph.Authentication -MinimumVersion 2.0.0
        $graphParameters = @{
            TenantId = $Tenant; UseDeviceAuthentication = $true; ContextScope = 'Process'; NoWelcome = $true
            Scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All'); ErrorAction = 'Stop'
        }
        if ($SetupClientId) { $graphParameters.ClientId = $SetupClientId }
        Connect-MgGraph @graphParameters
        try {
            $displayName = if ($AppDisplayName) { $AppDisplayName } else { $GeneratedCertSubject }
            $provisioned = New-PnpLinuxEntraApp -DisplayName $displayName -CerPath $generated.CerPath
            $entryName = if ($ProfileName) { $ProfileName } else { $displayName }
            Save-PnpLinuxProfile -Path $CertificateMapPath -Entry @{
                name = $entryName; tenant = $Tenant; siteUrl = $SiteUrl; appId = $provisioned.AppId
                certificateThumbprint = $generated.Thumbprint; certificateFilePath = $generated.PfxPath
            }
            Write-Host "Provisioned app: $($provisioned.AppId). Profile saved: $CertificateMapPath"
            Write-Warning 'New app permissions may take time to replicate. Connect from a fresh pwsh session.'
        }
        finally { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
        return
    }

    if ($ProfileName -or -not ($Tenant -and $SiteUrl -and $AppId -and $CertificateFilePath)) {
        $profiles = @(Read-PnpProfiles -Path $CertificateMapPath | Where-Object {
            (-not $ProfileName -or $_.name -eq $ProfileName) -and
            (-not $Tenant -or $_.tenant -eq $Tenant) -and
            (-not $SiteUrl -or $_.siteUrl -eq $SiteUrl)
        })
        if ($profiles.Count -ne 1) { throw "Expected one matching profile in '$CertificateMapPath', found $($profiles.Count). Specify -ProfileName or explicit connection parameters." }
        $profile = $profiles[0]
        if (-not $Tenant) { $Tenant = $profile.tenant }
        if (-not $SiteUrl) { $SiteUrl = $profile.siteUrl }
        if (-not $AppId) { $AppId = $profile.appId }
        if (-not $CertificateFilePath) {
            $CertificateFilePath = $profile.certificateFilePath
            if ($CertificateFilePath -and -not [IO.Path]::IsPathRooted($CertificateFilePath)) {
                $CertificateFilePath = Join-Path (Split-Path -Parent $CertificateMapPath) $CertificateFilePath
            }
        }
        if (-not $CertificateThumbprint) { $CertificateThumbprint = $profile.certificateThumbprint }
    }
    if (-not ($Tenant -and $SiteUrl -and $AppId -and $CertificateFilePath)) {
        throw 'Connection requires Tenant, SiteUrl, AppId and CertificateFilePath. Windows thumbprint-only profiles need a certificateFilePath pointing to an exported PFX.'
    }
    if ($SiteUrl -notmatch '^https://' -or -not [uri]::IsWellFormedUriString($SiteUrl, [UriKind]::Absolute)) {
        throw 'The profile SiteUrl must be an absolute HTTPS SharePoint site URL.'
    }
    $CertificateFilePath = (Resolve-Path -LiteralPath $CertificateFilePath -ErrorAction Stop).ProviderPath
    $password = Get-PnpPfxPassword -Password $CertificatePassword -NoPrompt:$noprompt -EnvironmentVariable $CertificatePasswordEnvironmentVariable
    $localCertificate = Import-PnpLinuxPfx -Path $CertificateFilePath -Password $password
    if (-not $localCertificate.HasPrivateKey) { throw 'The certificate has no private key. Supply a PFX, not a public CER.' }
    if ($CertificateThumbprint -and $localCertificate.Thumbprint -ne ($CertificateThumbprint -replace '[^0-9a-fA-F]', '')) {
        throw 'The PFX thumbprint does not match the expected certificate thumbprint.'
    }
    if ($localCertificate.NotBefore -gt (Get-Date) -or $localCertificate.NotAfter -le (Get-Date)) { throw 'The certificate is not currently valid.' }
    if ($localCertificate.NotAfter -lt (Get-Date).AddDays(30)) { Write-Warning 'The certificate expires in fewer than 30 days.' }
    Import-PnpLinuxModule -Name PnP.PowerShell -MinimumVersion 3.0.0 -CheckUpdate:(-not $noupdate)
    Connect-PnPOnline -Url $SiteUrl -Tenant $Tenant -ClientId $AppId -CertificatePath $CertificateFilePath -CertificatePassword $password -ValidateConnection -ErrorAction Stop | Out-Null
    Write-Host "Connected to $SiteUrl using app $AppId and certificate $($localCertificate.Thumbprint)."
}
finally {
    if ($null -ne $localCertificate) { $localCertificate.Dispose() }
    if ($transcribing) { Stop-Transcript | Out-Null }
}