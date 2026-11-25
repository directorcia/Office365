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
    grants the selected SharePoint application permission (Sites.FullControl.All
    by default) and Graph Application.Read.All, and saves a connection profile.
    Sites.FullControl.All grants full control of all SharePoint sites.
    Sites.Read.All grants read access to all sites. Sites.Selected grants no site
    access until sites are assigned. Graph Application.Read.All reads directory
    applications. Provisioning requires an administrator able to consent to these
    application permissions. It signs in with the OAuth device-code flow and waits
    -DeviceCodeTimeoutSeconds (default 15 minutes) so an administrator can finish
    a relayed code. Polls retry transient network read timeouts until that window
    ends. The Microsoft Graph PowerShell SDK cancels its own device-code sign-in
    after 120 seconds, so this script acquires the token itself and passes it to
    Connect-MgGraph -AccessToken, which keeps the session in process scope.
    Pair -ProvisionEntraApp with -GenerateLocalCertificate, or reuse a PFX already
    on disk via -CertificateFilePath/-CertificatePath or -CertificateThumbprint.

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
    Existing encrypted PFX. Alias: -CertificatePath. Overrides the profile path
    when connecting. With -ProvisionEntraApp and without -GenerateLocalCertificate,
    reuses this PFX instead of minting a new certificate. Windows thumbprint-only
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
    With -ProvisionEntraApp and without -GenerateLocalCertificate, locates
    {thumbprint}.pfx in the certificate directory when -CertificateFilePath is omitted.
.PARAMETER CertificateMapPath
    JSON profile map; accepts a bare array or an object containing profiles.
.PARAMETER Tenant
    Tenant ID or domain for provisioning and connection, or saved profile filter.
    Profile lookup is case-insensitive and matches tenant, tenantId, organization,
    default domain and profile name. A vanity domain or tenant GUID matches a
    profile stored as the tenant's *.onmicrosoft.com domain.
.PARAMETER SiteUrl
    HTTPS SharePoint site URL. Required for provisioning; also a profile filter.
.PARAMETER AppId
    Application client ID for explicit connections.
.PARAMETER ProfileName
    Saved profile name to select. Defaults to AppDisplayName during provisioning.
    When several profiles match -Tenant, -ProfileName selects among them.
    Matching is case-insensitive.
.PARAMETER ProvisionEntraApp
    Interactive one-time app setup and admin consent. Use with
    -GenerateLocalCertificate, or with an existing PFX
    (-CertificateFilePath/-CertificatePath or -CertificateThumbprint).
    Cannot be combined with -noprompt.
.PARAMETER AppDisplayName
    App name; defaults to GeneratedCertSubject. Duplicate names cause an error.
.PARAMETER SetupClientId
    Optional tenant-approved public client ID for Microsoft Graph device sign-in.
    If omitted, uses the Microsoft Graph PowerShell client
    (14d82eec-204b-4c2f-b7e8-296a70dab67e).
.PARAMETER SharePointPermission
    SharePoint application permission granted during provisioning.
    Sites.FullControl.All (default), Sites.Read.All, or Sites.Selected.
    Graph Application.Read.All is always granted as well. Existing extra
    application permissions on a reused app are left in place.
.PARAMETER DeviceCodeTimeoutSeconds
    Seconds to wait for device-code sign-in during provisioning (1-7200).
    Defaults to 900. Transient network read timeouts are retried until this
    window ends. The sign-in service may still expire its code sooner.
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
.EXAMPLE
    ./o365-connect-pnp-cert-linux.ps1 -GenerateLocalCertificate -ProvisionEntraApp -Tenant contoso.onmicrosoft.com -SiteUrl https://contoso.sharepoint.com -SharePointPermission Sites.Selected -DeviceCodeTimeoutSeconds 900
    Provision with Sites.Selected and a 15-minute device-code window.
.EXAMPLE
    ./o365-connect-pnp-cert-linux.ps1 -ProvisionEntraApp -Tenant ciaops365.com -SiteUrl https://ciaops365.sharepoint.com -CertificatePath /home/admin/.local/share/ciaops/pnp/THUMBPRINT.pfx -SharePointPermission Sites.Read.All
    Reuse an existing PFX for a second tenant instead of minting a new certificate.
.NOTES
    Version: 1.1
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
    [Alias('CertificatePath')][string]$CertificateFilePath,
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
    [string]$SetupClientId,
    [ValidateSet('Sites.FullControl.All', 'Sites.Read.All', 'Sites.Selected')]
    [string]$SharePointPermission = 'Sites.FullControl.All',
    [ValidateRange(1, 7200)][int]$DeviceCodeTimeoutSeconds = 900
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
    $map = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -NoEnumerate -ErrorAction Stop
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
    param(
        [string]$DisplayName,
        [string]$CerPath,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [ValidateSet('Sites.FullControl.All', 'Sites.Read.All', 'Sites.Selected')]
        [string]$SharePointPermission = 'Sites.FullControl.All'
    )
    if ($null -eq $Certificate -and [string]::IsNullOrWhiteSpace($CerPath)) {
        throw 'Provisioning requires a public certificate.'
    }
    $targets = @(
        @{ AppId = '00000003-0000-0ff1-ce00-000000000000'; RoleName = $SharePointPermission },
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
    $ownsCertificate = $false
    $publicCertificate = $Certificate
    if ($null -eq $publicCertificate) {
        $publicCertificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($CerPath)
        $ownsCertificate = $true
    }
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
    finally { if ($ownsCertificate) { $publicCertificate.Dispose() } }
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

function Get-PnpTempDirectory {
    # The process temp variable is often unset on Linux. GetTempPath falls back
    # to /tmp instead of building a path from a null variable.
    $temp = [System.IO.Path]::GetTempPath()
    if ([string]::IsNullOrWhiteSpace($temp)) { throw 'No process temp directory is available.' }
    return $temp
}

function Get-PnpProvisionPermissionWarning {
    param([string]$SharePointPermission)
    switch ($SharePointPermission) {
        'Sites.FullControl.All' { return 'Provisioning grants full control of all SharePoint sites and Graph Application.Read.All.' }
        'Sites.Read.All' { return 'Provisioning grants read access to all SharePoint sites and Graph Application.Read.All.' }
        'Sites.Selected' { return 'Provisioning grants SharePoint Sites.Selected (no site access until individual sites are granted) and Graph Application.Read.All.' }
        default { return "Provisioning grants SharePoint $SharePointPermission and Graph Application.Read.All." }
    }
}

function Test-PnpSameText {
    param($Left, $Right)
    if ($null -eq $Left -or $null -eq $Right) { return $false }
    $leftText = ([string]$Left).Trim()
    $rightText = ([string]$Right).Trim()
    if ($leftText.Length -eq 0 -or $rightText.Length -eq 0) { return $false }
    return $leftText.Equals($rightText, [StringComparison]::OrdinalIgnoreCase)
}

function Add-PnpIdentityValue {
    param($Values, $Value)
    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return }
    $trimmed = $Value.Trim()
    foreach ($existing in $Values) {
        if ($existing.Equals($trimmed, [StringComparison]::OrdinalIgnoreCase)) { return }
    }
    [void]$Values.Add($trimmed)
}

function Get-PnpProfileIdentityValues {
    param($Profile)
    $values = [System.Collections.Generic.List[string]]::new()
    foreach ($name in @('tenant', 'tenantId', 'organization', 'defaultDomain', 'initialDomain', 'name')) {
        Add-PnpIdentityValue -Values $values -Value $Profile.$name
    }
    $organization = $Profile.organization
    if ($null -ne $organization -and $organization -isnot [string]) {
        foreach ($name in @('id', 'tenantId', 'defaultDomain', 'initialDomain')) {
            Add-PnpIdentityValue -Values $values -Value $organization.$name
        }
        foreach ($domain in @($organization.verifiedDomains)) {
            if ($domain -is [string]) { Add-PnpIdentityValue -Values $values -Value $domain }
            elseif ($null -ne $domain) { Add-PnpIdentityValue -Values $values -Value $domain.name }
        }
    }
    foreach ($domain in @($Profile.verifiedDomains)) {
        if ($domain -is [string]) { Add-PnpIdentityValue -Values $values -Value $domain }
        elseif ($null -ne $domain) { Add-PnpIdentityValue -Values $values -Value $domain.name }
    }
    return $values
}

function Test-PnpTransientNetworkError {
    param($ErrorRecord)
    if ($null -eq $ErrorRecord) { return $false }
    $status = 0
    try {
        if ($ErrorRecord.Exception.Response) { $status = [int]$ErrorRecord.Exception.Response.StatusCode }
    }
    catch { $status = 0 }
    if ($status -in 408, 429, 500, 502, 503, 504) { return $true }
    $exception = $ErrorRecord.Exception
    while ($null -ne $exception) {
        $typeName = $exception.GetType().FullName
        if ($typeName -match 'TaskCanceledException|OperationCanceledException|TimeoutException|IOException|SocketException|WebException') {
            return $true
        }
        if ($typeName -eq 'System.Net.Http.HttpRequestException') {
            $statusCode = $null
            try { $statusCode = $exception.StatusCode } catch { $statusCode = $null }
            if ($null -eq $statusCode -or [int]$statusCode -in 408, 429, 500, 502, 503, 504) { return $true }
        }
        $exception = $exception.InnerException
    }
    $message = [string]$ErrorRecord.Exception.Message
    return $message -match '(?i)timed out|timeout|unable to read data|response ended prematurely|connection reset|connection was closed|temporarily unavailable|the ssl connection could not be established|name or service not known|no such host|resource temporarily unavailable|broken pipe'
}

function Get-PnpOAuthError {
    param($ErrorRecord)
    $body = $null
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $body = [string]$ErrorRecord.ErrorDetails.Message
    }
    if ([string]::IsNullOrWhiteSpace($body)) { return $null }
    try {
        $parsed = $body | ConvertFrom-Json -ErrorAction Stop
    }
    catch { return $null }
    $code = $parsed.PSObject.Properties['error']
    if ($null -eq $code -or [string]::IsNullOrWhiteSpace([string]$code.Value)) { return $null }
    return $parsed
}

function Resolve-PnpDirectoryTenantId {
    param([string]$Identifier, [hashtable]$Cache)
    if ([string]::IsNullOrWhiteSpace($Identifier)) { return $null }
    $trimmed = $Identifier.Trim()
    $cacheKey = $trimmed.ToLowerInvariant()
    if ($Cache.ContainsKey($cacheKey)) { return $Cache[$cacheKey] }
    $parsed = [guid]::Empty
    if ([guid]::TryParse($trimmed, [ref]$parsed)) {
        $id = $parsed.ToString().ToLowerInvariant()
        $Cache[$cacheKey] = $id
        return $id
    }
    if ($trimmed -notmatch '^(?=.{1,253}$)[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$') {
        $Cache[$cacheKey] = $null
        return $null
    }
    $resolved = $null
    for ($attempt = 1; $attempt -le 2 -and [string]::IsNullOrWhiteSpace($resolved); $attempt++) {
        try {
            $uri = "https://login.microsoftonline.com/$([uri]::EscapeDataString($trimmed))/v2.0/.well-known/openid-configuration"
            $configuration = Invoke-RestMethod -Method Get -Uri $uri -TimeoutSec 10 -ErrorAction Stop
            $issuer = [string]$configuration.issuer
            if ($issuer -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
                $resolved = $Matches[1].ToLowerInvariant()
            }
            else { break }
        }
        catch {
            if (-not (Test-PnpTransientNetworkError -ErrorRecord $_) -or $attempt -ge 2) { break }
            Start-Sleep -Seconds 1
        }
    }
    $Cache[$cacheKey] = $resolved
    return $resolved
}

function Test-PnpProfileMatchesTenant {
    param($Profile, [string]$TenantFilter, [hashtable]$ResolvedTenantIds)
    $identifiers = @(Get-PnpProfileIdentityValues -Profile $Profile)
    foreach ($identifier in $identifiers) {
        if (Test-PnpSameText $identifier $TenantFilter) { return $true }
    }
    $filterId = Resolve-PnpDirectoryTenantId -Identifier $TenantFilter -Cache $ResolvedTenantIds
    if ([string]::IsNullOrWhiteSpace($filterId)) { return $false }
    foreach ($name in @('tenant', 'tenantId', 'organization', 'defaultDomain', 'initialDomain')) {
        $identifier = $Profile.$name
        if ($identifier -isnot [string]) { continue }
        $identifierId = Resolve-PnpDirectoryTenantId -Identifier $identifier -Cache $ResolvedTenantIds
        if (-not [string]::IsNullOrWhiteSpace($identifierId) -and (Test-PnpSameText $identifierId $filterId)) { return $true }
    }
    $organization = $Profile.organization
    if ($null -ne $organization -and $organization -isnot [string]) {
        foreach ($name in @('id', 'tenantId', 'defaultDomain', 'initialDomain')) {
            $identifierId = Resolve-PnpDirectoryTenantId -Identifier ([string]$organization.$name) -Cache $ResolvedTenantIds
            if (-not [string]::IsNullOrWhiteSpace($identifierId) -and (Test-PnpSameText $identifierId $filterId)) { return $true }
        }
    }
    return $false
}

function Format-PnpProfileList {
    param($Profiles)
    $items = @($Profiles | Where-Object { $null -ne $_ })
    if ($items.Count -eq 0) { return 'none' }
    return (@($items | ForEach-Object {
        $organization = $_.'organization'
        if ($organization -isnot [string]) { $organization = '' }
        "name='$($_.'name')' tenant='$($_.'tenant')' tenantId='$($_.'tenantId')' organization='$organization' defaultDomain='$($_.'defaultDomain')' siteUrl='$($_.'siteUrl')'"
    }) -join '; ')
}

function Select-PnpLinuxProfile {
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Profiles,
        [string]$ProfileName,
        [string]$Tenant,
        [string]$SiteUrl,
        [string]$MapPath
    )
    $available = @($Profiles | Where-Object { $null -ne $_ })
    $candidates = $available
    if (-not [string]::IsNullOrWhiteSpace($SiteUrl)) {
        $candidates = @($candidates | Where-Object { $_.siteUrl -eq $SiteUrl })
    }
    if (-not [string]::IsNullOrWhiteSpace($Tenant)) {
        $resolvedTenantIds = @{}
        $candidates = @($candidates | Where-Object {
            Test-PnpProfileMatchesTenant -Profile $_ -TenantFilter $Tenant -ResolvedTenantIds $resolvedTenantIds
        })
    }
    if (-not [string]::IsNullOrWhiteSpace($ProfileName)) {
        $named = @($candidates | Where-Object { Test-PnpSameText $_.name $ProfileName })
        if ($named.Count -eq 0) {
            throw "No profile named '$ProfileName' matches the supplied filters in '$MapPath'. Compared tenant, tenantId, organization, default domain and profile name. Candidates: [$(Format-PnpProfileList $candidates)]. Available: [$(Format-PnpProfileList $available)]. Specify -ProfileName from that list."
        }
        $candidates = $named
    }
    if ($candidates.Count -eq 1) { return $candidates[0] }
    if ($candidates.Count -eq 0) {
        $filters = @()
        if ($ProfileName) { $filters += "ProfileName='$ProfileName'" }
        if ($Tenant) { $filters += "Tenant='$Tenant'" }
        if ($SiteUrl) { $filters += "SiteUrl='$SiteUrl'" }
        $filterText = if ($filters.Count -gt 0) { $filters -join ', ' } else { 'none' }
        throw "No matching profile in '$MapPath' (filters: $filterText). Compared tenant, tenantId, organization, default domain and profile name. Available: [$(Format-PnpProfileList $available)]."
    }
    throw "Multiple profiles in '$MapPath' match. Specify -ProfileName to select one. Matches: [$(Format-PnpProfileList $candidates)]."
}

function Resolve-PnpExistingPfxPath {
    param([string]$CertificatePath, [string]$Thumbprint, [string]$Directory)
    if (-not [string]::IsNullOrWhiteSpace($CertificatePath)) {
        if (-not (Test-Path -LiteralPath $CertificatePath -PathType Leaf)) {
            throw "Certificate file not found: $CertificatePath"
        }
        return (Resolve-Path -LiteralPath $CertificatePath).ProviderPath
    }
    $normalized = ($Thumbprint -replace '[^0-9a-fA-F]', '').ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        throw '-ProvisionEntraApp requires -GenerateLocalCertificate or an existing PFX specified with -CertificateFilePath/-CertificatePath or -CertificateThumbprint.'
    }
    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
        throw "Certificate directory not found: $Directory. Pass -CertificateFilePath/-CertificatePath for a PFX outside the certificate directory."
    }
    $pfxMatches = @(Get-ChildItem -LiteralPath $Directory -File -Filter '*.pfx' | Where-Object {
        ($_.BaseName -replace '[^0-9a-fA-F]', '').ToUpperInvariant() -eq $normalized
    })
    if ($pfxMatches.Count -eq 1) { return $pfxMatches[0].FullName }
    if ($pfxMatches.Count -gt 1) { throw "Multiple PFX files match thumbprint $normalized in $Directory." }
    throw "No PFX with thumbprint $normalized in $Directory. Pass -CertificateFilePath/-CertificatePath."
}

function Get-PnpJwtPayload {
    param([string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token)) { return $null }
    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { return $null }
    $payload = $parts[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
        1 { return $null }
    }
    try {
        $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload))
        $parsed = $json | ConvertFrom-Json -AsHashtable -ErrorAction Stop
        if ($parsed -is [System.Collections.IDictionary]) { return $parsed }
    }
    catch { return $null }
    return $null
}

function Get-PnpRemainingSeconds {
    param([datetime]$Deadline)
    $remaining = ($Deadline - (Get-Date)).TotalSeconds
    if ($remaining -lt 1) { return 0 }
    return [int][Math]::Floor($remaining)
}

function Get-PnpDeviceCodeAccessToken {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [string]$ClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e',
        [string[]]$Scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All'),
        [Parameter(Mandatory)][int]$TimeoutSeconds
    )
    if ($TimeoutSeconds -lt 1) { throw 'Device-code timeout must be at least 1 second.' }
    $scope = (@($Scopes) | ForEach-Object {
        if ($_ -match '^https?://') { $_ } else { "https://graph.microsoft.com/$_" }
    }) -join ' '
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $tenantSegment = [uri]::EscapeDataString($TenantId.Trim())
    $deviceUri = "https://login.microsoftonline.com/$tenantSegment/oauth2/v2.0/devicecode"
    $tokenUri = "https://login.microsoftonline.com/$tenantSegment/oauth2/v2.0/token"
    $device = $null
    while ((Get-PnpRemainingSeconds -Deadline $deadline) -gt 0) {
        $requestTimeout = [Math]::Min(60, (Get-PnpRemainingSeconds -Deadline $deadline))
        if ($requestTimeout -lt 1) { break }
        try {
            $device = Invoke-RestMethod -Method Post -Uri $deviceUri -Body @{ client_id = $ClientId; scope = $scope } `
                -ContentType 'application/x-www-form-urlencoded' -TimeoutSec $requestTimeout -ErrorAction Stop
            break
        }
        catch {
            $oauthError = Get-PnpOAuthError -ErrorRecord $_
            if ($null -ne $oauthError) {
                $description = [string]$oauthError.error_description
                if ([string]::IsNullOrWhiteSpace($description)) { $description = [string]$oauthError.error }
                throw "Device-code request failed ($($oauthError.error)): $description"
            }
            if (Test-PnpTransientNetworkError -ErrorRecord $_) {
                Write-Warning "Device-code request hit a transient network error ($($_.Exception.Message)). Retrying until the sign-in window ends."
                $pause = [Math]::Min(2, (Get-PnpRemainingSeconds -Deadline $deadline))
                if ($pause -ge 1) { Start-Sleep -Seconds $pause }
                continue
            }
            throw
        }
    }
    if ($null -eq $device -or [string]::IsNullOrWhiteSpace([string]$device.device_code) -or [string]::IsNullOrWhiteSpace([string]$device.user_code)) {
        throw "Device-code sign-in timed out after $TimeoutSeconds seconds. Complete the browser sign-in sooner or raise -DeviceCodeTimeoutSeconds."
    }
    $pollInterval = 5
    if ($null -ne $device.interval -and [int]$device.interval -gt 0) { $pollInterval = [int]$device.interval }
    Write-Host ''
    Write-Host '--- Graph device-code sign-in ---'
    if (-not [string]::IsNullOrWhiteSpace([string]$device.verification_uri)) {
        Write-Host "Sign in at: $($device.verification_uri)"
    }
    Write-Host "Device code: $($device.user_code)"
    if (-not [string]::IsNullOrWhiteSpace([string]$device.message)) { Write-Host $device.message }
    Write-Host "This window waits up to $TimeoutSeconds seconds. Relay the code to an administrator if needed."
    Write-Host '---------------------------------'
    while ((Get-PnpRemainingSeconds -Deadline $deadline) -gt 0) {
        $sleepFor = [Math]::Min($pollInterval, (Get-PnpRemainingSeconds -Deadline $deadline))
        if ($sleepFor -ge 1) { Start-Sleep -Seconds $sleepFor }
        $requestTimeout = [Math]::Min(60, (Get-PnpRemainingSeconds -Deadline $deadline))
        if ($requestTimeout -lt 1) { break }
        try {
            $token = Invoke-RestMethod -Method Post -Uri $tokenUri -Body @{
                grant_type = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id = $ClientId
                device_code = [string]$device.device_code
            } -ContentType 'application/x-www-form-urlencoded' -TimeoutSec $requestTimeout -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace([string]$token.access_token)) {
                throw 'Token endpoint returned no access_token.'
            }
            return [string]$token.access_token
        }
        catch {
            $oauthError = Get-PnpOAuthError -ErrorRecord $_
            if ($null -ne $oauthError -and -not [string]::IsNullOrWhiteSpace([string]$oauthError.error)) {
                # continue must sit directly in catch. Inside switch it does not resume this loop.
                $oauthCode = [string]$oauthError.error.ToLowerInvariant()
                if ($oauthCode -eq 'authorization_pending') { continue }
                if ($oauthCode -eq 'slow_down') { $pollInterval += 5; continue }
                if ($oauthCode -eq 'authorization_declined') { throw 'User declined the device-code authorization request.' }
                if ($oauthCode -eq 'expired_token') { throw 'Device code expired before authorization was completed.' }
                if ($oauthCode -eq 'bad_verification_code') { throw 'Device code was rejected by the sign-in service.' }
                if ($oauthCode -eq 'access_denied') { throw 'Access was denied during device-code sign-in.' }
                $description = [string]$oauthError.error_description
                if ([string]::IsNullOrWhiteSpace($description)) { $description = [string]$oauthError.error }
                throw "Device-code token request failed ($($oauthError.error)): $description"
            }
            if (Test-PnpTransientNetworkError -ErrorRecord $_) {
                Write-Warning "Device-code poll hit a transient network error ($($_.Exception.Message)). Retrying until the sign-in window ends."
                continue
            }
            throw
        }
    }
    throw "Device-code sign-in timed out after $TimeoutSeconds seconds. Complete the browser sign-in sooner or raise -DeviceCodeTimeoutSeconds."
}

if (-not $IsLinux) { throw 'This script requires Linux and PowerShell 7.4.6 or later. Use o365-connect-pnp-cert.ps1 on Windows.' }
if ($GenerateLocalCertificate -and $UseCertificateAuth) { throw 'Specify exactly one of -GenerateLocalCertificate or -UseCertificateAuth.' }
if (-not $GenerateLocalCertificate -and -not $UseCertificateAuth -and -not $ProvisionEntraApp) {
    throw 'Specify exactly one of -GenerateLocalCertificate or -UseCertificateAuth.'
}
if ($UseCertificateAuth -and $ProvisionEntraApp) {
    throw '-ProvisionEntraApp cannot be combined with -UseCertificateAuth.'
}
if ($ProvisionEntraApp -and $noprompt) {
    throw '-ProvisionEntraApp requires interactive administrator sign-in (no -noprompt).'
}
if ($ProvisionEntraApp -and $GenerateLocalCertificate -and -not [string]::IsNullOrWhiteSpace($CertificateFilePath)) {
    throw '-ProvisionEntraApp cannot both -GenerateLocalCertificate and reuse -CertificateFilePath/-CertificatePath. Omit -GenerateLocalCertificate to use the existing PFX.'
}
if ($ProvisionEntraApp -and -not $GenerateLocalCertificate -and [string]::IsNullOrWhiteSpace($CertificateFilePath) -and [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
    throw '-ProvisionEntraApp requires -GenerateLocalCertificate or an existing PFX (-CertificateFilePath/-CertificatePath or -CertificateThumbprint).'
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
    $null = Get-PnpTempDirectory
    if ($enableLog) {
        $logDirectory = Initialize-PnpPrivateDirectory -Path $dataDirectory
        $logPath = Join-Path $logDirectory ("o365-connect-pnp-{0}.txt" -f [guid]::NewGuid())
        Write-PnpPrivateFile -Path $logPath -Bytes ([byte[]]::new(0))
        Start-Transcript -LiteralPath $logPath -Append -ErrorAction Stop | Out-Null
        $transcribing = $true
    }
    if ($GenerateLocalCertificate -or $ProvisionEntraApp) {
        $reuseCertificate = $null
        $disposeReuseCertificate = $false
        try {
            if ($GenerateLocalCertificate) {
                $password = Get-PnpPfxPassword -Password $GeneratedPfxPassword -NoPrompt:$noprompt -EnvironmentVariable $CertificatePasswordEnvironmentVariable
                $generated = New-PnpLinuxCertificate -SubjectName $GeneratedCertSubject -YearsValid $GeneratedCertYearsValid -OutputPath $GeneratedCertOutputPath -Password $password
                Write-Host "Public certificate: $($generated.CerPath)"
                Write-Host "Encrypted private key: $($generated.PfxPath)"
                Write-Host "Thumbprint: $($generated.Thumbprint); expires: $($generated.NotAfter.ToString('yyyy-MM-dd'))"
                if (-not $ProvisionEntraApp) { return }
                $pfxPath = $generated.PfxPath
                $cerPath = $generated.CerPath
                $thumbprint = $generated.Thumbprint
            }
            else {
                $pfxPath = Resolve-PnpExistingPfxPath -CertificatePath $CertificateFilePath -Thumbprint $CertificateThumbprint -Directory $GeneratedCertOutputPath
                $reusePassword = if ($null -ne $CertificatePassword) { $CertificatePassword } else { $GeneratedPfxPassword }
                $password = Get-PnpPfxPassword -Password $reusePassword -NoPrompt:$noprompt -EnvironmentVariable $CertificatePasswordEnvironmentVariable
                $reuseCertificate = Import-PnpLinuxPfx -Path $pfxPath -Password $password
                $disposeReuseCertificate = $true
                if (-not $reuseCertificate.HasPrivateKey) { throw 'The certificate has no private key. Supply a PFX, not a public CER.' }
                $normalizedThumbprint = ($CertificateThumbprint -replace '[^0-9a-fA-F]', '').ToUpperInvariant()
                if ($normalizedThumbprint -and $reuseCertificate.Thumbprint.ToUpperInvariant() -ne $normalizedThumbprint) {
                    throw 'The PFX thumbprint does not match -CertificateThumbprint.'
                }
                $thumbprint = $reuseCertificate.Thumbprint
                Write-Host "Reusing certificate: $pfxPath"
                Write-Host "Thumbprint: $thumbprint; expires: $($reuseCertificate.NotAfter.ToString('yyyy-MM-dd'))"
            }
            Write-Warning (Get-PnpProvisionPermissionWarning -SharePointPermission $SharePointPermission)
            Import-PnpLinuxModule -Name Microsoft.Graph.Authentication -MinimumVersion 2.0.0
            $graphClientId = if ($SetupClientId) { $SetupClientId } else { '14d82eec-204b-4c2f-b7e8-296a70dab67e' }
            $accessToken = Get-PnpDeviceCodeAccessToken -TenantId $Tenant -ClientId $graphClientId -TimeoutSeconds $DeviceCodeTimeoutSeconds
            $graphConnected = $false
            try {
                $tenantIdFromToken = $null
                $payload = Get-PnpJwtPayload -Token $accessToken
                if ($payload -and $payload['tid']) { $tenantIdFromToken = [string]$payload['tid'] }
                $secureToken = ConvertTo-SecureString -String $accessToken -AsPlainText -Force
                # Connect-MgGraph -AccessToken always uses process scope. -ContextScope is not
                # in that parameter set; the SDK's own device-code path cancels after 120 seconds.
                Connect-MgGraph -AccessToken $secureToken -NoWelcome -ErrorAction Stop
                $graphConnected = $true
                $displayName = if ($AppDisplayName) { $AppDisplayName } else { $GeneratedCertSubject }
                $provisionParameters = @{ DisplayName = $displayName; SharePointPermission = $SharePointPermission }
                if ($null -ne $reuseCertificate) { $provisionParameters.Certificate = $reuseCertificate }
                else { $provisionParameters.CerPath = $cerPath }
                $provisioned = New-PnpLinuxEntraApp @provisionParameters
                $entryName = if ($ProfileName) { $ProfileName } else { $displayName }
                $profileEntry = @{
                    name = $entryName; tenant = $Tenant; siteUrl = $SiteUrl; appId = $provisioned.AppId
                    certificateThumbprint = $thumbprint; certificateFilePath = $pfxPath
                }
                if (-not [string]::IsNullOrWhiteSpace($tenantIdFromToken)) { $profileEntry.tenantId = $tenantIdFromToken }
                Save-PnpLinuxProfile -Path $CertificateMapPath -Entry $profileEntry
                Write-Host "Provisioned app: $($provisioned.AppId). Profile saved: $CertificateMapPath"
                Write-Warning 'New app permissions may take time to replicate. Connect from a fresh pwsh session.'
            }
            finally {
                Remove-Variable -Name accessToken -ErrorAction SilentlyContinue
                if ($graphConnected) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
            }
        }
        finally {
            if ($disposeReuseCertificate -and $null -ne $reuseCertificate) { $reuseCertificate.Dispose() }
        }
        return
    }

    if ($ProfileName -or -not ($Tenant -and $SiteUrl -and $AppId -and $CertificateFilePath)) {
        $profile = Select-PnpLinuxProfile -Profiles @(Read-PnpProfiles -Path $CertificateMapPath) `
            -ProfileName $ProfileName -Tenant $Tenant -SiteUrl $SiteUrl -MapPath $CertificateMapPath
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