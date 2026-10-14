<#
.SYNOPSIS
    Generate file-based certificates and connect to Exchange Online on Linux.
.DESCRIPTION
    Linux counterpart of o365-connect-exo-cert.ps1. Select exactly one mode:
    -GenerateLocalCertificate creates an RSA-2048/SHA256 certificate, a public
    DER .cer, and an encrypted private-key .pfx using .NET (no OpenSSL CLI needed).
    -UseCertificateAuth loads a PFX into memory and passes it to Exchange Online.

    Certificates default to $XDG_DATA_HOME/ciaops/exo (or ~/.local/share/ciaops/exo).
    Profiles default to $XDG_CONFIG_HOME/ciaops/exo/o365-exo-cert-auth.json
    (or ~/.config/ciaops/exo/o365-exo-cert-auth.json). Dedicated storage directories
    are restricted to mode 700 and files to mode 600, independent of the umask.
    Custom output/map directories must also be dedicated to this script: their
    permissions are restricted to 700. Symbolic links are not accepted there.

    -ProvisionEntraApp additionally creates/reuses an Entra app, uploads the public
    certificate, grants Exchange.ManageAsApp, and assigns the Exchange service
    principal to Organization Management (full Exchange administration).
    Provisioning requires an administrator able to consent to application
    permissions and manage Exchange role groups. Both sign-ins use device codes.
    Only Microsoft.Graph.Authentication is needed for provisioning; no Graph
    application permissions are granted to the Exchange app itself.

    Requires Linux, PowerShell 7.4+, and ExchangeOnlineManagement 3.5.0+.
    Missing modules are installed for CurrentUser, without sudo. Run this script
    in the PowerShell session where you want to use the Exchange cmdlets.
.PARAMETER GenerateLocalCertificate
    Generate a certificate. Always exports an encrypted PFX, because Linux does
    not use the Windows certificate store. Generation alone needs no modules.
.PARAMETER UseCertificateAuth
    Connect using CertificateFilePath/Organization/AppId or a saved profile.
.PARAMETER GeneratedCertSubject
    Common name for the generated certificate. Defaults to O365-EXO-AppAuth.
.PARAMETER GeneratedCertYearsValid
    Certificate lifetime in years (1-100). Defaults to 2.
.PARAMETER GeneratedCertOutputPath
    Dedicated certificate directory. Defaults to the XDG data location above.
.PARAMETER GeneratedPfxPassword
    Nonempty SecureString for the generated PFX; prompted for if omitted.
.PARAMETER CertificateFilePath
    Existing encrypted PFX. Overrides the profile path. No private key is imported
    into a persistent certificate store.
.PARAMETER CertificatePassword
    PFX password for connection; prompted for if omitted. For unattended use,
    retrieve a SecureString from a secret vault and pass it with -noprompt.
    Passwords are never written to profiles. Do not use Linux ConvertFrom-SecureString
    output as a substitute for a secret vault: it is not Windows DPAPI encryption.
.PARAMETER CertificatePasswordEnvironmentVariable
    Name of an environment variable containing the PFX password, injected by your
    scheduler or secret manager. Used for generation and connection when the
    corresponding SecureString password parameter is omitted. The value is never
    printed or saved. Environment variables are not a secret vault and may be
    accessible to other processes running as the same user; prefer passing a
    SecureString from a vault when running from a PowerShell automation script.
.PARAMETER CertificateThumbprint
    Optional expected thumbprint; also read from profiles and checked against PFX.
.PARAMETER CertificateMapPath
    JSON profile map; accepts a bare array or an object containing profiles.
.PARAMETER Tenant
    Provisioning tenant ID/domain, or profile filter when connecting.
.PARAMETER Organization
    Primary tenant .onmicrosoft.com domain. Required for provisioning if Tenant is
    a GUID; otherwise defaults to Tenant. Also usable for explicit connections.
.PARAMETER AppId
    Application client ID for explicit connections.
.PARAMETER ProfileName
    Saved profile name to select. Defaults to AppDisplayName during provisioning.
.PARAMETER ProvisionEntraApp
    Generate mode only: perform interactive one-time app and Exchange RBAC setup.
.PARAMETER AppDisplayName
    App name for provisioning; defaults to GeneratedCertSubject. Multiple existing
    apps with this name cause an error rather than choosing an arbitrary app.
.PARAMETER SetupClientId
    Optional tenant-approved public client ID for Microsoft Graph device sign-in.
    If omitted, uses the Microsoft Graph PowerShell client.
.PARAMETER noprompt
    Disable prompts. Alias: -Headless. Requires a password parameter or a named
    password environment variable, and preinstalled modules. Skips update checks.
    Cannot be combined with ProvisionEntraApp, which requires admin sign-in.
.PARAMETER noupdate
    Skip the Exchange module Gallery version check. Updates otherwise require
    confirmation and are never attempted in -noprompt mode.
.PARAMETER enableLog
    Save a timestamped transcript in the private certificate data directory.
    Provisioning logs may contain short-lived device codes; protect these logs.
.EXAMPLE
    ./o365-connect-exo-cert-linux.ps1 -GenerateLocalCertificate -ProvisionEntraApp -Tenant contoso.onmicrosoft.com
    Create encrypted certificate files, provision the app and save a profile.
.EXAMPLE
    ./o365-connect-exo-cert-linux.ps1 -UseCertificateAuth -Tenant contoso.onmicrosoft.com
    Prompt for the saved profile's PFX password and connect in this session.
.EXAMPLE
    ./o365-connect-exo-cert-linux.ps1 -UseCertificateAuth -Organization contoso.onmicrosoft.com -AppId 11111111-2222-3333-4444-555555555555 -CertificateFilePath /home/admin/.local/share/ciaops/exo/auth.pfx -CertificatePassword $vaultPassword -noprompt -noupdate
    Connect unattended with a SecureString retrieved separately from a secret vault.
.EXAMPLE
    pwsh -NoProfile -NonInteractive -File ./o365-connect-exo-cert-linux.ps1 -UseCertificateAuth -Tenant ciaopslabs.onmicrosoft.com -Headless -CertificatePasswordEnvironmentVariable EXO_PFX_PASSWORD
    Connect without prompts using EXO_PFX_PASSWORD injected into the process by
    a scheduler or secret manager. Provision the profile and install modules first.
    To run Exchange commands, invoke this script from the same PowerShell job that
    runs those commands; a connection ends when its pwsh process exits.
.NOTES
    Documentation: https://github.com/directorcia/Office365/wiki/Exchange-Online-certificate-authentication-on-Linux
    CIAOPS - Script provided as is. Use at own risk. No guarantees or warranty provided.
#>
#requires -Version 7.4
[CmdletBinding()]
param(
    [Alias('Headless')][switch]$noprompt,
    [switch]$noupdate,
    [switch]$enableLog,
    [switch]$GenerateLocalCertificate,
    [switch]$UseCertificateAuth,
    [string]$GeneratedCertSubject = 'O365-EXO-AppAuth',
    [ValidateRange(1, 100)][int]$GeneratedCertYearsValid = 2,
    [string]$GeneratedCertOutputPath,
    [securestring]$GeneratedPfxPassword,
    [string]$CertificateFilePath,
    [securestring]$CertificatePassword,
    [string]$CertificatePasswordEnvironmentVariable,
    [string]$CertificateThumbprint,
    [string]$CertificateMapPath,
    [string]$Tenant,
    [string]$Organization,
    [string]$AppId,
    [string]$ProfileName,
    [switch]$ProvisionEntraApp,
    [string]$AppDisplayName,
    [string]$SetupClientId
)

function Initialize-ExoPrivateDirectory {
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

function Write-ExoPrivateFile {
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

function Read-ExoProfiles {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $map = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -ErrorAction Stop
    if ($map -is [System.Collections.IDictionary] -and $map.Contains('profiles')) {
        return $map.profiles
    }
    if ($map -is [array]) { return $map }
    throw "Invalid profile map: $Path. Expected an array or an object containing profiles."
}

function Save-ExoLinuxProfile {
    param([string]$Path, [hashtable]$Entry)
    $directory = Initialize-ExoPrivateDirectory -Path (Split-Path -Parent $Path)
    $lockPath = Join-Path $directory '.profiles.lock'
    $lockOptions = [IO.FileStreamOptions]::new()
    $lockOptions.Mode = [IO.FileMode]::OpenOrCreate
    $lockOptions.Access = [IO.FileAccess]::ReadWrite
    $lockOptions.Share = [IO.FileShare]::None
    $lockOptions.UnixCreateMode = [IO.UnixFileMode]::UserRead -bor [IO.UnixFileMode]::UserWrite
    $lockStream = [IO.FileStream]::new($lockPath, $lockOptions)
    $temporaryPath = Join-Path $directory ([guid]::NewGuid().ToString() + '.tmp')
    try {
        $profiles = @(Read-ExoProfiles -Path $Path | Where-Object {
            -not ($_.appId -eq $Entry.appId -and $_.name -eq $Entry.name)
        }) + @($Entry)
        $json = @{ profiles = $profiles } | ConvertTo-Json -Depth 10
        Write-ExoPrivateFile -Path $temporaryPath -Bytes ([Text.Encoding]::UTF8.GetBytes($json))
        [IO.File]::Move($temporaryPath, $Path, $true)
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force }
        $lockStream.Dispose()
    }
}

function Get-ExoPfxPassword {
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

function New-ExoLinuxCertificate {
    param([string]$SubjectName, [int]$YearsValid, [string]$OutputPath, [securestring]$Password)
    $directory = Initialize-ExoPrivateDirectory -Path $OutputPath
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
        Write-ExoPrivateFile -Path $pfxPath -Bytes ($certificate.Export([Security.Cryptography.X509Certificates.X509ContentType]::Pfx, $Password))
        Write-ExoPrivateFile -Path $cerPath -Bytes ($certificate.Export([Security.Cryptography.X509Certificates.X509ContentType]::Cert))
        return @{ Thumbprint = $certificate.Thumbprint; PfxPath = $pfxPath; CerPath = $cerPath; NotAfter = $certificate.NotAfter }
    }
    finally {
        if ($null -ne $certificate) { $certificate.Dispose() }
        $rsa.Dispose()
    }
}

function Import-ExoLinuxPfx {
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

function Import-ExoLinuxModule {
    param([string]$Name, [version]$MinimumVersion, [switch]$CheckUpdate)
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

function Invoke-ExoLinuxGraph {
    param([string]$Method = 'GET', [string]$Path, [hashtable]$Body)
    $parameters = @{ Method = $Method; Uri = "https://graph.microsoft.com/v1.0/$Path"; OutputType = 'Hashtable'; ErrorAction = 'Stop' }
    if ($null -ne $Body) {
        $parameters.Body = $Body | ConvertTo-Json -Depth 20 -Compress
        $parameters.ContentType = 'application/json'
    }
    Invoke-MgGraphRequest @parameters
}

function New-ExoLinuxEntraApp {
    param([string]$DisplayName, [string]$CerPath)
    $exchangeResource = '00000002-0000-0ff1-ce00-000000000000'
    $resourceFilter = [uri]::EscapeDataString("appId eq '$exchangeResource'")
    $resources = Invoke-ExoLinuxGraph -Path "servicePrincipals?`$filter=$resourceFilter"
    $resource = @($resources.value)[0]
    $role = @($resource.appRoles | Where-Object { $_.value -eq 'Exchange.ManageAsApp' -and $_.isEnabled -and $_.allowedMemberTypes -contains 'Application' })
    if ($role.Count -ne 1) { throw 'Unable to resolve Exchange.ManageAsApp in the provisioning tenant.' }
    $appFilter = [uri]::EscapeDataString("displayName eq '$($DisplayName.Replace("'", "''"))'")
    $apps = Invoke-ExoLinuxGraph -Path "applications?`$filter=$appFilter"
    if (@($apps.value).Count -gt 1) { throw "Multiple applications named '$DisplayName'. Use a unique -AppDisplayName." }
    $application = @($apps.value)[0]
    if ($null -eq $application) {
        $application = Invoke-ExoLinuxGraph -Method POST -Path 'applications' -Body @{
            displayName = $DisplayName; signInAudience = 'AzureADMyOrg'
            requiredResourceAccess = @(@{ resourceAppId = $exchangeResource; resourceAccess = @(@{ id = $role[0].id; type = 'Role' }) })
        }
    }
    $publicCertificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($CerPath)
    try {
        $existing = Invoke-ExoLinuxGraph -Path "applications/$($application.id)?`$select=keyCredentials"
        $keyIdentifier = [Convert]::ToBase64String($publicCertificate.GetCertHash())
        $keys = @($existing.keyCredentials | Where-Object { $_.customKeyIdentifier -ne $keyIdentifier })
        $keys += @(@{
            type = 'AsymmetricX509Cert'; usage = 'Verify'; displayName = "EXO-$($publicCertificate.Thumbprint)"
            key = [Convert]::ToBase64String($publicCertificate.RawData); customKeyIdentifier = $keyIdentifier
        })
        Invoke-ExoLinuxGraph -Method PATCH -Path "applications/$($application.id)" -Body @{ keyCredentials = $keys } | Out-Null
    }
    finally { $publicCertificate.Dispose() }
    $principalFilter = [uri]::EscapeDataString("appId eq '$($application.appId)'")
    $principals = Invoke-ExoLinuxGraph -Path "servicePrincipals?`$filter=$principalFilter"
    $principal = @($principals.value)[0]
    if ($null -eq $principal) {
        $principal = Invoke-ExoLinuxGraph -Method POST -Path 'servicePrincipals' -Body @{ appId = $application.appId }
    }
    $assignments = Invoke-ExoLinuxGraph -Path "servicePrincipals/$($principal.id)/appRoleAssignments"
    if (-not ($assignments.value | Where-Object { $_.resourceId -eq $resource.id -and $_.appRoleId -eq $role[0].id })) {
        Invoke-ExoLinuxGraph -Method POST -Path "servicePrincipals/$($principal.id)/appRoleAssignments" -Body @{
            principalId = $principal.id; resourceId = $resource.id; appRoleId = $role[0].id
        } | Out-Null
    }
    return @{ AppId = $application.appId; SpObjId = $principal.id }
}

if (-not $IsLinux) { throw 'This script requires Linux and PowerShell 7.4 or later. Use o365-connect-exo-cert.ps1 on Windows.' }
if ($GenerateLocalCertificate -eq $UseCertificateAuth) { throw 'Specify exactly one of -GenerateLocalCertificate or -UseCertificateAuth.' }
if ($ProvisionEntraApp -and (-not $GenerateLocalCertificate -or $noprompt)) {
    throw '-ProvisionEntraApp requires -GenerateLocalCertificate and interactive administrator sign-in (no -noprompt).'
}
if ($ProvisionEntraApp) {
    if (-not $Tenant) { throw '-ProvisionEntraApp requires -Tenant.' }
    if (-not $Organization) { $Organization = $Tenant }
    if ($Organization -notmatch '^[a-zA-Z0-9-]+\.onmicrosoft\.com$') {
        throw 'Provisioning requires -Organization to be the primary tenant .onmicrosoft.com domain.'
    }
}
$dataRoot = if ($env:XDG_DATA_HOME -and [IO.Path]::IsPathRooted($env:XDG_DATA_HOME)) { $env:XDG_DATA_HOME } else { Join-Path $HOME '.local/share' }
$configRoot = if ($env:XDG_CONFIG_HOME -and [IO.Path]::IsPathRooted($env:XDG_CONFIG_HOME)) { $env:XDG_CONFIG_HOME } else { Join-Path $HOME '.config' }
$dataDirectory = Join-Path $dataRoot 'ciaops/exo'
if (-not $GeneratedCertOutputPath) { $GeneratedCertOutputPath = $dataDirectory }
if (-not $CertificateMapPath) { $CertificateMapPath = Join-Path $configRoot 'ciaops/exo/o365-exo-cert-auth.json' }
$GeneratedCertOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($GeneratedCertOutputPath)
$CertificateMapPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CertificateMapPath)
$transcribing = $false
$localCertificate = $null
try {
    if ($enableLog) {
        $logDirectory = Initialize-ExoPrivateDirectory -Path $dataDirectory
        $logPath = Join-Path $logDirectory ("o365-connect-exo-{0}.txt" -f [guid]::NewGuid())
        Write-ExoPrivateFile -Path $logPath -Bytes ([byte[]]::new(0))
        Start-Transcript -LiteralPath $logPath -Append -ErrorAction Stop | Out-Null
        $transcribing = $true
    }
    if ($GenerateLocalCertificate) {
        $password = Get-ExoPfxPassword -Password $GeneratedPfxPassword -NoPrompt:$noprompt -EnvironmentVariable $CertificatePasswordEnvironmentVariable
        $generated = New-ExoLinuxCertificate -SubjectName $GeneratedCertSubject -YearsValid $GeneratedCertYearsValid -OutputPath $GeneratedCertOutputPath -Password $password
        Write-Host "Public certificate: $($generated.CerPath)"
        Write-Host "Encrypted private key: $($generated.PfxPath)"
        Write-Host "Thumbprint: $($generated.Thumbprint); expires: $($generated.NotAfter.ToString('yyyy-MM-dd'))"
        if (-not $ProvisionEntraApp) { return }
        Import-ExoLinuxModule -Name ExchangeOnlineManagement -MinimumVersion 3.5.0 -CheckUpdate:(-not $noupdate)
        Import-ExoLinuxModule -Name Microsoft.Graph.Authentication -MinimumVersion 2.0.0
        $graphParameters = @{
            TenantId = $Tenant; UseDeviceAuthentication = $true; ContextScope = 'Process'; NoWelcome = $true
            Scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All'); ErrorAction = 'Stop'
        }
        if ($SetupClientId) { $graphParameters.ClientId = $SetupClientId }
        Connect-MgGraph @graphParameters
        try {
            $displayName = if ($AppDisplayName) { $AppDisplayName } else { $GeneratedCertSubject }
            $provisioned = New-ExoLinuxEntraApp -DisplayName $displayName -CerPath $generated.CerPath
            Write-Host "Provisioned app: $($provisioned.AppId). Assigning full Exchange administration rights."
            Connect-ExchangeOnline -Device -Organization $Organization -ShowBanner:$false -ErrorAction Stop
            try {
                $graphTenantId = (Get-MgContext -ErrorAction Stop).TenantId
                $exchangeTenantId = (Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1).TenantID
                $domains = @(Get-AcceptedDomain -ErrorAction Stop | ForEach-Object { $_.DomainName.ToString() })
                if (-not $graphTenantId -or $exchangeTenantId -ne $graphTenantId -or $domains -notcontains $Organization) {
                    throw "The Graph and Exchange sign-ins must be for '$Organization' in the same tenant; refusing RBAC changes."
                }
                $exoPrincipal = Get-ServicePrincipal -ErrorAction Stop | Where-Object { $_.AppId -eq $provisioned.AppId } | Select-Object -First 1
                if ($null -eq $exoPrincipal) {
                    New-ServicePrincipal -AppId $provisioned.AppId -ObjectId $provisioned.SpObjId -DisplayName "EXO-App-$($provisioned.AppId)" -ErrorAction Stop | Out-Null
                    $exoPrincipal = Get-ServicePrincipal -Identity $provisioned.SpObjId -ErrorAction Stop
                }
                try { Add-RoleGroupMember -Identity 'Organization Management' -Member $exoPrincipal.Identity -ErrorAction Stop | Out-Null }
                catch {
                    if ($_.Exception.Message -notmatch 'already a member|already exists|is already present') { throw }
                }
            }
            finally { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
            $entryName = if ($ProfileName) { $ProfileName } else { $displayName }
            Save-ExoLinuxProfile -Path $CertificateMapPath -Entry @{
                name = $entryName; tenant = $Tenant; organization = $Organization; appId = $provisioned.AppId
                certificateThumbprint = $generated.Thumbprint; certificateFilePath = $generated.PfxPath
            }
            Write-Host "Profile saved: $CertificateMapPath"
            Write-Warning 'New app permissions and Exchange RBAC may take 15-30 minutes to replicate.'
        }
        finally { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
        return
    }

    if ($ProfileName -or -not ($Organization -and $AppId -and $CertificateFilePath)) {
        $profiles = @(Read-ExoProfiles -Path $CertificateMapPath | Where-Object {
            (-not $ProfileName -or $_.name -eq $ProfileName) -and
            (-not $Tenant -or $_.tenant -eq $Tenant -or $_.organization -eq $Tenant) -and
            (-not $Organization -or $_.organization -eq $Organization)
        })
        if ($profiles.Count -ne 1) { throw "Expected one matching profile in '$CertificateMapPath', found $($profiles.Count). Specify -ProfileName or explicit connection parameters." }
        $profile = $profiles[0]
        if (-not $Organization) { $Organization = $profile.organization }
        if (-not $AppId) { $AppId = $profile.appId }
        if (-not $CertificateFilePath) {
            $CertificateFilePath = $profile.certificateFilePath
            if ($CertificateFilePath -and -not [IO.Path]::IsPathRooted($CertificateFilePath)) {
                $CertificateFilePath = Join-Path (Split-Path -Parent $CertificateMapPath) $CertificateFilePath
            }
        }
        if (-not $CertificateThumbprint) { $CertificateThumbprint = $profile.certificateThumbprint }
    }
    if (-not ($Organization -and $AppId -and $CertificateFilePath)) {
        throw 'Connection requires Organization, AppId and CertificateFilePath. Windows thumbprint-only profiles need a certificateFilePath pointing to an exported PFX.'
    }
    $password = Get-ExoPfxPassword -Password $CertificatePassword -NoPrompt:$noprompt -EnvironmentVariable $CertificatePasswordEnvironmentVariable
    $localCertificate = Import-ExoLinuxPfx -Path $CertificateFilePath -Password $password
    if (-not $localCertificate.HasPrivateKey) { throw 'The certificate has no private key. Supply a PFX, not a public CER.' }
    if ($CertificateThumbprint -and $localCertificate.Thumbprint -ne ($CertificateThumbprint -replace '[^0-9a-fA-F]', '')) {
        throw 'The PFX thumbprint does not match the expected certificate thumbprint.'
    }
    if ($localCertificate.NotBefore -gt (Get-Date) -or $localCertificate.NotAfter -le (Get-Date)) { throw 'The certificate is not currently valid.' }
    if ($localCertificate.NotAfter -lt (Get-Date).AddDays(30)) { Write-Warning 'The certificate expires in fewer than 30 days.' }
    Import-ExoLinuxModule -Name ExchangeOnlineManagement -MinimumVersion 3.5.0 -CheckUpdate:(-not $noupdate)
    Connect-ExchangeOnline -Organization $Organization -AppId $AppId -Certificate $localCertificate -ShowBanner:$false -ErrorAction Stop | Out-Null
    Write-Host "Connected to $Organization using app $AppId and certificate $($localCertificate.Thumbprint)."
}
finally {
    if ($null -ne $localCertificate) { $localCertificate.Dispose() }
    if ($transcribing) { Stop-Transcript | Out-Null }
}