<#
.SYNOPSIS
    Generate file-based certificates and connect to Microsoft Teams on Linux.
.DESCRIPTION
    Linux counterpart of o365-connect-tms-cert.ps1, following the Exchange and
    PnP Linux certificate scripts. Select exactly one mode:
    -GenerateLocalCertificate creates an RSA-2048/SHA256 certificate, a public
    DER .cer and an encrypted private-key .pfx using .NET (no OpenSSL needed).
    -UseCertificateAuth loads the PFX with an ephemeral private key and passes
    the certificate object to Connect-MicrosoftTeams, not a store thumbprint.

    Certificates default to $XDG_DATA_HOME/ciaops/tms (or ~/.local/share/ciaops/tms).
    Profiles default to $XDG_CONFIG_HOME/ciaops/tms/o365-tms-cert-auth.json
    (or ~/.config/ciaops/tms/o365-tms-cert-auth.json). Dedicated storage directories
    are restricted to mode 700 and files to mode 600, independent of the umask.
    Custom output/map directories must also be dedicated to this script: their
    permissions are restricted to 700. Symbolic links are not accepted there.

    -ProvisionEntraApp creates/reuses an Entra app, uploads the public certificate,
    grants the documented Graph permissions for Teams app-only administration,
    assigns the Teams Administrator directory role and saves a connection profile.
    This grants broad Teams administration and write access to groups, Teams,
    channels and the app catalog, but not Directory.ReadWrite.All. Provisioning
    uses Graph device-code sign-in and requires an administrator allowed to grant
    Graph application consent and assign directory roles, such as a Privileged
    Role Administrator or Global Administrator.

    Requires Linux, PowerShell 7.4+ and MicrosoftTeams 5.0.0+ for connections.
    Missing modules are installed for CurrentUser without sudo. Generation alone
    needs no modules; provisioning needs only Microsoft.Graph.Authentication.
    Run in the same PowerShell session as subsequent Teams commands. Successful
    connections are left open. Some Teams cmdlets do not support app-only auth;
    see the linked Microsoft documentation. Administrative Unit scenarios may
    require additional permissions not granted by this script.
.PARAMETER GenerateLocalCertificate
    Generate a certificate. Always exports an encrypted PFX on Linux.
.PARAMETER UseCertificateAuth
    Connect using Tenant/AppId/CertificateFilePath or a saved profile.
.PARAMETER GeneratedCertSubject
    Certificate common name and default app name. Defaults to O365-TMS-AppAuth.
.PARAMETER GeneratedCertYearsValid
    Certificate lifetime in years (1-100). Defaults to 2.
.PARAMETER GeneratedCertOutputPath
    Dedicated certificate directory. Defaults to the XDG data location above.
.PARAMETER GeneratedPfxPassword
    Nonempty SecureString for the generated PFX; prompted for if omitted.
.PARAMETER CertificateFilePath
    Existing encrypted PFX. Overrides the profile path. Windows thumbprint-only
    profiles need this parameter or an added certificateFilePath property.
.PARAMETER CertificatePassword
    PFX password for connection; prompted for if omitted. For unattended use,
    pass a SecureString from a secret vault with -noprompt. Passwords are never
    saved in profiles. Linux ConvertFrom-SecureString is not DPAPI encryption.
.PARAMETER CertificatePasswordEnvironmentVariable
    Name of an environment variable containing the PFX password, injected by a
    scheduler or secret manager. Used when the relevant password is omitted.
    Never printed or saved. Environment values may be accessible to processes
    running as the same user; prefer a SecureString from a secret vault.
.PARAMETER CertificateThumbprint
    Optional expected thumbprint; also read from profiles and checked against PFX.
.PARAMETER CertificateMapPath
    JSON profile map; accepts a bare array or an object containing profiles.
.PARAMETER Tenant
    Tenant ID or domain for provisioning and connection, or saved profile filter.
    Alias: Organization, matching the Windows Teams connection script.
.PARAMETER AppId
    Application client ID for explicit connections.
.PARAMETER ProfileName
    Saved profile name to select. Defaults to AppDisplayName during provisioning.
.PARAMETER ProvisionEntraApp
    Generate mode only: interactive one-time app setup, consent and role assignment.
.PARAMETER AppDisplayName
    App name; defaults to GeneratedCertSubject. Duplicate names cause an error.
.PARAMETER SetupClientId
    Optional tenant-approved public client ID for Microsoft Graph device sign-in.
    If omitted, uses the Microsoft Graph PowerShell client.
.PARAMETER noprompt
    Disable prompts. Alias: Headless. Requires a supplied password and preinstalled
    modules. Skips update checks. Cannot be combined with ProvisionEntraApp.
.PARAMETER noupdate
    Skip the Teams module Gallery version check. Updates otherwise need confirmation.
.PARAMETER enableLog
    Save a transcript in the private data directory. Provisioning transcripts may
    contain short-lived device codes; protect these logs.
.EXAMPLE
    ./o365-connect-tms-cert-linux.ps1 -GenerateLocalCertificate -ProvisionEntraApp -Tenant contoso.onmicrosoft.com
    Create encrypted certificate files, provision the app and save a profile.
.EXAMPLE
    ./o365-connect-tms-cert-linux.ps1 -UseCertificateAuth -ProfileName O365-TMS-AppAuth
    Prompt for the saved profile's PFX password and connect in this session.
.EXAMPLE
    ./o365-connect-tms-cert-linux.ps1 -UseCertificateAuth -Tenant contoso.onmicrosoft.com -AppId 11111111-2222-3333-4444-555555555555 -CertificateFilePath /home/admin/auth.pfx -CertificatePassword $vaultPassword -Headless -noupdate
    Connect using an existing app and exported PFX (chmod 600 first).
.EXAMPLE
    ./o365-connect-tms-cert-linux.ps1 -UseCertificateAuth -ProfileName O365-TMS-AppAuth -Headless -CertificatePasswordEnvironmentVariable TMS_PFX_PASSWORD
    Connect without prompts using a password injected by a secret manager.
.NOTES
    CIAOPS - Script provided as is. Use at own risk. No guarantees or warranty provided.
.LINK
    https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-application-authentication
.LINK
    https://github.com/directorcia/Office365/wiki/Microsoft-Teams-Certificate-Authentication-on-Linux
#>
#requires -Version 7.4
[CmdletBinding()]
param(
    [Alias('Headless')][switch]$noprompt,
    [switch]$noupdate,
    [switch]$enableLog,
    [switch]$GenerateLocalCertificate,
    [switch]$UseCertificateAuth,
    [string]$GeneratedCertSubject = 'O365-TMS-AppAuth',
    [ValidateRange(1, 100)][int]$GeneratedCertYearsValid = 2,
    [string]$GeneratedCertOutputPath,
    [securestring]$GeneratedPfxPassword,
    [string]$CertificateFilePath,
    [securestring]$CertificatePassword,
    [string]$CertificatePasswordEnvironmentVariable,
    [string]$CertificateThumbprint,
    [string]$CertificateMapPath,
    [Alias('Organization')][string]$Tenant,
    [string]$AppId,
    [string]$ProfileName,
    [switch]$ProvisionEntraApp,
    [string]$AppDisplayName,
    [string]$SetupClientId
)

function Initialize-TeamsPrivateDirectory {
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
    $mode = [IO.UnixFileMode]::UserRead -bor [IO.UnixFileMode]::UserWrite -bor [IO.UnixFileMode]::UserExecute
    [IO.Directory]::CreateDirectory($fullPath, $mode) | Out-Null
    [IO.File]::SetUnixFileMode($fullPath, $mode)
    return $fullPath
}

function Write-TeamsPrivateFile {
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

function Read-TeamsProfiles {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $map = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -NoEnumerate -ErrorAction Stop
    if ($map -is [System.Collections.IDictionary] -and $map.Contains('profiles')) { return $map.profiles }
    if ($map -is [array]) { return $map }
    throw "Invalid profile map: $Path. Expected an array or an object containing profiles."
}

function Save-TeamsLinuxProfile {
    param([string]$Path, [hashtable]$Entry)
    $directory = Initialize-TeamsPrivateDirectory -Path (Split-Path -Parent $Path)
    $lockOptions = [IO.FileStreamOptions]::new()
    $lockOptions.Mode = [IO.FileMode]::OpenOrCreate
    $lockOptions.Access = [IO.FileAccess]::ReadWrite
    $lockOptions.Share = [IO.FileShare]::None
    $lockOptions.UnixCreateMode = [IO.UnixFileMode]::UserRead -bor [IO.UnixFileMode]::UserWrite
    $lockStream = [IO.FileStream]::new((Join-Path $directory '.profiles.lock'), $lockOptions)
    $temporaryPath = Join-Path $directory ([guid]::NewGuid().ToString() + '.tmp')
    try {
        $profiles = @(Read-TeamsProfiles -Path $Path | Where-Object {
            -not ($_.appId -eq $Entry.appId -and $_.name -eq $Entry.name)
        }) + @($Entry)
        $json = @{ profiles = $profiles } | ConvertTo-Json -Depth 10
        Write-TeamsPrivateFile -Path $temporaryPath -Bytes ([Text.Encoding]::UTF8.GetBytes($json))
        [IO.File]::Move($temporaryPath, $Path, $true)
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force }
        $lockStream.Dispose()
    }
}

function Get-TeamsPfxPassword {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '',
        Justification = 'Converts a caller-injected process environment secret in memory; never prints or persists it. SecureString vault input is preferred.')]
    [CmdletBinding()]
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

function New-TeamsLinuxCertificate {
    param([string]$SubjectName, [int]$YearsValid, [string]$OutputPath, [securestring]$Password)
    $directory = Initialize-TeamsPrivateDirectory -Path $OutputPath
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
        Write-TeamsPrivateFile -Path $pfxPath -Bytes ($certificate.Export([Security.Cryptography.X509Certificates.X509ContentType]::Pfx, $Password))
        Write-TeamsPrivateFile -Path $cerPath -Bytes ($certificate.Export([Security.Cryptography.X509Certificates.X509ContentType]::Cert))
        return @{ Thumbprint = $certificate.Thumbprint; PfxPath = $pfxPath; CerPath = $cerPath; NotAfter = $certificate.NotAfter }
    }
    finally {
        if ($null -ne $certificate) { $certificate.Dispose() }
        $rsa.Dispose()
    }
}

function Import-TeamsLinuxPfx {
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

function Import-TeamsLinuxModule {
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
        try { $latest = Find-Module -Name $Name -Repository PSGallery -ErrorAction Stop }
        catch { Write-Warning "Module update check failed: $($_.Exception.Message)"; $latest = $null }
        if ($null -ne $latest -and [version]$latest.Version -gt $available.Version -and
            (Read-Host "Install $Name $($latest.Version)? [y/N]") -eq 'y') {
            Install-Module -Name $Name -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop
        }
    }
    Import-Module -Name $Name -MinimumVersion $MinimumVersion -ErrorAction Stop
}

function Invoke-TeamsLinuxGraph {
    param([string]$Method = 'GET', [string]$Path, [hashtable]$Body)
    $parameters = @{ Method = $Method; Uri = "https://graph.microsoft.com/v1.0/$Path"; OutputType = 'Hashtable'; ErrorAction = 'Stop' }
    if ($null -ne $Body) {
        $parameters.Body = $Body | ConvertTo-Json -Depth 20 -Compress
        $parameters.ContentType = 'application/json'
    }
    Invoke-MgGraphRequest @parameters
}

function Get-TeamsLinuxGraphCollection {
    param([string]$Path)
    do {
        $page = Invoke-TeamsLinuxGraph -Path $Path
        $page.value
        $nextLink = $page.'@odata.nextLink'
        if ($nextLink) {
            $prefix = 'https://graph.microsoft.com/v1.0/'
            if (-not $nextLink.StartsWith($prefix, [StringComparison]::Ordinal)) {
                throw "Unexpected Microsoft Graph pagination URL: $nextLink"
            }
            $Path = $nextLink.Substring($prefix.Length)
        }
    } while ($nextLink)
}

function New-TeamsLinuxEntraApp {
    param([string]$DisplayName, [string]$CerPath)
    $graphAppId = '00000003-0000-0000-c000-000000000000'
    $resources = @(Get-TeamsLinuxGraphCollection -Path "servicePrincipals?`$filter=appId%20eq%20'$graphAppId'")
    if ($resources.Count -ne 1) { throw 'Unable to resolve the Microsoft Graph service principal.' }
    $graphResource = $resources[0]
    $roleIds = @(
        foreach ($roleName in @('Organization.Read.All', 'User.Read.All', 'Group.ReadWrite.All',
            'AppCatalog.ReadWrite.All', 'TeamSettings.ReadWrite.All', 'Channel.Delete.All',
            'ChannelSettings.ReadWrite.All', 'ChannelMember.ReadWrite.All')) {
            $roles = @($graphResource.appRoles | Where-Object {
                $_.value -eq $roleName -and $_.isEnabled -and $_.allowedMemberTypes -contains 'Application'
            })
            if ($roles.Count -ne 1) { throw "Unable to resolve Graph application permission '$roleName'." }
            $roles[0].id
        }
    )
    $roleDefinitions = @(Get-TeamsLinuxGraphCollection -Path "roleManagement/directory/roleDefinitions?`$filter=displayName%20eq%20'Teams%20Administrator'")
    if ($roleDefinitions.Count -ne 1) { throw 'Unable to resolve the Teams Administrator directory role.' }
    $appFilter = [uri]::EscapeDataString("displayName eq '$($DisplayName.Replace("'", "''"))'")
    $apps = @(Get-TeamsLinuxGraphCollection -Path "applications?`$filter=$appFilter")
    if ($apps.Count -gt 1) { throw "Multiple applications named '$DisplayName'. Use a unique -AppDisplayName." }
    $application = if ($apps.Count -eq 1) { $apps[0] } else {
        Invoke-TeamsLinuxGraph -Method POST -Path 'applications' -Body @{
            displayName = $DisplayName; signInAudience = 'AzureADMyOrg'
        }
    }
    $publicCertificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($CerPath)
    try {
        $existing = Invoke-TeamsLinuxGraph -Path "applications/$($application.id)?`$select=keyCredentials,requiredResourceAccess"
        $keyIdentifier = [Convert]::ToBase64String($publicCertificate.GetCertHash())
        $keys = @($existing.keyCredentials | Where-Object { $_.customKeyIdentifier -ne $keyIdentifier })
        foreach ($existingKey in $keys) {
            if ([string]::IsNullOrWhiteSpace($existingKey.key)) {
                throw "Graph did not return public key material for credential '$($existingKey.keyId)'. Refusing to remove existing certificates."
            }
        }
        $keys += @(@{
            type = 'AsymmetricX509Cert'; usage = 'Verify'; displayName = "TMS-$($publicCertificate.Thumbprint)"
            key = [Convert]::ToBase64String($publicCertificate.RawData); customKeyIdentifier = $keyIdentifier
            startDateTime = $publicCertificate.NotBefore.ToUniversalTime().ToString('o')
            endDateTime = $publicCertificate.NotAfter.ToUniversalTime().ToString('o')
        })
        $permissions = @($existing.requiredResourceAccess)
        $graphPermission = $permissions | Where-Object { $_.resourceAppId -eq $graphAppId } | Select-Object -First 1
        if ($null -eq $graphPermission) {
            $graphPermission = @{ resourceAppId = $graphAppId; resourceAccess = @() }
            $permissions += @($graphPermission)
        }
        foreach ($roleId in $roleIds) {
            if (-not ($graphPermission.resourceAccess | Where-Object { $_.id -eq $roleId -and $_.type -eq 'Role' })) {
                $graphPermission.resourceAccess = @($graphPermission.resourceAccess) + @(@{ id = $roleId; type = 'Role' })
            }
        }
        Invoke-TeamsLinuxGraph -Method PATCH -Path "applications/$($application.id)" -Body @{
            keyCredentials = $keys; requiredResourceAccess = $permissions
        } | Out-Null
    }
    finally { $publicCertificate.Dispose() }
    $principalFilter = [uri]::EscapeDataString("appId eq '$($application.appId)'")
    $principals = @(Get-TeamsLinuxGraphCollection -Path "servicePrincipals?`$filter=$principalFilter")
    if ($principals.Count -gt 1) { throw "Multiple service principals for app $($application.appId)." }
    $principal = if ($principals.Count -eq 1) { $principals[0] } else {
        Invoke-TeamsLinuxGraph -Method POST -Path 'servicePrincipals' -Body @{ appId = $application.appId }
    }
    $assignments = @(Get-TeamsLinuxGraphCollection -Path "servicePrincipals/$($principal.id)/appRoleAssignments")
    foreach ($roleId in $roleIds) {
        if (-not ($assignments | Where-Object { $_.resourceId -eq $graphResource.id -and $_.appRoleId -eq $roleId })) {
            Invoke-TeamsLinuxGraph -Method POST -Path "servicePrincipals/$($principal.id)/appRoleAssignments" -Body @{
                principalId = $principal.id; resourceId = $graphResource.id; appRoleId = $roleId
            } | Out-Null
        }
    }
    $directoryFilter = [uri]::EscapeDataString("principalId eq '$($principal.id)'")
    $directoryAssignments = @(Get-TeamsLinuxGraphCollection -Path "roleManagement/directory/roleAssignments?`$filter=$directoryFilter")
    if (-not ($directoryAssignments | Where-Object { $_.roleDefinitionId -eq $roleDefinitions[0].id -and $_.directoryScopeId -eq '/' })) {
        Invoke-TeamsLinuxGraph -Method POST -Path 'roleManagement/directory/roleAssignments' -Body @{
            principalId = $principal.id; roleDefinitionId = $roleDefinitions[0].id; directoryScopeId = '/'
        } | Out-Null
    }
    return @{ AppId = $application.appId }
}

if (-not $IsLinux) { throw 'This script requires Linux and PowerShell 7.4 or later. Use o365-connect-tms-cert.ps1 on Windows.' }
if ($GenerateLocalCertificate -eq $UseCertificateAuth) { throw 'Specify exactly one of -GenerateLocalCertificate or -UseCertificateAuth.' }
if ($ProvisionEntraApp -and (-not $GenerateLocalCertificate -or $noprompt)) {
    throw '-ProvisionEntraApp requires -GenerateLocalCertificate and interactive administrator sign-in (no -noprompt).'
}
if ($ProvisionEntraApp -and -not $Tenant) { throw '-ProvisionEntraApp requires -Tenant.' }
$dataRoot = if ($env:XDG_DATA_HOME -and [IO.Path]::IsPathRooted($env:XDG_DATA_HOME)) { $env:XDG_DATA_HOME } else { Join-Path $HOME '.local/share' }
$configRoot = if ($env:XDG_CONFIG_HOME -and [IO.Path]::IsPathRooted($env:XDG_CONFIG_HOME)) { $env:XDG_CONFIG_HOME } else { Join-Path $HOME '.config' }
$dataDirectory = Join-Path $dataRoot 'ciaops/tms'
if (-not $GeneratedCertOutputPath) { $GeneratedCertOutputPath = $dataDirectory }
if (-not $CertificateMapPath) { $CertificateMapPath = Join-Path $configRoot 'ciaops/tms/o365-tms-cert-auth.json' }
$GeneratedCertOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($GeneratedCertOutputPath)
$CertificateMapPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CertificateMapPath)
$transcribing = $false
$localCertificate = $null
$teamsConnected = $false
try {
    if ($enableLog) {
        $logDirectory = Initialize-TeamsPrivateDirectory -Path $dataDirectory
        $logPath = Join-Path $logDirectory ("o365-connect-tms-{0}.txt" -f [guid]::NewGuid())
        Write-TeamsPrivateFile -Path $logPath -Bytes ([byte[]]::new(0))
        Start-Transcript -LiteralPath $logPath -Append -ErrorAction Stop | Out-Null
        $transcribing = $true
    }
    if ($GenerateLocalCertificate) {
        $password = Get-TeamsPfxPassword -Password $GeneratedPfxPassword -NoPrompt:$noprompt -EnvironmentVariable $CertificatePasswordEnvironmentVariable
        $generated = New-TeamsLinuxCertificate -SubjectName $GeneratedCertSubject -YearsValid $GeneratedCertYearsValid -OutputPath $GeneratedCertOutputPath -Password $password
        Write-Host "Public certificate: $($generated.CerPath)"
        Write-Host "Encrypted private key: $($generated.PfxPath)"
        Write-Host "Thumbprint: $($generated.Thumbprint); expires: $($generated.NotAfter.ToString('yyyy-MM-dd'))"
        if (-not $ProvisionEntraApp) { return }
        Write-Warning 'Provisioning grants Teams Administrator and Graph write access to groups, Teams, channels and the app catalog.'
        Import-TeamsLinuxModule -Name Microsoft.Graph.Authentication -MinimumVersion 2.0.0
        $graphParameters = @{
            TenantId = $Tenant; UseDeviceAuthentication = $true; ContextScope = 'Process'; NoWelcome = $true
            Scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'RoleManagement.ReadWrite.Directory')
            ErrorAction = 'Stop'
        }
        if ($SetupClientId) { $graphParameters.ClientId = $SetupClientId }
        Connect-MgGraph @graphParameters
        try {
            $displayName = if ($AppDisplayName) { $AppDisplayName } else { $GeneratedCertSubject }
            $provisioned = New-TeamsLinuxEntraApp -DisplayName $displayName -CerPath $generated.CerPath
            $entryName = if ($ProfileName) { $ProfileName } else { $displayName }
            Save-TeamsLinuxProfile -Path $CertificateMapPath -Entry @{
                name = $entryName; tenant = $Tenant; organization = $Tenant; appId = $provisioned.AppId
                certificateThumbprint = $generated.Thumbprint; certificateFilePath = $generated.PfxPath
            }
            Write-Host "Provisioned app: $($provisioned.AppId). Profile saved: $CertificateMapPath"
            Write-Warning 'New app permissions and directory roles may take time to replicate. Connect from a fresh pwsh session.'
        }
        finally { Disconnect-MgGraph -ErrorAction Stop | Out-Null }
        return
    }

    if ($ProfileName -or -not ($Tenant -and $AppId -and $CertificateFilePath)) {
        $profiles = @(Read-TeamsProfiles -Path $CertificateMapPath | Where-Object {
            (-not $ProfileName -or $_.name -eq $ProfileName) -and
            (-not $Tenant -or $_.tenant -eq $Tenant -or $_.organization -eq $Tenant)
        })
        if ($profiles.Count -ne 1) { throw "Expected one matching profile in '$CertificateMapPath', found $($profiles.Count). Specify -ProfileName or explicit connection parameters." }
        $selectedProfile = $profiles[0]
        if (-not $Tenant) { $Tenant = if ($selectedProfile.organization) { $selectedProfile.organization } else { $selectedProfile.tenant } }
        if (-not $AppId) { $AppId = $selectedProfile.appId }
        if (-not $CertificateFilePath) {
            $CertificateFilePath = $selectedProfile.certificateFilePath
            if ($CertificateFilePath -and -not [IO.Path]::IsPathRooted($CertificateFilePath)) {
                $CertificateFilePath = Join-Path (Split-Path -Parent $CertificateMapPath) $CertificateFilePath
            }
        }
        if (-not $CertificateThumbprint) { $CertificateThumbprint = $selectedProfile.certificateThumbprint }
    }
    if (-not ($Tenant -and $AppId -and $CertificateFilePath)) {
        throw 'Connection requires Tenant, AppId and CertificateFilePath. Windows thumbprint-only profiles need a certificateFilePath pointing to an exported PFX.'
    }
    $password = Get-TeamsPfxPassword -Password $CertificatePassword -NoPrompt:$noprompt -EnvironmentVariable $CertificatePasswordEnvironmentVariable
    $localCertificate = Import-TeamsLinuxPfx -Path $CertificateFilePath -Password $password
    if (-not $localCertificate.HasPrivateKey) { throw 'The certificate has no private key. Supply a PFX, not a public CER.' }
    if ($CertificateThumbprint -and $localCertificate.Thumbprint -ne ($CertificateThumbprint -replace '[^0-9a-fA-F]', '')) {
        throw 'The PFX thumbprint does not match the expected certificate thumbprint.'
    }
    if ($localCertificate.NotBefore -gt (Get-Date) -or $localCertificate.NotAfter -le (Get-Date)) { throw 'The certificate is not currently valid.' }
    if ($localCertificate.NotAfter -lt (Get-Date).AddDays(30)) { Write-Warning 'The certificate expires in fewer than 30 days.' }
    Import-TeamsLinuxModule -Name MicrosoftTeams -MinimumVersion 5.0.0 -CheckUpdate:(-not $noupdate)
    Connect-MicrosoftTeams -TenantId $Tenant -ApplicationId $AppId -Certificate $localCertificate -ErrorAction Stop | Out-Null
    $teamsConnected = $true
    Write-Host "Connected to Teams tenant $Tenant using app $AppId and certificate $($localCertificate.Thumbprint)."
}
finally {
    # Teams may retain the certificate for token renewal in the live session.
    if ($null -ne $localCertificate -and -not $teamsConnected) { $localCertificate.Dispose() }
    if ($transcribing) { Stop-Transcript | Out-Null }
}
