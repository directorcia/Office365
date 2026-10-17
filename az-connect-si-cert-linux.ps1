<#
.SYNOPSIS
    Connect to Microsoft Sentinel using a certificate on Linux.
.DESCRIPTION
    Certificate-based counterpart of az-connect-si.ps1 for PowerShell 7.4+.
    Loads only Az.Accounts and Az.SecurityInsights, not the full Az module.
    One-time setup: use the existing SecurityMonitor app or a new Entra app.
    Assign its service principal the Microsoft Sentinel Responder role on the
    Sentinel workspace or its resource group to manage incidents. Register the
    public certificate on the app and keep the matching encrypted PFX, including
    its private key, on the computer running this script. This script does not
    create apps, certificates, or role assignments.

    Uses an encrypted PFX file, without importing it into a certificate store.
    Azure context autosave is disabled for this process. Run this script in the
    same PowerShell session as the Sentinel commands; exiting pwsh ends the session.
    The current process's Azure context is replaced, not other saved contexts.

    For profile-based runs, first connect once with -SaveProfile -ProfileName and
    explicit -Tenant, -AppId, -SubscriptionId and -CertificateFilePath values,
    as shown in the first example. A saved profile is optional when all four
    connection values are supplied explicitly on every run.

    Profiles default to $XDG_CONFIG_HOME/ciaops/sentinel/az-si-cert-auth.json
    or ~/.config/ciaops/sentinel/az-si-cert-auth.json. Passwords are never saved.
    On Linux the PFX must have owner-only permissions (chmod 600). Profile storage
    directories must be dedicated: saving restricts them to 700 and files to 600.
    Symbolic links are not accepted in profile storage or for the PFX itself.
.PARAMETER UseCertificateAuth
    Certificate authentication (the default and only sign-in mode).
.PARAMETER Tenant
    Entra tenant ID or domain. Alias: TenantId. Overrides the selected profile.
.PARAMETER AppId
    Service principal application/client ID. Alias: ApplicationId.
.PARAMETER SubscriptionId
    Target subscription GUID. Overrides the selected profile; no subscription picker.
.PARAMETER CertificateFilePath
    Encrypted PFX containing the private key. Overrides the selected profile.
.PARAMETER CertificatePassword
    SecureString PFX password, preferably retrieved from a secret vault.
.PARAMETER CertificatePasswordEnvironmentVariable
    Name of a process environment variable containing the PFX password. Used only
    if CertificatePassword is omitted. Inject it through your scheduler or secret
    manager. Environment variables are not a vault and may be accessible to other
    processes running as the same user. The value is never printed or saved.
.PARAMETER CertificateThumbprint
    Optional expected PFX thumbprint; also checked when loaded from a profile.
.PARAMETER ProfileName
    Select a saved profile. Without a name, exactly one matching profile is needed
    when explicit connection details are incomplete.
.PARAMETER CertificateMapPath
    JSON profile map: an array or an object containing a profiles array. Relative
    certificate paths in profiles resolve relative to this file.
.PARAMETER SaveProfile
    Save the successful connection's non-secret details. Requires ProfileName.
.PARAMETER noprompt
    Alias: Headless. Never prompts, installs, updates, or queries the Gallery.
    Requires preinstalled modules and a password parameter or environment variable.
.PARAMETER noupdate
    Skip Gallery update checks. Headless always skips them. Interactive module
    installation and updates use CurrentUser and confirmation, without sudo.
.EXAMPLE
    ./az-connect-si-cert-linux.ps1 -UseCertificateAuth -Tenant contoso.onmicrosoft.com -AppId 11111111-2222-3333-4444-555555555555 -SubscriptionId aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee -CertificateFilePath /home/admin/auth.pfx -ProfileName sentinel -SaveProfile -noupdate
    One-time profile setup after configuring the app, certificate, and RBAC role.
    Prompt for the PFX password, connect, and save non-secret connection details.
.EXAMPLE
    ./az-connect-si-cert-linux.ps1 -UseCertificateAuth -ProfileName sentinel -Headless -CertificatePasswordEnvironmentVariable SENTINEL_PFX_PASSWORD
    Connect in the current PowerShell process using a saved profile and a
    scheduler-injected password. Run the incident-close script afterward in this
    same process. For scheduling, use pwsh -NoProfile -NonInteractive -File with a
    wrapper script that invokes both scripts; running this connector alone with
    pwsh -File exits immediately after connecting and loses the process context.
.NOTES
    Authentication reference: https://learn.microsoft.com/powershell/module/az.accounts/connect-azaccount
    CIAOPS - Script provided as is. Use at own risk. No guarantees or warranty provided.
#>
#requires -Version 7.4
[CmdletBinding()]
param(
    [switch]$UseCertificateAuth = $true,
    [Alias('Headless')][switch]$noprompt,
    [switch]$noupdate,
    [Alias('TenantId')][string]$Tenant,
    [Alias('ApplicationId')][string]$AppId,
    [string]$SubscriptionId,
    [string]$CertificateFilePath,
    [securestring]$CertificatePassword,
    [string]$CertificatePasswordEnvironmentVariable,
    [string]$CertificateThumbprint,
    [string]$ProfileName,
    [string]$CertificateMapPath,
    [switch]$SaveProfile
)

function Read-SentinelProfiles {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $map = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -ErrorAction Stop
    if ($map -is [System.Collections.IDictionary] -and $map.Contains('profiles') -and $map.profiles -is [array]) {
        return $map.profiles
    }
    if ($map -is [array]) { return $map }
    throw "Invalid profile map: $Path. Expected an array or an object containing a profiles array."
}

function Save-SentinelProfile {
    param([string]$Path, [hashtable]$Entry)
    $directory = Split-Path -Parent $Path
    $ancestor = $Path
    while ($ancestor) {
        if ((Test-Path -LiteralPath $ancestor) -and
            ([IO.File]::GetAttributes($ancestor) -band [IO.FileAttributes]::ReparsePoint)) {
            throw "Symbolic links are not supported in profile storage: $ancestor"
        }
        $ancestor = [IO.Path]::GetDirectoryName($ancestor)
    }
    $fileMode = [IO.UnixFileMode]::UserRead -bor [IO.UnixFileMode]::UserWrite
    if ($IsLinux) {
        [IO.Directory]::CreateDirectory($directory, $fileMode -bor [IO.UnixFileMode]::UserExecute) | Out-Null
        [IO.File]::SetUnixFileMode($directory, $fileMode -bor [IO.UnixFileMode]::UserExecute)
    }
    else { [IO.Directory]::CreateDirectory($directory) | Out-Null }
    $lockOptions = [IO.FileStreamOptions]::new()
    $lockOptions.Mode = [IO.FileMode]::OpenOrCreate
    $lockOptions.Access = [IO.FileAccess]::ReadWrite
    $lockOptions.Share = [IO.FileShare]::None
    if ($IsLinux) { $lockOptions.UnixCreateMode = $fileMode }
    $lockStream = [IO.FileStream]::new((Join-Path $directory '.profiles.lock'), $lockOptions)
    $temporaryPath = Join-Path $directory ([guid]::NewGuid().ToString() + '.tmp')
    try {
        $profiles = @(Read-SentinelProfiles -Path $Path | Where-Object { $_.name -ne $Entry.name }) + @($Entry)
        $bytes = [Text.Encoding]::UTF8.GetBytes((@{ profiles = $profiles } | ConvertTo-Json -Depth 10))
        $options = [IO.FileStreamOptions]::new()
        $options.Mode = [IO.FileMode]::CreateNew
        $options.Access = [IO.FileAccess]::Write
        $options.Share = [IO.FileShare]::None
        if ($IsLinux) { $options.UnixCreateMode = $fileMode }
        $stream = [IO.FileStream]::new($temporaryPath, $options)
        try { $stream.Write($bytes, 0, $bytes.Length) }
        finally { $stream.Dispose() }
        [IO.File]::Move($temporaryPath, $Path, $true)
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force }
        $lockStream.Dispose()
    }
}

function Import-SentinelModule {
    param([string]$Name)
    $available = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if ($null -eq $available) {
        if ($noprompt) { throw "Required module '$Name' is missing. Install it before running headless: Install-Module -Name $Name -Scope CurrentUser -Repository PSGallery" }
        if ((Read-Host "Install $Name for CurrentUser? [y/N]") -ne 'y') { throw "Required module '$Name' was not installed." }
        Install-Module -Name $Name -Scope CurrentUser -Repository PSGallery -Force -Confirm:$false -ErrorAction Stop
    }
    elseif (-not $noprompt -and -not $noupdate) {
        try {
            $latest = Find-Module -Name $Name -Repository PSGallery -ErrorAction Stop
            if ([version]$latest.Version -gt $available.Version -and (Read-Host "Install $Name $($latest.Version)? [y/N]") -eq 'y') {
                Install-Module -Name $Name -Scope CurrentUser -Repository PSGallery -Force -Confirm:$false -ErrorAction Stop
            }
        }
        catch { Write-Warning "Module update check failed for ${Name}: $($_.Exception.Message)" }
    }
    Import-Module -Name $Name -Global -ErrorAction Stop
}

if (-not $UseCertificateAuth) { throw 'This script supports certificate authentication only.' }
if ($SaveProfile -and -not $ProfileName) { throw '-SaveProfile requires -ProfileName.' }
if (-not $CertificateMapPath) {
    $configRoot = if ($env:XDG_CONFIG_HOME) { $env:XDG_CONFIG_HOME } else { Join-Path $HOME '.config' }
    $CertificateMapPath = Join-Path $configRoot 'ciaops/sentinel/az-si-cert-auth.json'
}
$CertificateMapPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CertificateMapPath)

if (($ProfileName -and -not $SaveProfile) -or -not ($Tenant -and $AppId -and $SubscriptionId -and $CertificateFilePath)) {
    $profiles = @(Read-SentinelProfiles -Path $CertificateMapPath | Where-Object {
        ($_ -is [System.Collections.IDictionary]) -and
        (-not $ProfileName -or $_.name -eq $ProfileName) -and
        (-not $Tenant -or $ProfileName -or $_.tenant -eq $Tenant) -and
        (-not $AppId -or $ProfileName -or $_.appId -eq $AppId) -and
        (-not $SubscriptionId -or $ProfileName -or $_.subscriptionId -eq $SubscriptionId)
    })
    if ($profiles.Count -ne 1) {
        throw "Expected one matching profile in '$CertificateMapPath', found $($profiles.Count). Specify -ProfileName or explicit -Tenant, -AppId, -SubscriptionId and -CertificateFilePath."
    }
    $selectedProfile = $profiles[0]
    if (-not $Tenant) { $Tenant = $selectedProfile.tenant }
    if (-not $AppId) { $AppId = $selectedProfile.appId }
    if (-not $SubscriptionId) { $SubscriptionId = $selectedProfile.subscriptionId }
    if (-not $CertificateFilePath) {
        $CertificateFilePath = $selectedProfile.certificateFilePath
        if ($CertificateFilePath -and -not [IO.Path]::IsPathRooted($CertificateFilePath)) {
            $CertificateFilePath = Join-Path (Split-Path -Parent $CertificateMapPath) $CertificateFilePath
        }
        if (-not $CertificateThumbprint) { $CertificateThumbprint = $selectedProfile.certificateThumbprint }
    }
}
if (-not ($Tenant -and $AppId -and $SubscriptionId -and $CertificateFilePath)) {
    throw 'Connection requires Tenant, AppId, SubscriptionId and CertificateFilePath, explicitly or from a saved profile.'
}
$parsedId = [guid]::Empty
if (-not [guid]::TryParse($AppId, [ref]$parsedId)) { throw 'AppId must be an application/client GUID.' }
if (-not [guid]::TryParse($SubscriptionId, [ref]$parsedId)) { throw 'SubscriptionId must be a subscription GUID.' }

if ($null -eq $CertificatePassword -and $CertificatePasswordEnvironmentVariable) {
    $environmentPassword = [Environment]::GetEnvironmentVariable($CertificatePasswordEnvironmentVariable, 'Process')
    if ([string]::IsNullOrEmpty($environmentPassword)) {
        throw "PFX password environment variable '$CertificatePasswordEnvironmentVariable' is missing or empty."
    }
    try { $CertificatePassword = ConvertTo-SecureString -String $environmentPassword -AsPlainText -Force }
    finally { $environmentPassword = $null }
}
if ($null -eq $CertificatePassword) {
    if ($noprompt) { throw 'Headless mode requires -CertificatePassword (SecureString) or -CertificatePasswordEnvironmentVariable.' }
    $CertificatePassword = Read-Host -Prompt 'PFX password' -AsSecureString
}
if ($CertificatePassword.Length -eq 0) { throw 'An empty PFX password is not permitted.' }
$CertificateFilePath = (Resolve-Path -LiteralPath $CertificateFilePath -ErrorAction Stop).ProviderPath
if ([IO.File]::GetAttributes($CertificateFilePath) -band [IO.FileAttributes]::ReparsePoint) {
    throw 'The PFX must not be a symbolic link.'
}
if ($IsLinux) {
    $allowedMode = [IO.UnixFileMode]::UserRead -bor [IO.UnixFileMode]::UserWrite
    if (([IO.File]::GetUnixFileMode($CertificateFilePath) -band (-bnot $allowedMode)) -ne 0) {
        throw 'PFX permissions must be owner-only. Restrict the file with chmod 600 before connecting.'
    }
}
$localCertificate = $null
try {
    $localCertificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $CertificateFilePath, $CertificatePassword, [Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet)
    if (-not $localCertificate.HasPrivateKey) { throw 'The certificate has no private key. Supply a PFX, not a public CER.' }
    if ($CertificateThumbprint -and $localCertificate.Thumbprint -ne ($CertificateThumbprint -replace '[^0-9a-fA-F]', '')) {
        throw 'The PFX thumbprint does not match the expected certificate thumbprint.'
    }
    if ($localCertificate.NotBefore -gt (Get-Date) -or $localCertificate.NotAfter -le (Get-Date)) { throw 'The certificate is not currently valid.' }
    if ($localCertificate.NotAfter -lt (Get-Date).AddDays(30)) { Write-Warning 'The certificate expires in fewer than 30 days.' }

    Import-SentinelModule -Name Az.Accounts
    Import-SentinelModule -Name Az.SecurityInsights
    Disable-AzContextAutosave -Scope Process -ErrorAction Stop | Out-Null
    Clear-AzContext -Scope Process -Force -ErrorAction Stop | Out-Null
    try {
        Connect-AzAccount -ServicePrincipal -ApplicationId $AppId -Tenant $Tenant -Subscription $SubscriptionId `
            -CertificatePath $CertificateFilePath -CertificatePassword $CertificatePassword `
            -SkipContextPopulation -Scope Process -ErrorAction Stop | Out-Null
        $context = Set-AzContext -SubscriptionId $SubscriptionId -Tenant $Tenant -Scope Process -ErrorAction Stop
        if ($context.Subscription.Id -ne $SubscriptionId) { throw 'Azure returned an unexpected subscription context.' }
    }
    catch {
        Clear-AzContext -Scope Process -Force -ErrorAction SilentlyContinue | Out-Null
        throw "Unable to connect to Azure subscription '${SubscriptionId}' using certificate authentication: $($_.Exception.Message)"
    }
    if ($SaveProfile) {
        Save-SentinelProfile -Path $CertificateMapPath -Entry @{
            name = $ProfileName; tenant = $Tenant; appId = $AppId; subscriptionId = $SubscriptionId
            certificateFilePath = $CertificateFilePath; certificateThumbprint = $localCertificate.Thumbprint
        }
        Write-Host "Profile saved: $CertificateMapPath"
    }
    Write-Host "Connected to Azure subscription $SubscriptionId using app $AppId. Az.SecurityInsights is loaded."
}
finally {
    if ($null -ne $localCertificate) { $localCertificate.Dispose() }
}