<#
.SYNOPSIS
    Connect to the SharePoint Online admin center using certificates on Linux.
.DESCRIPTION
    Linux admin-center entry point for o365-connect-pnp-cert-linux.ps1, which
    must be present beside this script. Requires Linux and PowerShell 7.4.6+.
    Uses PnP.PowerShell 3.0.0+ rather than the Windows-only SharePoint Online
    Management Shell. The resulting session supports PnP cmdlets such as
    Get-PnPTenantSite, NOT Get-SPOSite or other SPO Management Shell cmdlets.

    Select exactly one mode: -GenerateLocalCertificate creates an RSA-2048
    certificate, public CER and password-protected PFX; -UseCertificateAuth
    connects in the current PowerShell process. No sudo or certificate store
    is needed. Missing modules are installed for CurrentUser unless headless.

    -ProvisionEntraApp with generation creates/reuses an app, uploads the public
    certificate, and grants SharePoint Sites.FullControl.All and Graph
    Application.Read.All using interactive administrator device-code sign-in.
    Unlike the Windows SPO script, Graph Group.Read.All is not requested.

    Certificates default to $XDG_DATA_HOME/ciaops/spo (or ~/.local/share/ciaops/spo).
    Profiles default to $XDG_CONFIG_HOME/ciaops/spo/o365-spo-admin-cert-auth.json
    (or ~/.config/ciaops/spo/o365-spo-admin-cert-auth.json). The shared template
    restricts dedicated storage directories to 700 and files to 600 and rejects
    symbolic links in storage paths. Custom directories must be dedicated too.
    PFX files supplied for connection must have owner-only permissions.

    Profiles accept adminUrl (Windows SPO format) or siteUrl (Linux PnP format).
    New profiles use siteUrl to store the admin URL. Windows thumbprint-only
    profiles also need -CertificateFilePath or a certificateFilePath property.
    Passwords are never saved in profiles. Other certificate, logging and module
    options are forwarded unchanged to the Linux PnP template; see its help.
.PARAMETER AdminUrl
    HTTPS SharePoint admin-center URL. Alias: SiteUrl. When no profile supplies
    it, derived from a Tenant of the form contoso.onmicrosoft.com. Specify it
    explicitly for tenant GUIDs, custom domains or renamed SharePoint tenants.
.PARAMETER Tenant
    Tenant ID or domain. Also filters saved profiles when supplied.
.PARAMETER AppId
    Application client ID. With Tenant and CertificateFilePath, connects without
    reading profiles unless ProfileName is supplied. AdminUrl may be derived
    from an onmicrosoft.com tenant domain.
.PARAMETER ProfileName
    Saved profile to select. Multiple matches are errors, including interactively.
.PARAMETER CertificateMapPath
    JSON profile map, either an array or an object containing profiles.
.PARAMETER CertificateFilePath
    Encrypted PFX file. Overrides the selected profile's certificateFilePath.
.PARAMETER CertificatePassword
    SecureString PFX password. Prompted for if omitted, unless headless.
.PARAMETER CertificatePasswordEnvironmentVariable
    Name of an injected process environment variable containing the PFX password.
    Used when the relevant SecureString parameter is omitted. Prefer a vault;
    process environment values may be accessible to other same-user processes.
.PARAMETER noprompt
    Alias: Headless. Requires preinstalled modules and a supplied password.
    Cannot be combined with ProvisionEntraApp.
.EXAMPLE
    ./o365-connect-spo-cert-linux.ps1 -GenerateLocalCertificate -ProvisionEntraApp -Tenant contoso.onmicrosoft.com
    Generate an encrypted PFX, provision the application and save an admin profile.
.EXAMPLE
    ./o365-connect-spo-cert-linux.ps1 -UseCertificateAuth -ProfileName O365-SPO-AppAuth
    Connect using the saved profile and prompt for its PFX password.
.EXAMPLE
    ./o365-connect-spo-cert-linux.ps1 -UseCertificateAuth -Tenant contoso.onmicrosoft.com -AdminUrl https://contoso-admin.sharepoint.com -AppId 11111111-2222-3333-4444-555555555555 -CertificateFilePath /home/admin/auth.pfx -CertificatePassword $vaultPassword -Headless
    Connect unattended using an existing application and encrypted PFX.
.NOTES
    CIAOPS - Script provided as is. Use at own risk. No guarantees or warranty provided.
.LINK
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
    [string]$GeneratedCertSubject = 'O365-SPO-AppAuth',
    [ValidateRange(1, 100)][int]$GeneratedCertYearsValid = 2,
    [string]$GeneratedCertOutputPath,
    [securestring]$GeneratedPfxPassword,
    [string]$CertificateFilePath,
    [securestring]$CertificatePassword,
    [string]$CertificatePasswordEnvironmentVariable,
    [string]$CertificateThumbprint,
    [string]$CertificateMapPath,
    [string]$Tenant,
    [Alias('SiteUrl')][string]$AdminUrl,
    [string]$AppId,
    [string]$ProfileName,
    [switch]$ProvisionEntraApp,
    [string]$AppDisplayName,
    [string]$SetupClientId
)

if (-not $IsLinux) { throw 'This script requires Linux. Use o365-connect-spo-cert.ps1 on Windows.' }
if ($GenerateLocalCertificate -eq $UseCertificateAuth) { throw 'Specify exactly one of -GenerateLocalCertificate or -UseCertificateAuth.' }
if ($ProvisionEntraApp -and (-not $GenerateLocalCertificate -or $noprompt)) {
    throw '-ProvisionEntraApp requires -GenerateLocalCertificate and interactive administrator sign-in (no -noprompt).'
}
$templatePath = Join-Path $PSScriptRoot 'o365-connect-pnp-cert-linux.ps1'
if (-not (Test-Path -LiteralPath $templatePath -PathType Leaf)) { throw "Required Linux PnP template is missing: $templatePath" }
$dataRoot = if ($env:XDG_DATA_HOME -and [IO.Path]::IsPathRooted($env:XDG_DATA_HOME)) { $env:XDG_DATA_HOME } else { Join-Path $HOME '.local/share' }
$configRoot = if ($env:XDG_CONFIG_HOME -and [IO.Path]::IsPathRooted($env:XDG_CONFIG_HOME)) { $env:XDG_CONFIG_HOME } else { Join-Path $HOME '.config' }
if (-not $GeneratedCertOutputPath) { $GeneratedCertOutputPath = Join-Path $dataRoot 'ciaops/spo' }
if (-not $CertificateMapPath) { $CertificateMapPath = Join-Path $configRoot 'ciaops/spo/o365-spo-admin-cert-auth.json' }
$CertificateMapPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CertificateMapPath)

if ($UseCertificateAuth -and -not $ProfileName -and -not $AdminUrl -and $AppId -and $CertificateFilePath -and
    $Tenant -match '^([a-z0-9-]+)\.onmicrosoft\.com$') {
    $AdminUrl = "https://$($Matches[1])-admin.sharepoint.com"
}

if ($UseCertificateAuth -and ($ProfileName -or -not ($Tenant -and $AdminUrl -and $AppId -and $CertificateFilePath)) -and
    (Test-Path -LiteralPath $CertificateMapPath)) {
    $map = Get-Content -LiteralPath $CertificateMapPath -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -NoEnumerate -ErrorAction Stop
    $profiles = if ($map -is [System.Collections.IDictionary] -and $map.Contains('profiles')) { @($map.profiles) }
        elseif ($map -is [array]) { $map }
        else { throw "Invalid profile map: $CertificateMapPath. Expected an array or an object containing profiles." }
    $profiles = @($profiles | Where-Object {
        $profileUrl = if ($_.adminUrl) { $_.adminUrl } else { $_.siteUrl }
        (-not $ProfileName -or $_.name -eq $ProfileName) -and
        (-not $Tenant -or $_.tenant -eq $Tenant) -and
        (-not $AdminUrl -or $profileUrl -eq $AdminUrl)
    })
    if ($profiles.Count -ne 1) { throw "Expected one matching profile in '$CertificateMapPath', found $($profiles.Count). Specify -ProfileName or explicit connection parameters." }
    $selectedProfile = $profiles[0]
    if (-not $Tenant) { $Tenant = $selectedProfile.tenant }
    if (-not $AdminUrl) { $AdminUrl = if ($selectedProfile.adminUrl) { $selectedProfile.adminUrl } else { $selectedProfile.siteUrl } }
    if (-not $AppId) { $AppId = $selectedProfile.appId }
    if (-not $CertificateThumbprint) { $CertificateThumbprint = $selectedProfile.certificateThumbprint }
    if (-not $CertificateFilePath) {
        $CertificateFilePath = $selectedProfile.certificateFilePath
        if ($CertificateFilePath -and -not [IO.Path]::IsPathRooted($CertificateFilePath)) {
            $CertificateFilePath = Join-Path (Split-Path -Parent $CertificateMapPath) $CertificateFilePath
        }
    }
}
elseif ($UseCertificateAuth -and $ProfileName) { throw "Profile map not found: $CertificateMapPath" }

if (-not $AdminUrl -and $Tenant -match '^([a-z0-9-]+)\.onmicrosoft\.com$') {
    $AdminUrl = "https://$($Matches[1])-admin.sharepoint.com"
}
if ($AdminUrl) {
    $adminUri = $null
    if (-not [uri]::TryCreate($AdminUrl, [UriKind]::Absolute, [ref]$adminUri) -or
        $adminUri.Scheme -ne 'https' -or $adminUri.Host -notmatch '^[a-z0-9-]+-admin\.sharepoint\.(com|us|de|cn)$' -or
        $adminUri.AbsolutePath -ne '/' -or $adminUri.Query -or $adminUri.Fragment -or $adminUri.UserInfo -or -not $adminUri.IsDefaultPort) {
        throw '-AdminUrl must be an HTTPS SharePoint admin-center URL, for example https://contoso-admin.sharepoint.com.'
    }
}
if (($UseCertificateAuth -or $ProvisionEntraApp) -and -not $AdminUrl) {
    throw 'Specify -AdminUrl or select a profile containing adminUrl/siteUrl. Automatic derivation requires a tenant.onmicrosoft.com domain.'
}
if ($UseCertificateAuth -and -not ($Tenant -and $AppId -and $CertificateFilePath)) {
    throw 'Connection requires Tenant, AdminUrl, AppId and CertificateFilePath. Windows thumbprint-only profiles need an exported PFX.'
}

$connectionParameters = @{}
foreach ($parameterName in $PSBoundParameters.Keys) {
    if ($parameterName -notin @('AdminUrl', 'ProfileName')) { $connectionParameters[$parameterName] = $PSBoundParameters[$parameterName] }
}
$connectionParameters.GeneratedCertSubject = $GeneratedCertSubject
$connectionParameters.GeneratedCertOutputPath = $GeneratedCertOutputPath
$connectionParameters.CertificateMapPath = $CertificateMapPath
if ($Tenant) { $connectionParameters.Tenant = $Tenant }
if ($AdminUrl) { $connectionParameters.SiteUrl = $AdminUrl }
if ($AppId) { $connectionParameters.AppId = $AppId }
if ($CertificateFilePath) { $connectionParameters.CertificateFilePath = $CertificateFilePath }
if ($CertificateThumbprint) { $connectionParameters.CertificateThumbprint = $CertificateThumbprint }
if ($GenerateLocalCertificate -and $ProfileName) { $connectionParameters.ProfileName = $ProfileName }
& $templatePath @connectionParameters