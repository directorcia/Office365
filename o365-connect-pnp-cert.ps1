param(
    [switch]$noprompt = $false,   ## if -noprompt used then user will not be asked for any input
    [switch]$noupdate = $false,   ## if -noupdate used then module will not be checked for more recent version
    [switch]$enableLog = $false,  ## if -enableLog create a transcript log file

    [switch]$GenerateLocalCertificate = $false,
    [switch]$UseCertificateAuth = $false,

    [string]$GeneratedCertSubject = "O365-PNP-AppAuth",
    [int]$GeneratedCertYearsValid = 2,
    [string]$GeneratedCertOutputPath = "",  ## defaults to parent of script directory at runtime
    [switch]$ExportGeneratedPfx = $false,
    [securestring]$GeneratedPfxPassword,

    [string]$Tenant,
    [string]$ProfileName,
    [string]$SiteUrl,           ## SharePoint site URL to connect to (e.g. https://contoso.sharepoint.com)
    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$CertificateMapPath = "",  ## defaults to o365-pnp-cert-auth.json in parent of script directory

    ## When used with -GenerateLocalCertificate, also create the Entra app, upload the cert,
    ## grant Sites.FullControl.All, and update the profile map automatically.
    [switch]$ProvisionEntraApp = $false,
    [string]$AppDisplayName = "",
    ## Client ID of a public client app registered in your tenant for device-code Graph auth.
    ## Defaults to the well-known Azure PowerShell public client.
    [string]$SetupClientId = "1950a258-227b-4e31-a9cf-717495945fc2",
    ## Opt-in only. Clipboard copy can leak auth codes in shared/RDP sessions.
    [switch]$CopyDeviceCodeToClipboard = $false
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.
Description - SharePoint Online connect script with two modes:
1. GenerateLocalCertificate: create/export local cert files.
2. UseCertificateAuth: connect to SharePoint Online with existing app/cert.
Usage - For setup and execution examples, see:
https://github.com/directorcia/Office365/wiki/Connect-to-SharePoint-Online-with-Certificates
Source - https://github.com/directorcia/Office365/blob/master/o365-connect-pnp-cert.ps1
#>

## Resolve paths relative to the script file itself, not the caller's working directory.
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptParentDir = Split-Path -Parent $scriptDir
if ([string]::IsNullOrWhiteSpace($GeneratedCertOutputPath)) { $GeneratedCertOutputPath = $scriptParentDir }
if ([string]::IsNullOrWhiteSpace($CertificateMapPath))      { $CertificateMapPath = Join-Path $scriptParentDir 'o365-pnp-cert-auth.json' }

## Shared output colors passed explicitly to functions to avoid hidden script-scope coupling.
$Colors = @{
    SystemMessage  = "cyan"
    ProcessMessage = "green"
    ErrorMessage   = "red"
    WarningMessage = "yellow"
}

## Well-known service principal app IDs used during provisioning.
$SpoResourceAppId   = "00000003-0000-0ff1-ce00-000000000000"
$GraphResourceAppId = "00000003-0000-0000-c000-000000000000"

## Tracks whether this run opened a cert-auth SPO session that should be closed on error.
$disconnectCertificateAuthOnError = $false

## Resolve the executable of the current host so elevated module installs target the same runtime
## (PS5 vs PS7) and module path as the running script. Get-Command pwsh/powershell would pick a
## different version, causing Install-Module to land in an unreachable module path.
$elevatedShellPath = (Get-Process -Id $PID).MainModule.FileName
if ([string]::IsNullOrWhiteSpace($elevatedShellPath) -or -not (Test-Path -LiteralPath $elevatedShellPath)) {
    throw "Unable to resolve current PowerShell host executable path for elevated module operations (PID $PID)."
}

function Get-ScriptInvocationArguments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,
        [Parameter(Mandatory = $true)]
        [hashtable]$BoundParameters,
        [Parameter(Mandatory = $false)]
        [string[]]$UnboundArguments = @()
    )

    $childArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath)
    foreach ($paramName in $BoundParameters.Keys) {
        $paramValue = $BoundParameters[$paramName]
        if ($paramValue -is [switch]) {
            if ($paramValue.IsPresent) {
                $childArgs += "-$paramName"
            }
        }
        elseif ($null -ne $paramValue) {
            $childArgs += "-$paramName"
            if ($paramValue -is [System.Array]) {
                foreach ($arrayEntry in $paramValue) {
                    $childArgs += [string]$arrayEntry
                }
            }
            else {
                $childArgs += [string]$paramValue
            }
        }
    }

    if ($UnboundArguments.Count -gt 0) {
        $childArgs += $UnboundArguments
    }

    return $childArgs
}

$scriptInvocationArguments = Get-ScriptInvocationArguments -ScriptPath $PSCommandPath -BoundParameters $PSBoundParameters -UnboundArguments $args

if ($PSVersionTable.PSEdition -ne 'Core' -or $PSVersionTable.PSVersion.Major -lt 7) {
    $pwshPath = Get-Command pwsh -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1
    if (-not [string]::IsNullOrWhiteSpace($pwshPath) -and (Test-Path -LiteralPath $pwshPath)) {
        Write-Host -ForegroundColor $Colors.WarningMessage "Launching SharePoint connection in PowerShell 7 to avoid the legacy Windows PowerShell PnP assembly mismatch."
        & $pwshPath @scriptInvocationArguments
        exit $LASTEXITCODE
    }

    throw "This script requires PowerShell 7+. Please run it with 'pwsh'."
}

$pwshPath = (Get-Command pwsh -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1)
if ([string]::IsNullOrWhiteSpace($pwshPath) -or -not (Test-Path -LiteralPath $pwshPath)) {
    throw "PowerShell 7 executable 'pwsh' not found."
}

if ($env:O365_PNP_ISOLATED -ne '1') {
    $loadedPnpModules = @(Get-Module -Name PnP.PowerShell -All -ErrorAction SilentlyContinue)
    if ($loadedPnpModules.Count -gt 0) {
        Write-Host -ForegroundColor $Colors.WarningMessage "PnP.PowerShell is already loaded in this session. Launching an isolated PowerShell 7 process to avoid the assembly collision."
        $env:O365_PNP_ISOLATED = '1'
        & $pwshPath @scriptInvocationArguments
        exit $LASTEXITCODE
    }
}

function Import-PnPModuleWithCompat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Colors
    )

    # If the module is already loaded in this session, reuse it. Calling Import-Module
    # -Force on an already-resident PnP.PowerShell causes a .NET assembly collision
    # because the CLR AppDomain cannot unload an assembly once it is loaded.
    $alreadyLoaded = Get-Module -Name PnP.PowerShell | Select-Object -First 1
    if ($null -ne $alreadyLoaded) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "PnP.PowerShell $($alreadyLoaded.Version) already loaded in this session."
        return $alreadyLoaded
    }

    $candidateModules = @(Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending)
    if ($candidateModules.Count -eq 0) {
        throw "PnP.PowerShell module is required."
    }

    $preferredModule = $candidateModules | Where-Object { $_.Path -notlike '*WindowsPowerShell\Modules*' } | Select-Object -First 1

    if ($null -eq $preferredModule) {
        Write-Host -ForegroundColor $Colors.WarningMessage "PnP.PowerShell is only installed under the Windows PowerShell module path. Installing a PowerShell 7 compatible copy to CurrentUser."
        Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop | Out-Null

        $candidateModules = @(Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending)
        $preferredModule = $candidateModules | Where-Object { $_.Path -notlike '*WindowsPowerShell\Modules*' } | Select-Object -First 1
        if ($null -eq $preferredModule) {
            throw "Unable to locate a PowerShell 7 compatible PnP.PowerShell installation."
        }
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Loading PnP.PowerShell from $($preferredModule.Path)"

    # Import without -Force. Using -Force causes PowerShell to attempt to reload
    # the DLL even when it is already resident in the AppDomain, which throws
    # "assembly with same name is already loaded". Without -Force, the module is
    # loaded once and subsequent calls simply reuse the already-resident assembly.
    Import-Module -Name $preferredModule.Path -ErrorAction Stop | Out-Null

    $loadedModule = Get-Module PnP.PowerShell | Select-Object -First 1
    if ($null -eq $loadedModule) {
        throw "PnP.PowerShell was not loaded successfully."
    }

    return $loadedModule
}

function Resolve-SpoCertificateProfile {
    <#
    .SYNOPSIS
        Load and filter the JSON certificate profile map, returning the matching profile entry.
    .DESCRIPTION
        Reads the map file at -Path, applies optional -TenantFilter, -ProfileFilter, and
        -SiteUrlFilter, and returns the single matching profile object. Prompts interactively
        when multiple profiles match and -NoPrompt is not specified. Returns $null when the map
        file is absent.
    .OUTPUTS
        PSCustomObject  The selected profile entry, or $null if the map file is absent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Path,
        [Parameter(Mandatory = $false)]
        [string]$TenantFilter,
        [Parameter(Mandatory = $false)]
        [string]$ProfileFilter,
        [Parameter(Mandatory = $false)]
        [string]$SiteUrlFilter,
        [Parameter(Mandatory = $false)]
        [switch]$NoPrompt,
        [Parameter(Mandatory = $true)]
        [hashtable]$Colors
    )

    Write-Verbose "Resolving certificate profile from map path: $Path"
    Write-Debug "Resolving certificate profile from map path: $Path"

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -Path $Path)) {
        Write-Debug "Certificate map file missing or not provided."
        return $null
    }

    try {
        $raw = Get-Content -Path $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Unable to parse certificate mapping file '$Path'. $($_.Exception.Message)"
    }

    $profileItems = @()
    if ($raw -is [System.Array]) {
        $profileItems = @($raw)
    }
    elseif ($null -ne $raw.profiles) {
        $profileItems = @($raw.profiles)
    }

    if ($profileItems.Count -eq 0) {
        throw "No profiles found in certificate mapping file '$Path'."
    }

    Write-Verbose "Certificate map loaded: $($profileItems.Count) profile(s); applying filters."
    $candidateProfiles = $profileItems
    if (-not [string]::IsNullOrWhiteSpace($ProfileFilter)) {
        $candidateProfiles = @($candidateProfiles | Where-Object { $_.name -eq $ProfileFilter })
    }
    if (-not [string]::IsNullOrWhiteSpace($TenantFilter)) {
        $candidateProfiles = @($candidateProfiles | Where-Object { $_.tenant -eq $TenantFilter -or $_.organization -eq $TenantFilter })
    }
    if (-not [string]::IsNullOrWhiteSpace($SiteUrlFilter)) {
        $candidateProfiles = @($candidateProfiles | Where-Object { $_.siteUrl -eq $SiteUrlFilter })
    }

    if ($candidateProfiles.Count -eq 0) {
        $appliedFilters = @()
        if (-not [string]::IsNullOrWhiteSpace($ProfileFilter)) { $appliedFilters += "ProfileName='$ProfileFilter'" }
        if (-not [string]::IsNullOrWhiteSpace($TenantFilter))  { $appliedFilters += "Tenant='$TenantFilter'" }
        if (-not [string]::IsNullOrWhiteSpace($SiteUrlFilter)) { $appliedFilters += "SiteUrl='$SiteUrlFilter'" }
        $filterDesc = if ($appliedFilters.Count -gt 0) { " (filters: $($appliedFilters -join ', '))" } else { " (no filters applied)" }

        $availableDesc = ($profileItems | ForEach-Object {
            "name='$($_.name)' tenant='$($_.tenant)' siteUrl='$($_.siteUrl)'"
        }) -join '; '

        throw "No matching certificate profile found in '$Path'$filterDesc. Available profiles: [$availableDesc]"
    }

    if ($candidateProfiles.Count -eq 1 -or $NoPrompt) {
        if ($candidateProfiles.Count -gt 1 -and $NoPrompt) {
            throw "Multiple matching profiles found in '$Path'. Specify -ProfileName, -Tenant, or -SiteUrl."
        }
        return $candidateProfiles[0]
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Multiple matching certificate profiles found:"
    for ($index = 0; $index -lt $candidateProfiles.Count; $index++) {
        $displayName = if ([string]::IsNullOrWhiteSpace($candidateProfiles[$index].name)) { "(unnamed)" } else { $candidateProfiles[$index].name }
        Write-Host -ForegroundColor $Colors.ProcessMessage ("[{0}] {1} | Tenant={2} | SiteUrl={3} | AppId={4}" -f ($index + 1), $displayName, $candidateProfiles[$index].tenant, $candidateProfiles[$index].siteUrl, $candidateProfiles[$index].appId)
    }

    do {
        $choice = Read-Host -Prompt "Select profile number"
        [int]$parsedChoice = 0
        $validSelection = [int]::TryParse($choice, [ref]$parsedChoice) -and $parsedChoice -ge 1 -and $parsedChoice -le $candidateProfiles.Count
    } until ($validSelection)

    return $candidateProfiles[$parsedChoice - 1]
}

function New-SpoLocalCertificate {
    <#
    .SYNOPSIS
        Generate a self-signed RSA-2048 certificate for SharePoint Online app authentication.
    .DESCRIPTION
        Creates the certificate in Cert:\CurrentUser\My, exports the public key as a .cer file,
        and optionally exports a password-protected .pfx. -YearsValid must be between 1 and 100.
        Returns an object with Thumbprint, Subject, NotAfter, CerPath, and PfxPath.
    .OUTPUTS
        PSCustomObject  Certificate metadata and paths to the exported .cer and optional .pfx.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SubjectName,
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 100)]
        [int]$YearsValid,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputPath,
        [Parameter(Mandatory = $false)]
        [switch]$ExportPfx,
        [Parameter(Mandatory = $false)]
        [securestring]$PfxPassword,
        [Parameter(Mandatory = $false)]
        [switch]$NoPrompt,
        [Parameter(Mandatory = $false)]
        [string]$FriendlyName = ""
    )

    Write-Debug "Starting local certificate generation."
    Write-Verbose "Creating self-signed certificate: Subject='$SubjectName', YearsValid=$YearsValid, OutputPath='$OutputPath'"

    if (-not (Test-Path -Path $OutputPath)) {
        Write-Debug "Creating certificate output directory: $OutputPath"
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    $certificate = New-SelfSignedCertificate -Subject "CN=$SubjectName" -CertStoreLocation "Cert:\CurrentUser\My" -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears($YearsValid) -KeySpec Signature -ErrorAction Stop

    $resolvedFriendlyName = if ([string]::IsNullOrWhiteSpace($FriendlyName)) { $SubjectName } else { $FriendlyName }
    $certificate.FriendlyName = $resolvedFriendlyName
    Write-Debug "Certificate friendly name set to: $resolvedFriendlyName"

    $safeSubject = ($SubjectName -replace '[^A-Za-z0-9\-_.]', '-')
    $fileBase    = "{0}-{1}" -f $safeSubject, $certificate.Thumbprint
    $cerPath     = Join-Path -Path $OutputPath -ChildPath "$fileBase.cer"

    Export-Certificate -Cert $certificate -FilePath $cerPath -Type CERT -Force -ErrorAction Stop | Out-Null

    $pfxPath = $null
    if ($ExportPfx) {
        $securePfxPassword = $PfxPassword

        if ($null -eq $securePfxPassword -and -not $NoPrompt) {
            $securePfxPassword = Read-Host -Prompt "Enter password for generated PFX file" -AsSecureString
        }

        if ($null -eq $securePfxPassword) {
            throw "ExportGeneratedPfx requires GeneratedPfxPassword (or prompt input when noprompt is not used)."
        }

        $pfxPath = Join-Path -Path $OutputPath -ChildPath "$fileBase.pfx"
        Export-PfxCertificate -Cert $certificate -FilePath $pfxPath -Password $securePfxPassword -Force -ErrorAction Stop | Out-Null
    }

    return [PSCustomObject]@{
        Thumbprint = $certificate.Thumbprint
        Subject    = $certificate.Subject
        NotAfter   = $certificate.NotAfter
        CerPath    = $cerPath
        PfxPath    = $pfxPath
    }
}

function Get-DeviceCodeGraphToken {
    <#
    .SYNOPSIS
        Authenticate to Microsoft Graph using the device-code OAuth2 flow.
    .DESCRIPTION
        Initiates the device-code OAuth2 flow against the given tenant, opens the verification
        URL in the default browser, optionally copies the one-time code to the clipboard, and
        polls the token endpoint until the user completes sign-in or the device code expires.
        Requires delegated permissions; intended for interactive provisioning sessions.
    .OUTPUTS
        String  Raw OAuth2 access token string suitable for use in Authorization headers.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TenantId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ClientId,
        [Parameter(Mandatory = $false)][string]$Scope = "https://graph.microsoft.com/Application.ReadWrite.All https://graph.microsoft.com/AppRoleAssignment.ReadWrite.All",
        [Parameter(Mandatory = $false)][switch]$CopyCodeToClipboard,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    Write-Debug "Requesting device code for tenant: $TenantId, client: $ClientId"

    try {
        $deviceCodeResponse = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
            -Body @{ client_id = $ClientId; scope = $Scope } `
            -ContentType "application/x-www-form-urlencoded" `
            -ErrorAction Stop
    }
    catch {
        throw "Device code request failed. Check -SetupClientId and -Tenant. Error: $($_.Exception.Message)"
    }

    if ($CopyCodeToClipboard) {
        ## Clipboard copy is opt-in only to reduce exposure on shared/RDP sessions.
        Set-Clipboard -Value $deviceCodeResponse.user_code
    }

    Write-Host -ForegroundColor $Colors.SystemMessage "`n--- Graph Authentication Required ---"
    Write-Host -ForegroundColor $Colors.SystemMessage "Opening browser: $($deviceCodeResponse.verification_uri)"
    if ($CopyCodeToClipboard) {
        Write-Host -ForegroundColor $Colors.SystemMessage "Device code (copied to clipboard): $($deviceCodeResponse.user_code)"
    }
    else {
        Write-Host -ForegroundColor $Colors.SystemMessage "Device code: $($deviceCodeResponse.user_code)"
        Write-Host -ForegroundColor $Colors.WarningMessage "Clipboard copy is disabled by default for security on shared/RDP sessions. Use -CopyDeviceCodeToClipboard to enable it."
    }
    Write-Host -ForegroundColor $Colors.SystemMessage "Paste the code in the browser and sign in, then return here."
    Write-Host -ForegroundColor $Colors.SystemMessage "-------------------------------------`n"

    ## Open the verification URL in the default browser
    Start-Process $deviceCodeResponse.verification_uri

    $tokenBody = @{
        grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
        client_id   = $ClientId
        device_code = $deviceCodeResponse.device_code
    }

    $deadline     = (Get-Date).AddSeconds($deviceCodeResponse.expires_in)
    $pollInterval = [int]$deviceCodeResponse.interval

    ## NOTE: 'continue' inside a switch/catch does not continue the while loop in PowerShell.
    ## Use an explicit $keepPolling flag instead.
    :pollLoop while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $pollInterval
        try {
            $tokenResponse = Invoke-RestMethod -Method Post `
                -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
                -Body $tokenBody `
                -ContentType "application/x-www-form-urlencoded" `
                -ErrorAction Stop
            Write-Debug "Graph token acquired."
            return $tokenResponse.access_token
        }
        catch {
            $errorContent = $null
            try { $errorContent = ($_.ErrorDetails.Message | ConvertFrom-Json) } catch {}
            if ($null -ne $errorContent) {
                switch ($errorContent.error) {
                    "authorization_pending" {
                        Write-Debug "Waiting for user to complete sign-in..."
                        continue pollLoop
                    }
                    "slow_down" {
                        $pollInterval += 5
                        continue pollLoop
                    }
                    "authorization_declined" { throw "User declined the authorization request." }
                    "expired_token"          { throw "Device code expired before authorization was completed." }
                    default                  { throw "Token exchange failed ($($errorContent.error)): $($errorContent.error_description)" }
                }
            }
            throw "Token exchange failed: $($_.Exception.Message)"
        }
    }
    throw "Device code authorization timed out."
}

function Invoke-SpoGraphRequest {
    <#
    .SYNOPSIS
        Helper: make an authenticated Graph REST call and return the parsed response.
    .DESCRIPTION
        Wraps Invoke-RestMethod with an exponential-backoff retry loop. Respects Retry-After
        response headers for HTTP 429 and 5xx status codes and matches common throttle/transient
        error message patterns. Per-retry delay is capped at 30 seconds. Throws on non-retriable
        errors or once MaxRetries is exhausted.
    .OUTPUTS
        PSObject  The parsed JSON response body returned by Microsoft Graph.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateSet('Get', 'Post', 'Patch', 'Put', 'Delete')][string]$Method,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Uri,
        [Parameter(Mandatory = $false)][object]$Body,
        [Parameter(Mandatory = $false)][ValidateRange(0, 10)][int]$MaxRetries = 5,
        [Parameter(Mandatory = $false)][ValidateRange(1, 60)][int]$InitialRetryDelaySeconds = 2
    )

    $headers = @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" }
    $params  = @{ Method = $Method; Uri = $Uri; Headers = $headers; ErrorAction = "Stop" }
    if ($null -ne $Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
    }

    $attempt = 0
    while ($true) {
        try {
            return Invoke-RestMethod @params
        }
        catch {
            $attempt++

            $detail = $null
            try { $detail = ($_.ErrorDetails.Message | ConvertFrom-Json).error.message } catch {}
            $msg = if ($null -ne $detail) { $detail } else { $_.Exception.Message }

            $statusCode = $null
            $retryAfterSeconds = $null
            $response = $null
            try { $response = $_.Exception.Response } catch {}
            if ($null -ne $response) {
                try { $statusCode = [int]$response.StatusCode } catch {}
                try {
                    $retryAfterValue = $response.Headers['Retry-After']
                    if ($null -ne $retryAfterValue) {
                        ## PS5.1 returns a string; PS7 returns IEnumerable<string>.
                        ## Normalise to string before parsing to avoid [0] giving the first character on PS5.
                        $retryAfterStr = if ($retryAfterValue -is [string]) { $retryAfterValue } else { @($retryAfterValue) | Select-Object -First 1 }
                        [int]$parsedRetryAfter = 0
                        if (-not [string]::IsNullOrWhiteSpace($retryAfterStr) -and [int]::TryParse($retryAfterStr, [ref]$parsedRetryAfter)) {
                            $retryAfterSeconds = $parsedRetryAfter
                        }
                    }
                }
                catch {}
            }

            $isRetriableStatus = ($statusCode -in @(429, 500, 502, 503, 504))
            $isRetriableMessage = ($msg -match '(?i)too many requests|throttl|rate limit|temporar|timeout|try again')
            $shouldRetry = ($attempt -le $MaxRetries) -and ($isRetriableStatus -or $isRetriableMessage)

            if (-not $shouldRetry) {
                throw "Graph call failed [$Method $Uri]: $msg"
            }

            $backoffSeconds = [math]::Pow(2, ($attempt - 1)) * $InitialRetryDelaySeconds
            $delaySeconds = if ($null -ne $retryAfterSeconds -and $retryAfterSeconds -gt 0) { [int]$retryAfterSeconds } else { [int][math]::Min(30, $backoffSeconds) }
            Write-Debug "Graph call retry $attempt/$MaxRetries after ${delaySeconds}s for [$Method $Uri]. Status=$statusCode"
            Start-Sleep -Seconds $delaySeconds
        }
    }
}

function Resolve-SpoTenantRootSiteUrl {
    <#
    .SYNOPSIS
        Resolve the tenant root SharePoint Online URL from tenant input and Graph metadata.
    .DESCRIPTION
        Uses these strategies in order:
        1) Explicit -SiteUrl value (passed in via ExistingSiteUrl)
        2) Direct derivation from an *.onmicrosoft.com tenant
        3) Microsoft Graph organization verifiedDomains lookup (isInitial domain)
        4) Best-effort custom domain prefix fallback (for example, contoso.com -> contoso.sharepoint.com)

        If none succeed, prompts (unless -NoPrompt) or throws.
    .OUTPUTS
        String  Tenant root SharePoint URL (for example, https://contoso.sharepoint.com)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ExistingSiteUrl,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantHint,
        [Parameter(Mandatory = $false)]
        [string]$AccessToken,
        [Parameter(Mandatory = $false)]
        [switch]$NoPrompt,
        [Parameter(Mandatory = $true)]
        [hashtable]$Colors
    )

    if (-not [string]::IsNullOrWhiteSpace($ExistingSiteUrl)) {
        Write-Debug "Using caller-provided SiteUrl: $ExistingSiteUrl"
        return $ExistingSiteUrl
    }

    if ($TenantHint -match '^([^.]+)\.onmicrosoft\.com$') {
        $derivedFromInitial = "https://$($Matches[1]).sharepoint.com"
        Write-Debug "Derived site URL directly from onmicrosoft tenant: $derivedFromInitial"
        return $derivedFromInitial
    }

    $fallbackFromCustomDomain = $null
    if ($TenantHint -match '^([^.]+)\.[^.]+(?:\.[^.]+)*$') {
        $fallbackFromCustomDomain = "https://$($Matches[1]).sharepoint.com"
        Write-Debug "Prepared custom-domain fallback candidate: $fallbackFromCustomDomain"
    }

    if (-not [string]::IsNullOrWhiteSpace($AccessToken)) {
        try {
            Write-Host -ForegroundColor $Colors.ProcessMessage "Resolving tenant SharePoint host from Graph verified domains..."
            $orgResponse = Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Get -Uri "https://graph.microsoft.com/v1.0/organization?`$select=verifiedDomains" -ErrorAction Stop

            $org = $null
            if ($null -ne $orgResponse.value -and $orgResponse.value.Count -gt 0) {
                $org = $orgResponse.value[0]
            }

            $initialDomain = $null
            if ($null -ne $org -and $null -ne $org.verifiedDomains) {
                $initialDomain = @($org.verifiedDomains | Where-Object { $_.isInitial -eq $true -and $_.name -like '*.onmicrosoft.com' } | Select-Object -First 1).name
            }

            if (-not [string]::IsNullOrWhiteSpace($initialDomain) -and $initialDomain -match '^([^.]+)\.onmicrosoft\.com$') {
                $resolvedFromGraph = "https://$($Matches[1]).sharepoint.com"
                Write-Debug "Resolved site URL from Graph initial domain '$initialDomain': $resolvedFromGraph"
                return $resolvedFromGraph
            }

            Write-Debug "Graph organization query succeeded but no initial onmicrosoft domain was returned."
        }
        catch {
            Write-Debug "Graph-based tenant root resolution failed: $($_.Exception.Message)"
            Write-Host -ForegroundColor $Colors.WarningMessage "Could not resolve initial tenant domain from Graph; using fallback logic."
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($fallbackFromCustomDomain)) {
        Write-Host -ForegroundColor $Colors.WarningMessage "Using custom-domain fallback for SharePoint URL: $fallbackFromCustomDomain"
        return $fallbackFromCustomDomain
    }

    if (-not $NoPrompt) {
        $promptValue = $null
        do {
            $promptValue = Read-Host -Prompt "Enter the tenant root SharePoint URL (e.g. https://contoso.sharepoint.com)"
        } until (-not [string]::IsNullOrWhiteSpace($promptValue))
        return $promptValue
    }

    throw "-ProvisionEntraApp could not derive SiteUrl from -Tenant '$TenantHint'. Supply -SiteUrl explicitly (e.g. https://contoso.sharepoint.com)."
}

function Set-SpoProfileMapEntry {
    <#
    .SYNOPSIS
        Atomically upsert a certificate profile entry in the JSON profile map file.
    .DESCRIPTION
        Acquires a named system mutex scoped to the resolved map file path, reads the existing
        map, replaces the entry matching the same tenant or appId (or appends a new entry), then
        writes the result via a temp-file atomic replace. Handles abandoned mutex from prior
        crashes. Concurrent callers are serialised rather than silently racing.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$MapPath,
        [Parameter(Mandatory = $true)][object]$ProfileEntry,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $fullMapPath = [System.IO.Path]::GetFullPath($MapPath)
    $hashBytes = [System.Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::UTF8.GetBytes($fullMapPath.ToLowerInvariant()))
    $hashHex = ([System.BitConverter]::ToString($hashBytes)).Replace('-', '')
    $mutexName = "Global\CIAOPS_SPO_PROFILEMAP_$hashHex"

    $mutex = New-Object System.Threading.Mutex($false, $mutexName)
    $hasHandle = $false
    $tempPath = "$fullMapPath.$PID.tmp"

    try {
        try {
            $hasHandle = $mutex.WaitOne([TimeSpan]::FromSeconds(30))
        }
        catch [System.Threading.AbandonedMutexException] {
            $hasHandle = $true
            Write-Debug "Profile map mutex was abandoned by a previous run; continuing with recovered lock ownership."
        }
        if (-not $hasHandle) {
            throw "Timed out waiting for profile map lock: $fullMapPath"
        }
        Write-Verbose "Profile map lock acquired for: $fullMapPath"
        Write-Verbose "Profile map lock acquired for: $fullMapPath"

        $mapData = @{ profiles = @() }
        if (Test-Path -Path $fullMapPath) {
            try {
                $raw = Get-Content -Path $fullMapPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                if ($null -ne $raw.profiles) { $mapData = $raw }
            }
            catch {
                Write-Debug "Could not parse existing profile map inside lock - will overwrite."
            }
        }

        $profileList = [System.Collections.Generic.List[object]]::new()
        foreach ($p in $mapData.profiles) { $profileList.Add($p) }

        ## Replace existing entry for the same tenant/app, or append.
        ## This preserves profiles that share the same display name across different tenants.
        $existingIdx = $null
        for ($i = 0; $i -lt $profileList.Count; $i++) {
            $sameApp = (-not [string]::IsNullOrWhiteSpace($profileList[$i].appId) -and $profileList[$i].appId -eq $ProfileEntry.appId)
            $sameTenant = ((-not [string]::IsNullOrWhiteSpace($profileList[$i].tenant) -and $profileList[$i].tenant -eq $ProfileEntry.tenant) -or
                (-not [string]::IsNullOrWhiteSpace($profileList[$i].organization) -and $profileList[$i].organization -eq $ProfileEntry.organization))

            if ($sameApp -or $sameTenant) { $existingIdx = $i; break }
        }
        if ($null -ne $existingIdx) { $profileList[$existingIdx] = $ProfileEntry } else { $profileList.Add($ProfileEntry) }

        @{ profiles = $profileList.ToArray() } | ConvertTo-Json -Depth 5 | Set-Content -Path $tempPath -Encoding UTF8
        Move-Item -Path $tempPath -Destination $fullMapPath -Force
        Write-Host -ForegroundColor $Colors.ProcessMessage "Profile map updated: $fullMapPath"
    }
    finally {
        if (Test-Path -Path $tempPath) {
            Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
        }
        if ($hasHandle) {
            $mutex.ReleaseMutex()
        }
        $mutex.Dispose()
    }
}

function Get-CertClientAssertionToken {
    <#
    .SYNOPSIS
        Acquire a Graph access token using a JWT client assertion signed with a local certificate.
        No user interaction required - uses the OAuth2 client_credentials flow.
    .DESCRIPTION
        Builds a signed RS256 JWT (header, payload, PKCS#1 v1.5 signature) from the provided
        certificate's RSA private key, then exchanges it at the tenant token endpoint for an
        OAuth2 access token. The assertion is valid for 10 minutes; the returned token has the
        lifetime granted by Entra ID (typically one hour).
    .OUTPUTS
        String  Raw OAuth2 access token string.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TenantId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AppId,
        [Parameter(Mandatory = $true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [string]$Scope = "https://graph.microsoft.com/.default"
    )

    ## Build the x5t (base64url of cert SHA-1 thumbprint bytes) for the JWT header.
    $thumbprintBytes = $Certificate.GetCertHash()
    $x5t = [System.Convert]::ToBase64String($thumbprintBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    ## [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() is portable across PS5.1 and PS7.
    ## Get-Date -UFormat %s is only available on PS7+ and fails silently on PS5.1.
    $now           = [int][DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

    ## JWT header and payload - both must be base64url encoded.
    $headerJson  = '{"alg":"RS256","typ":"JWT","x5t":"' + $x5t + '"}'
    $payloadJson = '{"aud":"' + $tokenEndpoint + '","iss":"' + $AppId + '","sub":"' + $AppId + '","jti":"' + [System.Guid]::NewGuid().ToString() + '","nbf":' + $now + ',"exp":' + ($now + 600) + '}'

    $headerB64   = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($headerJson)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $payloadB64  = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payloadJson)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $signingInput = [System.Text.Encoding]::UTF8.GetBytes("$headerB64.$payloadB64")

    ## Sign with the certificate's RSA private key using PKCS#1 v1.5 / SHA-256.
    $rsa       = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    $sigBytes  = $rsa.SignData($signingInput, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $sigB64    = [System.Convert]::ToBase64String($sigBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    $clientAssertion = "$headerB64.$payloadB64.$sigB64"

    Write-Verbose "Requesting Graph token via client assertion for app $AppId in tenant $TenantId"
    Write-Debug "Requesting Graph token via client assertion for app $AppId in tenant $TenantId"

    try {
        $response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ErrorAction Stop `
            -ContentType "application/x-www-form-urlencoded" `
            -Body @{
                grant_type             = "client_credentials"
                client_id              = $AppId
                scope                  = $Scope
                client_assertion_type  = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
                client_assertion       = $clientAssertion
            }
        return $response.access_token
    }
    catch {
        $detail = $null
        try { $detail = ($_.ErrorDetails.Message | ConvertFrom-Json).error_description } catch {}
        throw "Client assertion token request failed: $(if ($detail) { $detail } else { $_.Exception.Message })"
    }
}

function Write-SpoCertConnectionDetails {
    <#
    .SYNOPSIS
        Display local certificate details and the matching Entra ID app/keyCredential after connecting.
    .DESCRIPTION
        Prints the local certificate's friendly name, subject, thumbprint, issuer, and validity
        window. Then acquires a client-assertion Graph token to read the app registration and
        match the certificate against stored keyCredentials. Falls back gracefully with a
        descriptive warning when the app lacks Application.Read.All permission.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TenantId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AppId,
        [Parameter(Mandatory = $true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$LocalCert,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $sep = "-" * 60

    Write-Host -ForegroundColor $Colors.ProcessMessage "`n$sep"
    Write-Host -ForegroundColor $Colors.ProcessMessage "  LOCAL CERTIFICATE"
    Write-Host -ForegroundColor $Colors.ProcessMessage $sep
    Write-Host -ForegroundColor $Colors.ProcessMessage ("  Friendly Name : {0}" -f $(if ($LocalCert.FriendlyName) { $LocalCert.FriendlyName } else { "(none)" }))
    Write-Host -ForegroundColor $Colors.ProcessMessage ("  Subject       : {0}" -f $LocalCert.Subject)
    Write-Host -ForegroundColor $Colors.ProcessMessage ("  Thumbprint    : {0}" -f $LocalCert.Thumbprint)
    Write-Host -ForegroundColor $Colors.ProcessMessage ("  Issuer        : {0}" -f $LocalCert.Issuer)
    Write-Host -ForegroundColor $Colors.ProcessMessage ("  Valid From    : {0}" -f $LocalCert.NotBefore.ToString('yyyy-MM-dd HH:mm:ss'))
    Write-Host -ForegroundColor $Colors.ProcessMessage ("  Valid To      : {0}" -f $LocalCert.NotAfter.ToString('yyyy-MM-dd HH:mm:ss'))

    ## Attempt to fetch Entra app details using a client assertion (no user interaction).
    try {
        $graphToken = Get-CertClientAssertionToken -TenantId $TenantId -AppId $AppId -Certificate $LocalCert
        $graphBase  = "https://graph.microsoft.com/v1.0"
        $appFilter  = [uri]::EscapeDataString("appId eq '" + ($AppId -replace "'", "''") + "'")
        $appResult  = Invoke-SpoGraphRequest -AccessToken $graphToken -Method Get `
                          -Uri "$graphBase/applications?`$filter=$appFilter&`$select=displayName,keyCredentials"
        $appObj     = $appResult.value | Select-Object -First 1

        Write-Host -ForegroundColor $Colors.ProcessMessage "`n$sep"
        Write-Host -ForegroundColor $Colors.ProcessMessage "  ENTRA ID APP REGISTRATION"
        Write-Host -ForegroundColor $Colors.ProcessMessage $sep
        Write-Host -ForegroundColor $Colors.ProcessMessage ("  Display Name  : {0}" -f $appObj.displayName)
        Write-Host -ForegroundColor $Colors.ProcessMessage ("  App ID        : {0}" -f $AppId)

        ## Find the keyCredential whose customKeyIdentifier matches this cert's SHA-1 thumbprint.
        $thumbBase64 = [System.Convert]::ToBase64String($LocalCert.GetCertHash())
        $matchingKey = $appObj.keyCredentials | Where-Object { $_.customKeyIdentifier -eq $thumbBase64 } | Select-Object -First 1

        if ($null -ne $matchingKey) {
            Write-Host -ForegroundColor $Colors.ProcessMessage ("  Cert Label    : {0}" -f $matchingKey.displayName)
            Write-Host -ForegroundColor $Colors.ProcessMessage ("  Cert Start    : {0}" -f ([datetime]$matchingKey.startDateTime).ToString('yyyy-MM-dd HH:mm:ss'))
            Write-Host -ForegroundColor $Colors.ProcessMessage ("  Cert End      : {0}" -f ([datetime]$matchingKey.endDateTime).ToString('yyyy-MM-dd HH:mm:ss'))
        }
        else {
            Write-Host -ForegroundColor $Colors.WarningMessage "  Matching key credential not found in Entra app (thumbprint mismatch or cert not yet uploaded)."
        }

        ## Show all other certs registered on the same app.
        $otherKeys = @($appObj.keyCredentials | Where-Object { $_.customKeyIdentifier -ne $thumbBase64 })
        if ($otherKeys.Count -gt 0) {
            Write-Host -ForegroundColor $Colors.ProcessMessage ("`n  Other registered certs on this app ({0}):" -f $otherKeys.Count)
            foreach ($key in $otherKeys) {
                Write-Host -ForegroundColor $Colors.ProcessMessage ("    - {0}  [{1} -> {2}]" -f $key.displayName, ([datetime]$key.startDateTime).ToString('yyyy-MM-dd'), ([datetime]$key.endDateTime).ToString('yyyy-MM-dd'))
            }
        }
    }
    catch {
        $entraDetailError = $_.Exception.Message
        if ($entraDetailError -match "Insufficient privileges") {
            Write-Host -ForegroundColor $Colors.WarningMessage "`n  Entra ID cert details skipped: this app lacks Graph read permission for applications."
            Write-Host -ForegroundColor $Colors.WarningMessage "  Local cert details above are valid; SPO connection is unaffected."
            Write-Host -ForegroundColor $Colors.WarningMessage "  To enable Entra matching details, grant Microsoft Graph Application.Read.All (Application) and admin consent."
        }
        else {
            Write-Host -ForegroundColor $Colors.WarningMessage "`n  (Entra ID cert details unavailable: $entraDetailError)"
        }
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "$sep`n"
}

function Get-SpoProvisioningRoleTargets {
    <#
    .SYNOPSIS
        Resolve the SharePoint Online and Microsoft Graph service principals and required app role IDs.
    .DESCRIPTION
        Queries Microsoft Graph to locate the SharePoint Online service principal (filtered by
        -SpoResourceAppId) and its Sites.FullControl.All role, and the Microsoft Graph service
        principal and its Application.Read.All role. Returns all four objects for use by
        downstream provisioning helpers.
    .OUTPUTS
        PSCustomObject  SpoServicePrincipal, SpoSitesFullRole, GraphServicePrincipal, GraphReadAllRole.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$SpoResourceAppId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$GraphResourceAppId,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $graphBase = "https://graph.microsoft.com/v1.0"

    Write-Host -ForegroundColor $Colors.ProcessMessage "Locating SharePoint Online service principal in directory..."
    $spoSpFilter = [uri]::EscapeDataString("appId eq '$SpoResourceAppId'")
    $spoSpResult = Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals?`$filter=$spoSpFilter"
    if ($null -eq $spoSpResult.value -or $spoSpResult.value.Count -eq 0) {
        throw "SharePoint Online service principal not found. Ensure SharePoint Online is provisioned in this tenant."
    }
    $spoSp = $spoSpResult.value[0]
    $sitesFullRole = $spoSp.appRoles | Where-Object { $_.value -eq "Sites.FullControl.All" }
    if ($null -eq $sitesFullRole) {
        throw "Sites.FullControl.All role not found on SharePoint Online service principal."
    }
    Write-Debug "Sites.FullControl.All role ID: $($sitesFullRole.id)"

    Write-Host -ForegroundColor $Colors.ProcessMessage "Locating Microsoft Graph service principal in directory..."
    $graphSpFilter = [uri]::EscapeDataString("appId eq '$GraphResourceAppId'")
    $graphSpResult = Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals?`$filter=$graphSpFilter"
    if ($null -eq $graphSpResult.value -or $graphSpResult.value.Count -eq 0) {
        throw "Microsoft Graph service principal not found in this tenant."
    }
    $graphSp = $graphSpResult.value[0]
    $graphReadAllRole = $graphSp.appRoles | Where-Object { $_.value -eq "Application.Read.All" -and $_.allowedMemberTypes -contains "Application" } | Select-Object -First 1
    if ($null -eq $graphReadAllRole) {
        throw "Application.Read.All app role not found on Microsoft Graph service principal."
    }
    Write-Debug "Graph Application.Read.All role ID: $($graphReadAllRole.id)"

    return [PSCustomObject]@{
        SpoServicePrincipal   = $spoSp
        SpoSitesFullRole      = $sitesFullRole
        GraphServicePrincipal = $graphSp
        GraphReadAllRole      = $graphReadAllRole
    }
}

function Get-OrCreateSpoEntraApplication {
    <#
    .SYNOPSIS
        Return an existing Entra app registration by display name, or create a new one.
    .DESCRIPTION
        Searches Graph for an app registration matching -DisplayName. If found, returns it
        unchanged. If not found, creates a new single-tenant app with requiredResourceAccess
        entries for Sites.FullControl.All and Graph Application.Read.All, then returns the
        new app object.
    .OUTPUTS
        PSObject  The Entra app registration object from Microsoft Graph.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DisplayName,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$SpoSitesFullRoleId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$GraphReadAllRoleId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$SpoResourceAppId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$GraphResourceAppId,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $graphBase = "https://graph.microsoft.com/v1.0"

    Write-Host -ForegroundColor $Colors.ProcessMessage "Checking for existing app registration: $DisplayName..."
    $appFilter = [uri]::EscapeDataString("displayName eq '" + ($DisplayName -replace "'", "''") + "'")
    $existingApps = Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/applications?`$filter=$appFilter"

    if ($existingApps.value.Count -gt 0) {
        Write-Host -ForegroundColor $Colors.WarningMessage "App '$DisplayName' already exists - reusing existing registration."
        return $existingApps.value[0]
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Creating app registration: $DisplayName..."
    $newAppBody = @{
        displayName            = $DisplayName
        signInAudience         = "AzureADMyOrg"
        requiredResourceAccess = @(
            @{
                resourceAppId  = $SpoResourceAppId
                resourceAccess = @(
                    @{ id = $SpoSitesFullRoleId; type = "Role" }
                )
            },
            @{
                resourceAppId  = $GraphResourceAppId
                resourceAccess = @(
                    @{ id = $GraphReadAllRoleId; type = "Role" }
                )
            }
        )
    }
    $appObject = Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/applications" -Body $newAppBody
    Write-Host -ForegroundColor $Colors.ProcessMessage "App created. Object ID: $($appObject.id) | App ID: $($appObject.appId)"
    return $appObject
}

function Set-SpoEntraApplicationCertificate {
    <#
    .SYNOPSIS
        Upload a local certificate to an Entra app registration and verify it was stored.
    .DESCRIPTION
        Exports the DER bytes of the certificate matching -Thumbprint from Cert:\CurrentUser\My,
        patches the app's keyCredentials via Graph while preserving any existing certificates,
        then re-reads the app to confirm the upload succeeded before returning.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AppObjectId,
        [Parameter(Mandatory = $true)][ValidatePattern('^[0-9A-Fa-f]{40}$')][string]$Thumbprint,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $graphBase = "https://graph.microsoft.com/v1.0"

    Write-Host -ForegroundColor $Colors.ProcessMessage "Uploading certificate to app registration..."

    $certStoreObj = Get-Item "Cert:\CurrentUser\My\$Thumbprint" -ErrorAction Stop
    $cerBytes = $certStoreObj.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    $cerBase64 = [System.Convert]::ToBase64String($cerBytes)
    $customKeyIdBase64 = [System.Convert]::ToBase64String($certStoreObj.GetCertHash())

    if ($cerBytes.Length -eq 0) {
        throw "Certificate export produced empty bytes for thumbprint '$Thumbprint'. Cert may not be in Cert:\CurrentUser\My."
    }
    Write-Debug "Cert bytes: $($cerBytes.Length) | customKeyIdentifier: $customKeyIdBase64"

    $existingApp = Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/applications/${AppObjectId}?`$select=keyCredentials"
    $existingKeys = @($existingApp.keyCredentials | Where-Object { $_.customKeyIdentifier -ne $customKeyIdBase64 })
    Write-Debug "Existing key credentials on app (excluding this thumbprint): $($existingKeys.Count)"

    $newKey = @{
        type                = "AsymmetricX509Cert"
        usage               = "Verify"
        key                 = $cerBase64
        displayName         = "$($env:COMPUTERNAME) - SPO-Auth-$Thumbprint"
        customKeyIdentifier = $customKeyIdBase64
    }

    $certPatch = @{ keyCredentials = @($existingKeys) + @($newKey) }
    Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Patch -Uri "$graphBase/applications/$AppObjectId" -Body $certPatch | Out-Null

    $appVerify = Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/applications/${AppObjectId}?`$select=keyCredentials"
    $storedKey = $appVerify.keyCredentials | Where-Object { $_.customKeyIdentifier -eq $customKeyIdBase64 } | Select-Object -First 1
    if ($null -eq $storedKey) {
        $stored = ($appVerify.keyCredentials | ForEach-Object { $_.customKeyIdentifier }) -join ', '
        throw "Certificate upload verification failed. Expected customKeyIdentifier '$customKeyIdBase64' not found in app keyCredentials. Stored: $stored"
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Certificate uploaded and verified in app registration."
}

function Get-OrCreateSpoEntraServicePrincipal {
    <#
    .SYNOPSIS
        Return the service principal for the given app ID, creating it if it does not yet exist.
    .DESCRIPTION
        Queries Graph for a service principal matching -AppId. Returns the existing object when
        found, otherwise posts a new service principal and returns the created object.
    .OUTPUTS
        PSObject  The service principal object from Microsoft Graph.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AppId,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $graphBase = "https://graph.microsoft.com/v1.0"

    Write-Host -ForegroundColor $Colors.ProcessMessage "Ensuring service principal exists..."
    $spFilter = [uri]::EscapeDataString("appId eq '$AppId'")
    $existingSp = Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals?`$filter=$spFilter"

    if ($existingSp.value.Count -gt 0) {
        $spObject = $existingSp.value[0]
        Write-Host -ForegroundColor $Colors.ProcessMessage "Service principal already exists. Object ID: $($spObject.id)"
        return $spObject
    }

    $spObject = Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/servicePrincipals" -Body @{ appId = $AppId }
    Write-Host -ForegroundColor $Colors.ProcessMessage "Service principal created. Object ID: $($spObject.id)"
    return $spObject
}

function Set-SpoEntraAppRoleAssignments {
    <#
    .SYNOPSIS
        Grant Sites.FullControl.All and Graph Application.Read.All to the provisioned service principal.
    .DESCRIPTION
        Checks existing app role assignments for the service principal and grants each missing
        role. SPO is checked first, then Graph, with a fresh assignment fetch before each grant
        to avoid stale data. Already-granted roles are skipped silently.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ServicePrincipalId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$SpoServicePrincipalId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$SpoSitesFullRoleId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$GraphServicePrincipalId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$GraphReadAllRoleId,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $graphBase = "https://graph.microsoft.com/v1.0"

    Write-Host -ForegroundColor $Colors.ProcessMessage "Granting Sites.FullControl.All permission (admin consent)..."
    $spoAssignments = Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals/$ServicePrincipalId/appRoleAssignments"
    $alreadyGranted = $spoAssignments.value | Where-Object { $_.appRoleId -eq $SpoSitesFullRoleId -and $_.resourceId -eq $SpoServicePrincipalId }
    if ($null -ne $alreadyGranted) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Sites.FullControl.All already granted - skipping."
    }
    else {
        $roleBody = @{ principalId = $ServicePrincipalId; resourceId = $SpoServicePrincipalId; appRoleId = $SpoSitesFullRoleId }
        Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/servicePrincipals/$ServicePrincipalId/appRoleAssignments" -Body $roleBody | Out-Null
        Write-Host -ForegroundColor $Colors.ProcessMessage "Sites.FullControl.All granted."
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Granting Microsoft Graph Application.Read.All permission (admin consent)..."
    ## Re-fetch assignments after the SPO grant check/post so this remains correct if grant order changes.
    $graphAssignments = Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals/$ServicePrincipalId/appRoleAssignments"
    $graphAlreadyGranted = $graphAssignments.value | Where-Object { $_.appRoleId -eq $GraphReadAllRoleId -and $_.resourceId -eq $GraphServicePrincipalId }
    if ($null -ne $graphAlreadyGranted) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Microsoft Graph Application.Read.All already granted - skipping."
    }
    else {
        $graphRoleBody = @{ principalId = $ServicePrincipalId; resourceId = $GraphServicePrincipalId; appRoleId = $GraphReadAllRoleId }
        Invoke-SpoGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/servicePrincipals/$ServicePrincipalId/appRoleAssignments" -Body $graphRoleBody | Out-Null
        Write-Host -ForegroundColor $Colors.ProcessMessage "Microsoft Graph Application.Read.All granted."
    }
}

function Invoke-SpoAppProvisioning {
    <#
    .SYNOPSIS
        Orchestrates Entra app provisioning for SPO by invoking focused helper functions.
    .DESCRIPTION
        Calls Get-SpoProvisioningRoleTargets, Get-OrCreateSpoEntraApplication,
        Set-SpoEntraApplicationCertificate, Get-OrCreateSpoEntraServicePrincipal, and
        Set-SpoEntraAppRoleAssignments in sequence. Requires FullLanguage mode. All resource-app
        IDs are passed explicitly; no hidden scope reads.
    .OUTPUTS
        PSCustomObject  AppId (Entra client app ID), AppObjId (app object ID), SpObjId (service principal object ID).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DisplayName,
        [Parameter(Mandatory = $true)][string]$CerFilePath,
        [Parameter(Mandatory = $true)][ValidatePattern('^[0-9A-Fa-f]{40}$')][string]$Thumbprint,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$SpoResourceAppId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$GraphResourceAppId,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    ## Provisioning helpers use static .NET method calls that are blocked in Constrained Language Mode.
    if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
        throw "App provisioning requires FullLanguage mode. Current mode: $($ExecutionContext.SessionState.LanguageMode). Run the script in a standard PowerShell session (not one with Constrained Language Mode active)."
    }

    ## CerFilePath retained for backward compatibility in script interface.
    if (-not [string]::IsNullOrWhiteSpace($CerFilePath) -and -not (Test-Path -Path $CerFilePath)) {
        throw "Certificate file path '$CerFilePath' was provided but does not exist."
    }

    $targets = Get-SpoProvisioningRoleTargets -AccessToken $AccessToken -SpoResourceAppId $SpoResourceAppId -GraphResourceAppId $GraphResourceAppId -Colors $Colors
    $appObject = Get-OrCreateSpoEntraApplication `
        -AccessToken $AccessToken `
        -DisplayName $DisplayName `
        -SpoSitesFullRoleId $targets.SpoSitesFullRole.id `
        -GraphReadAllRoleId $targets.GraphReadAllRole.id `
        -SpoResourceAppId $SpoResourceAppId `
        -GraphResourceAppId $GraphResourceAppId `
        -Colors $Colors

    Set-SpoEntraApplicationCertificate -AccessToken $AccessToken -AppObjectId $appObject.id -Thumbprint $Thumbprint -Colors $Colors

    $spObject = Get-OrCreateSpoEntraServicePrincipal -AccessToken $AccessToken -AppId $appObject.appId -Colors $Colors

    Set-SpoEntraAppRoleAssignments `
        -AccessToken $AccessToken `
        -ServicePrincipalId $spObject.id `
        -SpoServicePrincipalId $targets.SpoServicePrincipal.id `
        -SpoSitesFullRoleId $targets.SpoSitesFullRole.id `
        -GraphServicePrincipalId $targets.GraphServicePrincipal.id `
        -GraphReadAllRoleId $targets.GraphReadAllRole.id `
        -Colors $Colors

    return [PSCustomObject]@{
        AppId    = $appObject.appId
        AppObjId = $appObject.id
        SpObjId  = $spObject.id
    }
}

function Write-SpoConnectedSite {
    <#
    .SYNOPSIS
        Prints the active SharePoint Online connection URL and web title.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$RequestedSiteUrl,
        [Parameter(Mandatory = $true)]
        [hashtable]$Colors
    )

    $connectedUrl   = $null
    $connectedTitle = $null

    try {
        $connection = Get-PnPConnection -ErrorAction Stop
        if ($null -ne $connection) {
            $connectedUrl = $connection.Url
        }
    }
    catch {
        Write-Debug "Get-PnPConnection unavailable or no active connection."
    }

    if (-not [string]::IsNullOrWhiteSpace($connectedUrl)) {
        try {
            $web            = Get-PnPWeb -ErrorAction Stop
            $connectedTitle = $web.Title
        }
        catch {
            Write-Debug "Could not retrieve web title from current connection."
        }
    }

    if ([string]::IsNullOrWhiteSpace($connectedUrl)) {
        $connectedUrl = "(unable to determine from current session)"
    }

    $titleDisplay = if (-not [string]::IsNullOrWhiteSpace($connectedTitle)) { " ($connectedTitle)" } else { "" }
    Write-Host -ForegroundColor $Colors.SystemMessage "Connected SharePoint site: ${connectedUrl}${titleDisplay}"

    if (-not [string]::IsNullOrWhiteSpace($RequestedSiteUrl) -and
        $connectedUrl -ne "(unable to determine from current session)" -and
        $RequestedSiteUrl.TrimEnd('/') -ne $connectedUrl.TrimEnd('/')) {
        Write-Host -ForegroundColor $Colors.WarningMessage "Requested site was '$RequestedSiteUrl' but active connection reports '$connectedUrl'."
    }
}

Clear-Host

## Enforce TLS 1.2 minimum. Required for .NET Framework / PowerShell 5.1; .NET 6+ (PS7+) defaults to TLS 1.2+ already.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ($enableLog) {
    $logPath = Join-Path $scriptParentDir 'o365-connect-spo.txt'
    Write-Host "Script activity logged at $logPath"
    Start-Transcript $logPath | Out-Null
}

try {
    Write-Host -ForegroundColor $Colors.SystemMessage "SharePoint Online Connection script started`n"
    Write-Host -ForegroundColor $Colors.ProcessMessage "Prompt =", (-not $noprompt)

    if (($GenerateLocalCertificate -and $UseCertificateAuth) -or (-not $GenerateLocalCertificate -and -not $UseCertificateAuth)) {
        throw "Specify exactly one mode: -GenerateLocalCertificate or -UseCertificateAuth."
    }

    if (Get-Module -ListAvailable -Name PnP.PowerShell) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "PnP.PowerShell module installed"
    }
    else {
        Write-Host -ForegroundColor $Colors.WarningMessage -BackgroundColor $Colors.ErrorMessage "[001] - PnP.PowerShell module not installed`n"
        if (-not $noprompt) {
            do {
                $response = Read-Host -Prompt "`nDo you wish to install the PnP.PowerShell module (Y/N)?"
            } until (-not [string]::IsNullOrWhiteSpace($response))

            if ($response -ne 'Y' -and $response -ne 'y') {
                throw "PnP.PowerShell module is required."
            }
        }

        Write-Host -ForegroundColor $Colors.ProcessMessage "Installing PnP.PowerShell module - Administration escalation required"
        Start-Process $elevatedShellPath -Verb runAs -ArgumentList "Install-Module -Name PnP.PowerShell -Force -Confirm:`$false" -Wait -WindowStyle Hidden
        Write-Host -ForegroundColor $Colors.ProcessMessage "PnP.PowerShell module installed"
    }

    if (-not $noupdate) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Checking whether newer version of PnP.PowerShell module is available"
        $version          = Get-InstalledModule -Name PnP.PowerShell -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        $psgalleryVersion = Find-Module -Name PnP.PowerShell -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1

        $localVersion  = if ($null -ne $version)          { $version.Version -as [string] }          else { $null }
        $onlineVersion = if ($null -ne $psgalleryVersion) { $psgalleryVersion.Version -as [string] } else { $null }

        if ($null -eq $localVersion -or $null -eq $onlineVersion) {
            Write-Host -ForegroundColor $Colors.WarningMessage "Unable to compare module versions - skipping update check."
        }
        elseif ([version]$localVersion -lt [version]$onlineVersion) {
            Write-Host -ForegroundColor $Colors.WarningMessage "Local module $localVersion is lower than Gallery module $onlineVersion"
            if (-not $noprompt) {
                do {
                    $updateResponse = Read-Host -Prompt "`nDo you wish to update the PnP.PowerShell module (Y/N)?"
                } until (-not [string]::IsNullOrWhiteSpace($updateResponse))

                if ($updateResponse -eq 'Y' -or $updateResponse -eq 'y') {
                    Write-Host -ForegroundColor $Colors.ProcessMessage "Updating PnP.PowerShell module - Administration escalation required"
                    Start-Process $elevatedShellPath -Verb runAs -ArgumentList "Update-Module -Name PnP.PowerShell -Force -Confirm:`$false" -Wait -WindowStyle Hidden
                }
            }
            else {
                Write-Host -ForegroundColor $Colors.ProcessMessage "Updating PnP.PowerShell module - Administration escalation required"
                Start-Process $elevatedShellPath -Verb runAs -ArgumentList "Update-Module -Name PnP.PowerShell -Force -Confirm:`$false" -Wait -WindowStyle Hidden
            }
        }
        else {
            Write-Host -ForegroundColor $Colors.ProcessMessage "Local module $localVersion is current"
        }
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Loading PnP.PowerShell module"
    Import-PnPModuleWithCompat -Colors $Colors | Out-Null

    if ($GenerateLocalCertificate) {
        ## Resolve tenant before cert generation so the friendly name can include the tenant domain.
        $provisionTenant = $Tenant
        if ($ProvisionEntraApp) {
            if ([string]::IsNullOrWhiteSpace($provisionTenant) -and -not $noprompt) {
                do {
                    $provisionTenant = Read-Host -Prompt "Enter tenant ID or domain for Graph authentication (e.g. contoso.onmicrosoft.com)"
                } until (-not [string]::IsNullOrWhiteSpace($provisionTenant))
            }
            if ([string]::IsNullOrWhiteSpace($provisionTenant)) {
                throw "-ProvisionEntraApp requires -Tenant (tenant ID or .onmicrosoft.com domain)."
            }
        }

        $certFriendlyName     = if (-not [string]::IsNullOrWhiteSpace($provisionTenant)) { "$GeneratedCertSubject - $provisionTenant" } else { $GeneratedCertSubject }
        $generatedCertificate = New-SpoLocalCertificate -SubjectName $GeneratedCertSubject -YearsValid $GeneratedCertYearsValid -OutputPath $GeneratedCertOutputPath -ExportPfx:$ExportGeneratedPfx -PfxPassword $GeneratedPfxPassword -NoPrompt:$noprompt -FriendlyName $certFriendlyName

        Write-Host -ForegroundColor $Colors.ProcessMessage "Generated certificate thumbprint: $($generatedCertificate.Thumbprint)"
        Write-Host -ForegroundColor $Colors.ProcessMessage "Public certificate exported to: $($generatedCertificate.CerPath)"
        if (-not [string]::IsNullOrWhiteSpace($generatedCertificate.PfxPath)) {
            Write-Host -ForegroundColor $Colors.ProcessMessage "PFX exported to: $($generatedCertificate.PfxPath)"
        }

        if ($ProvisionEntraApp) {
            ## Display name for the new Entra app - defaults to cert subject.
            $resolvedDisplayName = if ([string]::IsNullOrWhiteSpace($AppDisplayName)) { $GeneratedCertSubject } else { $AppDisplayName }

            Write-Host -ForegroundColor $Colors.ProcessMessage "`nStarting Graph authentication for app provisioning..."
            $graphToken = Get-DeviceCodeGraphToken -TenantId $provisionTenant -ClientId $SetupClientId -CopyCodeToClipboard:$CopyDeviceCodeToClipboard -Colors $Colors

            ## Resolve default site URL after auth so we can use Graph tenant metadata where possible.
            $defaultSiteUrl = Resolve-SpoTenantRootSiteUrl -ExistingSiteUrl $SiteUrl -TenantHint $provisionTenant -AccessToken $graphToken -NoPrompt:$noprompt -Colors $Colors

            $provisionResult = Invoke-SpoAppProvisioning `
                -AccessToken $graphToken `
                -DisplayName $resolvedDisplayName `
                -CerFilePath $generatedCertificate.CerPath `
                -Thumbprint $generatedCertificate.Thumbprint `
                -SpoResourceAppId $SpoResourceAppId `
                -GraphResourceAppId $GraphResourceAppId `
                -Colors $Colors

            ## Update / create profile map entry.
            $mapPath      = $CertificateMapPath
            $profileEntry = [PSCustomObject]@{
                name                  = $resolvedDisplayName
                tenant                = $provisionTenant
                siteUrl               = $defaultSiteUrl
                appId                 = $provisionResult.AppId
                certificateThumbprint = $generatedCertificate.Thumbprint
            }

            Set-SpoProfileMapEntry -MapPath $mapPath -ProfileEntry $profileEntry -Colors $Colors

            Write-Host -ForegroundColor $Colors.SystemMessage "`n=== Provisioning complete ==="
            Write-Host -ForegroundColor $Colors.ProcessMessage "App ID:           $($provisionResult.AppId)"
            Write-Host -ForegroundColor $Colors.ProcessMessage "Cert Thumbprint:  $($generatedCertificate.Thumbprint)"
            Write-Host -ForegroundColor $Colors.ProcessMessage "Tenant:           $provisionTenant"
            Write-Host -ForegroundColor $Colors.ProcessMessage "Default Site URL: $defaultSiteUrl"
            Write-Host -ForegroundColor $Colors.WarningMessage "IMPORTANT: New app role grants can take 15-30 minutes to replicate across services."
            Write-Host -ForegroundColor $Colors.WarningMessage "Certificate-based Connect-PnPOnline attempts may fail during this window even when provisioning succeeded."
            Write-Host -ForegroundColor $Colors.ProcessMessage "`nConnect any time using:"
            Write-Host -ForegroundColor $Colors.ProcessMessage "  .\o365-connect-pnp-cert.ps1 -UseCertificateAuth -Tenant '$provisionTenant'"
        }

        Write-Host -ForegroundColor $Colors.SystemMessage "`nCertificate generation finished`n"
        exit 0
    }

    # --- UseCertificateAuth mode ---
    Write-Debug "Resolving profile and certificate auth inputs."
    $resolvedProfile = Resolve-SpoCertificateProfile -Path $CertificateMapPath -TenantFilter $Tenant -ProfileFilter $ProfileName -SiteUrlFilter $SiteUrl -NoPrompt:$noprompt -Colors $Colors
    if ($null -ne $resolvedProfile) {
        if ([string]::IsNullOrWhiteSpace($Tenant)) {
            $Tenant = $resolvedProfile.tenant
        }
        if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
            $SiteUrl = $resolvedProfile.siteUrl
        }
        if ([string]::IsNullOrWhiteSpace($AppId)) {
            $AppId = $resolvedProfile.appId
        }
        if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
            $CertificateThumbprint = $resolvedProfile.certificateThumbprint
        }
    }

    if ([string]::IsNullOrWhiteSpace($Tenant) -or [string]::IsNullOrWhiteSpace($SiteUrl) -or [string]::IsNullOrWhiteSpace($AppId) -or [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
        throw "UseCertificateAuth requires Tenant, SiteUrl, AppId, and CertificateThumbprint (directly or via CertificateMapPath profile)."
    }

    ## Verify the certificate is present in the local store before attempting to connect.
    $thumbprintLookup = $CertificateThumbprint.Trim()
    $localCert = $null
    try {
        $localCert = Get-Item "Cert:\CurrentUser\My\$thumbprintLookup" -ErrorAction Stop
    }
    catch {
        $localCert = $null
    }

    if ($null -eq $localCert) {
        $storeCerts = @(Get-ChildItem Cert:\CurrentUser\My -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $thumbprintLookup })
        if ($storeCerts.Count -gt 0) {
            $localCert = $storeCerts[0]
        }
    }

    if ($null -eq $localCert) {
        throw "Certificate with thumbprint '$CertificateThumbprint' not found in Cert:\CurrentUser\My. Import the PFX or run -GenerateLocalCertificate -ProvisionEntraApp on this machine first."
    }

    ## Warn if the cert expires within 30 days.
    $daysUntilExpiry = ($localCert.NotAfter - (Get-Date)).Days
    if ($daysUntilExpiry -le 0) {
        throw "Certificate '$CertificateThumbprint' expired on $($localCert.NotAfter.ToString('yyyy-MM-dd')). Provision a new certificate."
    }
    elseif ($daysUntilExpiry -le 30) {
        Write-Host -ForegroundColor $Colors.WarningMessage "Warning: Certificate expires in $daysUntilExpiry day(s) on $($localCert.NotAfter.ToString('yyyy-MM-dd'))."
    }
    else {
        Write-Debug "Certificate valid for $daysUntilExpiry more day(s)."
    }

    ## Check for any existing PnP connection.
    $existingConnection = $null
    try { $existingConnection = Get-PnPConnection -ErrorAction Stop } catch { }

    if ($null -ne $existingConnection) {
        $normalizedExisting  = $existingConnection.Url.TrimEnd('/')
        $normalizedRequested = $SiteUrl.TrimEnd('/')
        if ($normalizedExisting -eq $normalizedRequested) {
            Write-Host -ForegroundColor $Colors.ProcessMessage "Already connected to $SiteUrl - skipping reconnect."
            Write-SpoConnectedSite -RequestedSiteUrl $SiteUrl -Colors $Colors
            Write-Host -ForegroundColor $Colors.SystemMessage "SharePoint Online certificate auth flow finished`n"
            exit 0
        }
        else {
            Write-Host -ForegroundColor $Colors.WarningMessage "Currently connected to '$($existingConnection.Url)'. A new connection to '$SiteUrl' will be made and the existing connection will be ended."
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            Write-Host -ForegroundColor $Colors.ProcessMessage "Previous connection disconnected."
        }
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Connecting to SharePoint Online with certificate authentication"
    Connect-PnPOnline -Url $SiteUrl -ClientId $AppId -Thumbprint $CertificateThumbprint -Tenant $Tenant -ErrorAction Stop
    $disconnectCertificateAuthOnError = $true
    Write-SpoConnectedSite -RequestedSiteUrl $SiteUrl -Colors $Colors
    Write-SpoCertConnectionDetails -TenantId $Tenant -AppId $AppId -LocalCert $localCert -Colors $Colors
    $disconnectCertificateAuthOnError = $false

    Write-Host -ForegroundColor $Colors.ProcessMessage "Connected to SharePoint Online`n"
    Write-Host -ForegroundColor $Colors.SystemMessage "SharePoint Online certificate auth flow finished`n"
}
catch {
    if ($disconnectCertificateAuthOnError) {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    Write-Host -ForegroundColor $Colors.ErrorMessage "Script failed: $($_.Exception.Message)"
    if ($UseCertificateAuth -and $_.Exception.Message -match '(?i)access denied|forbidden|unauthorized|insufficient privileges|aadsts|permission|not authorized') {
        Write-Host -ForegroundColor $Colors.WarningMessage "If this app/certificate was just provisioned, wait 15-30 minutes and try again due to RBAC replication lag."
    }
    exit 1
}
finally {
    if ($enableLog) {
        Stop-Transcript | Out-Null
    }
}
