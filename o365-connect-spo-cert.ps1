[CmdletBinding()]  ## FIX #15: enables -Debug and -Verbose switches at script level
param(
    [switch]$noprompt = $false,   ## if -noprompt used then user will not be asked for any input
    [switch]$noupdate = $false,   ## if -noupdate used then module will not be checked for more recent version
    [switch]$enableLog = $false,  ## if -enableLog create a transcript log file

    [switch]$GenerateLocalCertificate = $false,
    [switch]$UseCertificateAuth = $false,

    [string]$GeneratedCertSubject = "O365-SPO-AppAuth",
    [int]$GeneratedCertYearsValid = 2,
    [string]$GeneratedCertOutputPath = "",  ## defaults to parent of script directory at runtime
    [switch]$ExportGeneratedPfx = $false,
    [securestring]$GeneratedPfxPassword,

    [string]$Tenant,
    [string]$ProfileName,
    [string]$AdminUrl,          ## SharePoint admin center URL (e.g. https://contoso-admin.sharepoint.com)
    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$CertificateMapPath = "",  ## defaults to o365-spo-admin-cert-auth.json in parent of script directory

    ## When used with -GenerateLocalCertificate, also create the Entra app, upload the cert,
    ## grant Sites.FullControl.All, and update the profile map automatically.
    [switch]$ProvisionEntraApp = $false,
    [string]$AppDisplayName = "",
    [string]$SetupClientId = "1950a258-227b-4e31-a9cf-717495945fc2",
    [switch]$CopyDeviceCodeToClipboard = $false
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.
Description - Simplified SharePoint Online admin center connect script with two modes:
1. GenerateLocalCertificate: create/export local cert files.
2. UseCertificateAuth: connect to SharePoint admin center with existing app/cert.
Source - https://github.com/directorcia/Office365/blob/master/o365-connect-spo-cert.ps1
Documentation - https://github.com/directorcia/Office365/wiki/Certificate-based-authentication-for-SharePoint-Online
#>

## Resolve paths relative to the script file itself, not the caller's working directory.
$scriptDir       = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptParentDir = Split-Path -Parent $scriptDir
if ([string]::IsNullOrWhiteSpace($GeneratedCertOutputPath)) { $GeneratedCertOutputPath = $scriptParentDir }

if ([string]::IsNullOrWhiteSpace($CertificateMapPath)) {
    ## Parent directory is searched first so reads are consistent with where writes land.
    $candidateCertificateMapPaths = @(
        (Join-Path $scriptParentDir 'cert-export/o365-spo-admin-cert-auth.json'),
        (Join-Path $scriptParentDir 'o365-spo-admin-cert-auth.json'),
        (Join-Path $scriptDir 'cert-export/o365-spo-admin-cert-auth.json'),
        (Join-Path $scriptDir 'o365-spo-admin-cert-auth.json')
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($candidatePath in $candidateCertificateMapPaths) {
        if (Test-Path -LiteralPath $candidatePath) {
            $CertificateMapPath = $candidatePath
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($CertificateMapPath)) {
        ## Default write location: parent directory, matching GeneratedCertOutputPath convention.
        $CertificateMapPath = Join-Path $scriptParentDir 'cert-export/o365-spo-admin-cert-auth.json'
    }
}

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

## Resolve the executable of the current host so elevated module installs target the same runtime
## (PS5 vs PS7) and module path as the running script.
$elevatedShellPath = (Get-Process -Id $PID).MainModule.FileName
if ([string]::IsNullOrWhiteSpace($elevatedShellPath) -or -not (Test-Path -LiteralPath $elevatedShellPath)) {
    throw "Unable to resolve current PowerShell host executable path for elevated module operations (PID $PID)."
}

function Resolve-SpoAdminCertificateProfile {
    <#
    .SYNOPSIS
        Load and filter the JSON certificate profile map, returning the matching profile entry.
    .OUTPUTS
        PSCustomObject  The selected profile entry, or $null if the map file is absent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$Path,
        [Parameter(Mandatory = $false)][string]$TenantFilter,
        [Parameter(Mandatory = $false)][string]$ProfileFilter,
        [Parameter(Mandatory = $false)][string]$AdminUrlFilter,
        [Parameter(Mandatory = $false)][switch]$NoPrompt,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

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

    $candidateProfiles = $profileItems
    if (-not [string]::IsNullOrWhiteSpace($ProfileFilter)) {
        $candidateProfiles = @($candidateProfiles | Where-Object { $_.name -eq $ProfileFilter })
    }
    if (-not [string]::IsNullOrWhiteSpace($TenantFilter)) {
        $candidateProfiles = @($candidateProfiles | Where-Object { $_.tenant -eq $TenantFilter })
    }
    if (-not [string]::IsNullOrWhiteSpace($AdminUrlFilter)) {
        $candidateProfiles = @($candidateProfiles | Where-Object { $_.adminUrl -eq $AdminUrlFilter })
    }

    if ($candidateProfiles.Count -eq 0) {
        if ($profileItems.Count -eq 1) {
            Write-Debug "No exact profile match found for the supplied filters; using the only available profile."
            return $profileItems[0]
        }

        $appliedFilters = @()
        if (-not [string]::IsNullOrWhiteSpace($ProfileFilter))  { $appliedFilters += "ProfileName='$ProfileFilter'" }
        if (-not [string]::IsNullOrWhiteSpace($TenantFilter))   { $appliedFilters += "Tenant='$TenantFilter'" }
        if (-not [string]::IsNullOrWhiteSpace($AdminUrlFilter)) { $appliedFilters += "AdminUrl='$AdminUrlFilter'" }
        $filterDesc = if ($appliedFilters.Count -gt 0) { " (filters: $($appliedFilters -join ', '))" } else { " (no filters applied)" }

        $availableDesc = ($profileItems | ForEach-Object {
            "name='$($_.name)' tenant='$($_.tenant)' adminUrl='$($_.adminUrl)'"
        }) -join '; '

        throw "No matching certificate profile found in '$Path'$filterDesc. Available profiles: [$availableDesc]"
    }

    if ($candidateProfiles.Count -eq 1 -or $NoPrompt) {
        if ($candidateProfiles.Count -gt 1 -and $NoPrompt) {
            throw "Multiple matching profiles found in '$Path'. Specify -ProfileName, -Tenant, or -AdminUrl."
        }
        return $candidateProfiles[0]
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Multiple matching certificate profiles found:"
    for ($index = 0; $index -lt $candidateProfiles.Count; $index++) {
        $displayName = if ([string]::IsNullOrWhiteSpace($candidateProfiles[$index].name)) { "(unnamed)" } else { $candidateProfiles[$index].name }
        Write-Host -ForegroundColor $Colors.ProcessMessage ("[{0}] {1} | Tenant={2} | AdminUrl={3} | AppId={4}" -f ($index + 1), $displayName, $candidateProfiles[$index].tenant, $candidateProfiles[$index].adminUrl, $candidateProfiles[$index].appId)
    }

    do {
        $choice = Read-Host -Prompt "Select profile number"
        [int]$parsedChoice = 0
        $validSelection = [int]::TryParse($choice, [ref]$parsedChoice) -and $parsedChoice -ge 1 -and $parsedChoice -le $candidateProfiles.Count
    } until ($validSelection)

    return $candidateProfiles[$parsedChoice - 1]
}

function Resolve-SpoAdminCertificate {
    <#
    .SYNOPSIS
        Resolve a certificate from the current user or local machine personal store by thumbprint.
        If the private-key certificate is not yet present in the store, this will attempt to import
        it from a previously exported PFX file.
    .OUTPUTS
        X509Certificate2  The resolved certificate object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Thumbprint,
        [Parameter(Mandatory = $true)][hashtable]$Colors,
        [Parameter(Mandatory = $false)][string[]]$CandidatePfxPaths = @()
    )

    if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
        throw "Certificate thumbprint was empty."
    }

    $normalizedThumbprint = ($Thumbprint -replace '\s', '').ToUpperInvariant()
    $storePaths = @(
        'Cert:\CurrentUser\My',
        'Cert:\LocalMachine\My'
    )

    foreach ($storePath in $storePaths) {
        Write-Debug "Searching certificate store '$storePath' for thumbprint '$Thumbprint'."
        $matchingCert = Get-ChildItem -Path $storePath -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Thumbprint -and
                ((($_.Thumbprint) -replace '\s', '').ToUpperInvariant()) -eq $normalizedThumbprint
            } |
            Select-Object -First 1

        if ($null -ne $matchingCert) {
            if (-not $matchingCert.HasPrivateKey) {
                throw "Certificate '$Thumbprint' was found but does not expose a private key. Import the matching PFX or ensure the certificate was created with a private key."
            }

            Write-Debug "Resolved certificate '$Thumbprint' from $storePath."
            return $matchingCert
        }
    }

    $pfxSearchRoots = @(
        $script:scriptDir,
        $script:scriptParentDir,
        $env:TEMP,
        $PWD.Path
    ) + @($CandidatePfxPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    foreach ($searchRoot in ($pfxSearchRoots | Select-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace($searchRoot) -or -not (Test-Path -LiteralPath $searchRoot)) {
            continue
        }

        $candidatePfxs = @()
        if ((Test-Path -LiteralPath $searchRoot) -and (Get-Item -LiteralPath $searchRoot -ErrorAction SilentlyContinue).PSIsContainer) {
            $candidatePfxs = @(Get-ChildItem -Path $searchRoot -Filter '*.pfx' -File -Recurse -ErrorAction SilentlyContinue)
        }
        elseif ($searchRoot -like '*.pfx') {
            $candidatePfxs = @(Get-Item -LiteralPath $searchRoot -ErrorAction SilentlyContinue)
        }

        foreach ($candidatePfx in $candidatePfxs) {
            $candidateName = $candidatePfx.Name
            if ($candidateName -notmatch [regex]::Escape($normalizedThumbprint) -and $candidateName -notmatch [regex]::Escape(($Thumbprint -replace '\s', ''))) {
                continue
            }

            try {
                Write-Host -ForegroundColor $Colors.ProcessMessage "Importing certificate from PFX: $($candidatePfx.FullName)"
                $emptyPassword = [System.Security.SecureString]::new()
                $importedCert = Import-PfxCertificate -FilePath $candidatePfx.FullName -CertStoreLocation 'Cert:\CurrentUser\My' -Password $emptyPassword -Exportable -ErrorAction Stop
                if ($null -ne $importedCert) {
                    if (-not $importedCert.HasPrivateKey) {
                        throw "Imported certificate '$Thumbprint' does not expose a private key."
                    }
                    Write-Debug "Resolved certificate '$Thumbprint' from imported PFX $($candidatePfx.FullName)."
                    return $importedCert
                }
            }
            catch {
                Write-Debug "Unable to import certificate from PFX '$($candidatePfx.FullName)': $($_.Exception.Message)"
            }
        }
    }

    throw "Certificate with thumbprint '$Thumbprint' was not found in Cert:\CurrentUser\My or Cert:\LocalMachine\My and no importable PFX was found. Import the matching PFX or run -GenerateLocalCertificate on this machine first."
}

function Resolve-SpoAdminCertificateThumbprint {
    <#
    .SYNOPSIS
        Find a suitable local certificate thumbprint when one was not explicitly supplied.
    .OUTPUTS
        String  The resolved thumbprint or $null if no suitable certificate was found.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$PreferredSubject = "",
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $candidateSubjects = @()
    if (-not [string]::IsNullOrWhiteSpace($PreferredSubject)) {
        $candidateSubjects += $PreferredSubject
        $candidateSubjects += "CN=$PreferredSubject"
        $candidateSubjects += "O365-SPO-AppAuth"
    }
    else {
        $candidateSubjects += "O365-SPO-AppAuth"
    }

    $storePaths = @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')

    foreach ($storePath in $storePaths) {
        ## FIX #2: Force array with @() so .Count is reliable when Get-ChildItem returns a single object (PS5.1).
        $matchingCerts = @(Get-ChildItem -Path $storePath -ErrorAction SilentlyContinue |
            Where-Object { $_.HasPrivateKey -and $_.Thumbprint })

        if ($matchingCerts.Count -gt 0) {
            $preferredMatches = @($matchingCerts | Where-Object {
                $subjectText = [string]$_.Subject
                $friendlyName = [string]$_.FriendlyName
                foreach ($subject in $candidateSubjects) {
                    if ($subjectText -like "*$subject*" -or $friendlyName -like "*$subject*") {
                        return $true
                    }
                }
                return $false
            })

            if ($preferredMatches.Count -gt 0) {
                $selectedCert = $preferredMatches | Sort-Object NotAfter -Descending | Select-Object -First 1
                return $selectedCert.Thumbprint
            }

            $selectedCert = $matchingCerts | Sort-Object NotAfter -Descending | Select-Object -First 1
            if ($null -ne $selectedCert) {
                return $selectedCert.Thumbprint
            }
        }
    }

    return $null
}

function New-SpoAdminLocalCertificate {
    <#
    .SYNOPSIS
        Generate a self-signed RSA-2048 certificate for SharePoint Online admin app authentication.
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

    if (-not (Test-Path -Path $OutputPath)) {
        Write-Debug "Creating certificate output directory: $OutputPath"
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    $certificate = New-SelfSignedCertificate -Subject "CN=$SubjectName" -CertStoreLocation "Cert:\CurrentUser\My" -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears($YearsValid) -KeySpec Signature -ErrorAction Stop

    ## FIX #5: FriendlyName is Windows-only; guard against failure on PS7/Linux/macOS.
    $resolvedFriendlyName = if ([string]::IsNullOrWhiteSpace($FriendlyName)) { $SubjectName } else { $FriendlyName }
    try {
        $certificate.FriendlyName = $resolvedFriendlyName
        Write-Debug "Certificate friendly name set to: $resolvedFriendlyName"
    }
    catch {
        Write-Debug "Could not set FriendlyName (non-Windows platform or restricted store): $($_.Exception.Message)"
    }

    $safeSubject = ($SubjectName -replace '[^A-Za-z0-9\-_.]', '-')
    $fileBase    = "{0}-{1}" -f $safeSubject, $certificate.Thumbprint
    $cerPath     = Join-Path -Path $OutputPath -ChildPath "$fileBase.cer"

    Export-Certificate -Cert $certificate -FilePath $cerPath -Type CERT -Force -ErrorAction Stop | Out-Null

    $pfxPath = $null
    if ($ExportPfx) {
        $securePfxPassword = $PfxPassword

        if ($null -eq $securePfxPassword) {
            Write-Debug "No explicit PFX password supplied; using an empty password for local import compatibility."
            ## FIX #9: Warn prominently that an unprotected PFX is being written to disk.
            Write-Host -ForegroundColor "yellow" "WARNING: PFX will be exported with an EMPTY password. The private key is unprotected on disk. Secure or delete the PFX file after import."
            $securePfxPassword = [System.Security.SecureString]::new()
        }

        $pfxPath = Join-Path -Path $OutputPath -ChildPath "$fileBase.pfx"
        Export-PfxCertificate -Cert $certificate -FilePath $pfxPath -Password $securePfxPassword -Force -ErrorAction Stop | Out-Null
    }

    return [PSCustomObject]@{
        Thumbprint = $certificate.Thumbprint
        Subject    = $certificate.Subject
        NotAfter   = $certificate.NotAfter
        CerPath    = $cerPath
        PfxPath    = if ($null -ne $pfxPath) { $pfxPath } else { "" }  ## FIX #6: store empty string instead of $null for safe JSON serialisation
    }
}

function Get-DeviceCodeGraphToken {
    <#
    .SYNOPSIS
        Authenticate to Microsoft Graph using the device-code OAuth2 flow.
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

    Start-Process $deviceCodeResponse.verification_uri

    $tokenBody = @{
        grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
        client_id   = $ClientId
        device_code = $deviceCodeResponse.device_code
    }

    $deadline     = (Get-Date).AddSeconds($deviceCodeResponse.expires_in)
    ## FIX #7: Explicitly cast to [int] to avoid type-widening from JSON long on PS7.
    $pollInterval = [int]$deviceCodeResponse.interval

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

function Invoke-SpoAdminGraphRequest {
    <#
    .SYNOPSIS
        Helper: make an authenticated Graph REST call and return the parsed response.
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

            $statusCode        = $null
            $retryAfterSeconds = $null
            $response          = $null
            try { $response = $_.Exception.Response } catch {}
            if ($null -ne $response) {
                try { $statusCode = [int]$response.StatusCode } catch {}
                try {
                    ## FIX #4: Normalise Retry-After header access for PS5.1 (WebHeaderCollection) and PS7 (IEnumerable<string>).
                    $retryAfterRaw = $null
                    try {
                        ## PS7: HttpResponseMessage.Headers is a typed collection; use GetValues() if available.
                        $retryAfterRaw = $response.Headers.GetValues('Retry-After') | Select-Object -First 1
                    }
                    catch {
                        ## PS5.1: WebHeaderCollection supports string indexing directly.
                        try { $retryAfterRaw = [string]$response.Headers['Retry-After'] } catch {}
                    }

                    if (-not [string]::IsNullOrWhiteSpace($retryAfterRaw)) {
                        [int]$parsedRetryAfter = 0
                        if ([int]::TryParse($retryAfterRaw.Trim(), [ref]$parsedRetryAfter) -and $parsedRetryAfter -gt 0) {
                            $retryAfterSeconds = $parsedRetryAfter
                        }
                    }
                }
                catch {}
            }

            $isRetriableStatus  = ($statusCode -in @(429, 500, 502, 503, 504))
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

function Set-SpoAdminProfileMapEntry {
    <#
    .SYNOPSIS
        Atomically upsert a certificate profile entry in the JSON profile map file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$MapPath,
        [Parameter(Mandatory = $true)][object]$ProfileEntry,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $fullMapPath = [System.IO.Path]::GetFullPath($MapPath)

    ## FIX #11: Dispose SHA256 instance properly to avoid unmanaged resource leak.
    $sha256    = [System.Security.Cryptography.SHA256]::Create()
    $hashBytes = $null
    try {
        $hashBytes = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($fullMapPath.ToLowerInvariant()))
    }
    finally {
        $sha256.Dispose()
    }

    $hashHex   = ([System.BitConverter]::ToString($hashBytes)).Replace('-', '')
    $mutexName = "Global\CIAOPS_SPO_ADMIN_PROFILEMAP_$hashHex"

    $mutex     = New-Object System.Threading.Mutex($false, $mutexName)
    $hasHandle = $false
    $tempPath  = "$fullMapPath.$PID.tmp"

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

        ## Replace existing entry for the same tenant or appId, or append.
        $existingIdx = $null
        for ($i = 0; $i -lt $profileList.Count; $i++) {
            $sameApp    = (-not [string]::IsNullOrWhiteSpace($profileList[$i].appId)  -and $profileList[$i].appId   -eq $ProfileEntry.appId)
            $sameTenant = (-not [string]::IsNullOrWhiteSpace($profileList[$i].tenant) -and $profileList[$i].tenant  -eq $ProfileEntry.tenant)
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
    $now = [int][DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

    ## JWT header and payload - both must be base64url encoded.
    $headerJson  = '{"alg":"RS256","typ":"JWT","x5t":"' + $x5t + '"}'
    $payloadJson = '{"aud":"' + $tokenEndpoint + '","iss":"' + $AppId + '","sub":"' + $AppId + '","jti":"' + [System.Guid]::NewGuid().ToString() + '","nbf":' + $now + ',"exp":' + ($now + 600) + '}'

    $headerB64    = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($headerJson)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $payloadB64   = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payloadJson)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $signingInput = [System.Text.Encoding]::UTF8.GetBytes("$headerB64.$payloadB64")

    ## Sign with the certificate's RSA private key using PKCS#1 v1.5 / SHA-256.
    $rsa      = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    $sigBytes = $rsa.SignData($signingInput, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $sigB64   = [System.Convert]::ToBase64String($sigBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    $clientAssertion = "$headerB64.$payloadB64.$sigB64"

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

function Write-SpoAdminCertConnectionDetails {
    <#
    .SYNOPSIS
        Display local certificate details and the matching Entra ID app/keyCredential after connecting.
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
        $appResult  = Invoke-SpoAdminGraphRequest -AccessToken $graphToken -Method Get `
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

function Get-SpoAdminProvisioningRoleTargets {
    <#
    .SYNOPSIS
        Resolve the SharePoint Online and Microsoft Graph service principals and required app role IDs.
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
    $spoSpResult = Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals?`$filter=$spoSpFilter"
    if ($null -eq $spoSpResult.value -or $spoSpResult.value.Count -eq 0) {
        throw "SharePoint Online service principal not found. Ensure SharePoint Online is provisioned in this tenant."
    }
    $spoSp            = $spoSpResult.value[0]
    $sitesFullRole    = $spoSp.appRoles | Where-Object { $_.value -eq "Sites.FullControl.All" }
    if ($null -eq $sitesFullRole) {
        throw "Sites.FullControl.All role not found on SharePoint Online service principal."
    }
    Write-Debug "Sites.FullControl.All role ID: $($sitesFullRole.id)"

    Write-Host -ForegroundColor $Colors.ProcessMessage "Locating Microsoft Graph service principal in directory..."
    $graphSpFilter    = [uri]::EscapeDataString("appId eq '$GraphResourceAppId'")
    $graphSpResult    = Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals?`$filter=$graphSpFilter"
    if ($null -eq $graphSpResult.value -or $graphSpResult.value.Count -eq 0) {
        throw "Microsoft Graph service principal not found in this tenant."
    }
    $graphSp          = $graphSpResult.value[0]
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

function Get-OrCreateSpoAdminEntraApplication {
    <#
    .SYNOPSIS
        Return an existing Entra app registration by display name, or create a new one.
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
    $appFilter    = [uri]::EscapeDataString("displayName eq '" + ($DisplayName -replace "'", "''") + "'")
    $existingApps = Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/applications?`$filter=$appFilter"

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
                ## SharePoint Online
                resourceAppId  = $SpoResourceAppId
                resourceAccess = @(
                    @{ id = $SpoSitesFullRoleId; type = "Role" }
                )
            },
            @{
                ## Microsoft Graph
                resourceAppId  = $GraphResourceAppId
                resourceAccess = @(
                    @{ id = $GraphReadAllRoleId; type = "Role" }
                )
            }
        )
    }
    $appObject = Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/applications" -Body $newAppBody
    Write-Host -ForegroundColor $Colors.ProcessMessage "App created. Object ID: $($appObject.id) | App ID: $($appObject.appId)"
    return $appObject
}

function Set-SpoAdminEntraApplicationCertificate {
    <#
    .SYNOPSIS
        Upload a local certificate to an Entra app registration and verify it was stored.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AppObjectId,
        ## FIX #3: Removed strict 40-hex-char regex so thumbprints with spaces (Windows store quirk)
        ##         are normalised rather than rejected by ValidatePattern.
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Thumbprint,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $graphBase = "https://graph.microsoft.com/v1.0"

    ## Normalise the thumbprint before looking up the cert store.
    $normalizedThumbprint = ($Thumbprint -replace '\s', '').ToUpperInvariant()

    Write-Host -ForegroundColor $Colors.ProcessMessage "Uploading certificate to app registration..."

    $certStoreObj = Get-Item "Cert:\CurrentUser\My\$normalizedThumbprint" -ErrorAction SilentlyContinue
    if ($null -eq $certStoreObj) {
        throw "Certificate '$normalizedThumbprint' not found in Cert:\CurrentUser\My. Ensure the certificate was generated on this machine or imported first."
    }

    $cerBytes          = $certStoreObj.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    $cerBase64         = [System.Convert]::ToBase64String($cerBytes)
    $customKeyIdBase64 = [System.Convert]::ToBase64String($certStoreObj.GetCertHash())

    if ($cerBytes.Length -eq 0) {
        throw "Certificate export produced empty bytes for thumbprint '$normalizedThumbprint'. Cert may not be in Cert:\CurrentUser\My."
    }
    Write-Debug "Cert bytes: $($cerBytes.Length) | customKeyIdentifier: $customKeyIdBase64"

    ## Fetch existing certs so we append rather than replace (supports multiple machines / certs per app).
    $existingApp  = Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/applications/${AppObjectId}?`$select=keyCredentials"
    $existingKeys = @($existingApp.keyCredentials | Where-Object { $_.customKeyIdentifier -ne $customKeyIdBase64 })
    Write-Debug "Existing key credentials on app (excluding this thumbprint): $($existingKeys.Count)"

    $newKey = @{
        type                = "AsymmetricX509Cert"
        usage               = "Verify"
        key                 = $cerBase64
        displayName         = "O365-SPO-AppAuth"
        customKeyIdentifier = $customKeyIdBase64
    }

    $certPatch = @{ keyCredentials = @($existingKeys) + @($newKey) }
    Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Patch -Uri "$graphBase/applications/$AppObjectId" -Body $certPatch | Out-Null

    ## Verify the credential was actually stored with the correct customKeyIdentifier.
    $appVerify = Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/applications/${AppObjectId}?`$select=keyCredentials"
    $storedKey = $appVerify.keyCredentials | Where-Object { $_.customKeyIdentifier -eq $customKeyIdBase64 } | Select-Object -First 1
    if ($null -eq $storedKey) {
        $stored = ($appVerify.keyCredentials | ForEach-Object { $_.customKeyIdentifier }) -join ', '
        throw "Certificate upload verification failed. Expected customKeyIdentifier '$customKeyIdBase64' not found in app keyCredentials. Stored: $stored"
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Certificate uploaded and verified in app registration."
}

function Get-OrCreateSpoAdminEntraServicePrincipal {
    <#
    .SYNOPSIS
        Return the service principal for the given app ID, creating it if it does not yet exist.
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
    $spFilter   = [uri]::EscapeDataString("appId eq '$AppId'")
    $existingSp = Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals?`$filter=$spFilter"

    if ($existingSp.value.Count -gt 0) {
        $spObject = $existingSp.value[0]
        Write-Host -ForegroundColor $Colors.ProcessMessage "Service principal already exists. Object ID: $($spObject.id)"
        return $spObject
    }

    $spObject = Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/servicePrincipals" -Body @{ appId = $AppId }
    Write-Host -ForegroundColor $Colors.ProcessMessage "Service principal created. Object ID: $($spObject.id)"
    return $spObject
}

function Set-SpoAdminEntraAppRoleAssignments {
    <#
    .SYNOPSIS
        Grant Sites.FullControl.All and Graph Application.Read.All to the provisioned service principal.
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
    $spoAssignments = Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals/$ServicePrincipalId/appRoleAssignments"
    $alreadyGranted = $spoAssignments.value | Where-Object { $_.appRoleId -eq $SpoSitesFullRoleId -and $_.resourceId -eq $SpoServicePrincipalId }
    if ($null -ne $alreadyGranted) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Sites.FullControl.All already granted - skipping."
    }
    else {
        $roleBody = @{ principalId = $ServicePrincipalId; resourceId = $SpoServicePrincipalId; appRoleId = $SpoSitesFullRoleId }
        Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/servicePrincipals/$ServicePrincipalId/appRoleAssignments" -Body $roleBody | Out-Null
        Write-Host -ForegroundColor $Colors.ProcessMessage "Sites.FullControl.All granted."
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Granting Microsoft Graph Application.Read.All permission (admin consent)..."
    ## Re-fetch assignments after the SPO grant check/post so evaluation stays correct if grant order changes.
    $graphAssignments    = Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals/$ServicePrincipalId/appRoleAssignments"
    $graphAlreadyGranted = $graphAssignments.value | Where-Object { $_.appRoleId -eq $GraphReadAllRoleId -and $_.resourceId -eq $GraphServicePrincipalId }
    if ($null -ne $graphAlreadyGranted) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Microsoft Graph Application.Read.All already granted - skipping."
    }
    else {
        $graphRoleBody = @{ principalId = $ServicePrincipalId; resourceId = $GraphServicePrincipalId; appRoleId = $GraphReadAllRoleId }
        Invoke-SpoAdminGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/servicePrincipals/$ServicePrincipalId/appRoleAssignments" -Body $graphRoleBody | Out-Null
        Write-Host -ForegroundColor $Colors.ProcessMessage "Microsoft Graph Application.Read.All granted."
    }
}

function Invoke-SpoAdminEntraAppProvisioning {
    <#
    .SYNOPSIS
        Orchestrates Entra app provisioning for SharePoint admin by invoking focused helper functions.
    .OUTPUTS
        PSCustomObject  AppId, AppObjId, SpObjId.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DisplayName,
        [Parameter(Mandatory = $true)][string]$CerFilePath,
        ## FIX #3: Accept thumbprints with spaces; normalise inside the function.
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Thumbprint,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$SpoResourceAppId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$GraphResourceAppId,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    ## Provisioning uses static .NET method calls that are blocked in Constrained Language Mode.
    if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
        throw "App provisioning requires FullLanguage mode. Current mode: $($ExecutionContext.SessionState.LanguageMode). Run the script in a standard PowerShell session (not one with Constrained Language Mode active)."
    }

    if (-not [string]::IsNullOrWhiteSpace($CerFilePath) -and -not (Test-Path -Path $CerFilePath)) {
        throw "Certificate file path '$CerFilePath' was provided but does not exist."
    }

    ## Normalise before passing downstream.
    $normalizedThumbprint = ($Thumbprint -replace '\s', '').ToUpperInvariant()

    $targets   = Get-SpoAdminProvisioningRoleTargets -AccessToken $AccessToken -SpoResourceAppId $SpoResourceAppId -GraphResourceAppId $GraphResourceAppId -Colors $Colors
    $appObject = Get-OrCreateSpoAdminEntraApplication `
        -AccessToken        $AccessToken `
        -DisplayName        $DisplayName `
        -SpoSitesFullRoleId $targets.SpoSitesFullRole.id `
        -GraphReadAllRoleId $targets.GraphReadAllRole.id `
        -SpoResourceAppId   $SpoResourceAppId `
        -GraphResourceAppId $GraphResourceAppId `
        -Colors             $Colors

    Set-SpoAdminEntraApplicationCertificate -AccessToken $AccessToken -AppObjectId $appObject.id -Thumbprint $normalizedThumbprint -Colors $Colors

    $spObject = Get-OrCreateSpoAdminEntraServicePrincipal -AccessToken $AccessToken -AppId $appObject.appId -Colors $Colors

    Set-SpoAdminEntraAppRoleAssignments `
        -AccessToken             $AccessToken `
        -ServicePrincipalId      $spObject.id `
        -SpoServicePrincipalId   $targets.SpoServicePrincipal.id `
        -SpoSitesFullRoleId      $targets.SpoSitesFullRole.id `
        -GraphServicePrincipalId $targets.GraphServicePrincipal.id `
        -GraphReadAllRoleId      $targets.GraphReadAllRole.id `
        -Colors                  $Colors

    return [PSCustomObject]@{
        AppId    = $appObject.appId
        AppObjId = $appObject.id
        SpObjId  = $spObject.id
    }
}

function Resolve-SpoAdminUrl {
    <#
    .SYNOPSIS
        Derive the SharePoint admin center URL from a tenant identifier.
        Handles .onmicrosoft.com domains, tenant GUIDs (no URL derivable), explicit URLs,
        and custom domains. Returns the supplied ExplicitUrl unchanged when provided.
    .OUTPUTS
        String  The resolved admin center URL, or an empty string when it cannot be determined.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$ExplicitUrl = "",
        [Parameter(Mandatory = $false)][string]$TenantId = ""
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitUrl)) {
        Write-Debug "Using explicitly supplied admin URL: $ExplicitUrl"
        return $ExplicitUrl
    }

    if ([string]::IsNullOrWhiteSpace($TenantId)) {
        return ""
    }

    if ($TenantId -match '^([^.]+)\.onmicrosoft\.com$') {
        $url = "https://$($Matches[1])-admin.sharepoint.com"
        Write-Debug "Derived admin URL from .onmicrosoft.com tenant: $url"
        return $url
    }

    if ($TenantId -match '^https?://.+') {
        Write-Debug "Tenant value is already a URL: $TenantId"
        return $TenantId
    }

    ## Tenant GUID - no hostname to derive a URL from.
    if ($TenantId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
        Write-Debug "Tenant value is a GUID; cannot derive admin URL without a domain. Supply -AdminUrl explicitly."
        return ""
    }

    ## Best-effort fallback for custom domains (e.g. contoso.com -> https://contoso-admin.sharepoint.com).
    if ($TenantId -match '^([^.]+)\..+$') {
        $url = "https://$($Matches[1])-admin.sharepoint.com"
        Write-Debug "Derived admin URL from custom domain tenant: $url"
        return $url
    }

    Write-Debug "Could not derive admin URL from tenant value: $TenantId"
    return ""
}

function Write-SpoAdminConnectedCenter {
    <#
    .SYNOPSIS
        Prints the active SharePoint admin center URL and tenant title.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$RequestedAdminUrl,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $connectedTitle = $null

    try {
        $tenant = Get-SPOTenant -ErrorAction Stop
        if ($null -ne $tenant) {
            $connectedTitle = $tenant.Title
        }
    }
    catch {
        ## FIX #13: Surface the failure at debug level so transient errors aren't silently swallowed.
        Write-Debug "Get-SPOTenant failed or no active connection: $($_.Exception.Message)"
    }

    $titleDisplay = if (-not [string]::IsNullOrWhiteSpace($connectedTitle)) { " ($connectedTitle)" } else { "" }
    $displayUrl   = if (-not [string]::IsNullOrWhiteSpace($RequestedAdminUrl)) { $RequestedAdminUrl } else { "(unknown)" }
    Write-Host -ForegroundColor $Colors.SystemMessage "Connected SharePoint admin center: ${displayUrl}${titleDisplay}"
}

Clear-Host

## Enforce TLS 1.2 minimum.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ($enableLog) {
    $logPath = Join-Path $scriptParentDir 'o365-connect-spo-admin.txt'
    Write-Host "Script activity logged at $logPath"
    Start-Transcript $logPath | Out-Null
}

try {
    Write-Host -ForegroundColor $Colors.SystemMessage "SharePoint Online Admin Center Connection script started`n"
    Write-Host -ForegroundColor $Colors.ProcessMessage "Prompt =", (-not $noprompt)

    if (($GenerateLocalCertificate -and $UseCertificateAuth) -or (-not $GenerateLocalCertificate -and -not $UseCertificateAuth)) {
        throw "Specify exactly one mode: -GenerateLocalCertificate or -UseCertificateAuth."
    }

    if (Get-Module -ListAvailable -Name microsoft.online.sharepoint.powershell) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "SharePoint Online PowerShell module installed"
    }
    else {
        Write-Host -ForegroundColor $Colors.WarningMessage -BackgroundColor $Colors.ErrorMessage "[001] - SharePoint Online PowerShell module not installed`n"
        if (-not $noprompt) {
            do {
                $response = Read-Host -Prompt "`nDo you wish to install the SharePoint Online PowerShell module (Y/N)?"
            } until (-not [string]::IsNullOrWhiteSpace($response))

            if ($response -ne 'Y' -and $response -ne 'y') {
                throw "SharePoint Online PowerShell module is required."
            }
        }

        Write-Host -ForegroundColor $Colors.ProcessMessage "Installing SharePoint Online PowerShell module - Administration escalation required"
        ## FIX #16: Capture exit code from elevated install; fail early with a clear message if UAC was denied or install failed.
        $installProcess = Start-Process $elevatedShellPath -Verb runAs -ArgumentList "Install-Module -Name microsoft.online.sharepoint.powershell -Force -Confirm:`$false" -Wait -WindowStyle Hidden -PassThru
        if ($installProcess.ExitCode -ne 0) {
            throw "Elevated module install failed (exit code $($installProcess.ExitCode)). Run PowerShell as Administrator and retry, or install the module manually: Install-Module -Name microsoft.online.sharepoint.powershell"
        }
        Write-Host -ForegroundColor $Colors.ProcessMessage "SharePoint Online PowerShell module installed"
    }

    if (-not $noupdate) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Checking whether newer version of SharePoint Online PowerShell module is available"
        $version          = Get-InstalledModule -Name microsoft.online.sharepoint.powershell -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        $psgalleryVersion = Find-Module -Name microsoft.online.sharepoint.powershell -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1

        $localVersion  = if ($null -ne $version)          { $version.Version -as [string] }          else { $null }
        $onlineVersion = if ($null -ne $psgalleryVersion) { $psgalleryVersion.Version -as [string] } else { $null }

        if ($null -eq $localVersion -or $null -eq $onlineVersion) {
            Write-Host -ForegroundColor $Colors.WarningMessage "Unable to compare module versions - skipping update check."
        }
        elseif ([version]$localVersion -lt [version]$onlineVersion) {
            Write-Host -ForegroundColor $Colors.WarningMessage "Local module $localVersion is lower than Gallery module $onlineVersion"
            if (-not $noprompt) {
                do {
                    $updateResponse = Read-Host -Prompt "`nDo you wish to update the SharePoint Online PowerShell module (Y/N)?"
                } until (-not [string]::IsNullOrWhiteSpace($updateResponse))

                if ($updateResponse -eq 'Y' -or $updateResponse -eq 'y') {
                    Write-Host -ForegroundColor $Colors.ProcessMessage "Updating SharePoint Online PowerShell module - Administration escalation required"
                    ## FIX #16: Capture and check exit code for update as well.
                    $updateProcess = Start-Process $elevatedShellPath -Verb runAs -ArgumentList "Update-Module -Name microsoft.online.sharepoint.powershell -Force -Confirm:`$false" -Wait -WindowStyle Hidden -PassThru
                    if ($updateProcess.ExitCode -ne 0) {
                        Write-Host -ForegroundColor $Colors.WarningMessage "Module update may have failed (exit code $($updateProcess.ExitCode)). Continuing with current installed version."
                    }
                }
            }
            else {
                Write-Host -ForegroundColor $Colors.ProcessMessage "Updating SharePoint Online PowerShell module - Administration escalation required"
                $updateProcess = Start-Process $elevatedShellPath -Verb runAs -ArgumentList "Update-Module -Name microsoft.online.sharepoint.powershell -Force -Confirm:`$false" -Wait -WindowStyle Hidden -PassThru
                if ($updateProcess.ExitCode -ne 0) {
                    Write-Host -ForegroundColor $Colors.WarningMessage "Module update may have failed (exit code $($updateProcess.ExitCode)). Continuing with current installed version."
                }
            }
        }
        else {
            Write-Host -ForegroundColor $Colors.ProcessMessage "Local module $localVersion is current"
        }
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Loading SharePoint Online PowerShell module"
    $ps = $PSVersionTable.PSVersion
    if ($ps.Major -lt 6) {
        Import-Module microsoft.online.sharepoint.powershell -DisableNameChecking -ErrorAction Stop | Out-Null
    }
    else {
        ## Prefer native import in PowerShell 6+ so full cmdlet parameter sets are available.
        try {
            Write-Host -ForegroundColor $Colors.ProcessMessage "Using native module load for PowerShell 6+ (SkipEditionCheck)"
            Import-Module microsoft.online.sharepoint.powershell -DisableNameChecking -SkipEditionCheck -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Host -ForegroundColor $Colors.WarningMessage "Native load failed. Falling back to compatibility mode for PowerShell 6+"
            Import-Module microsoft.online.sharepoint.powershell -DisableNameChecking -UseWindowsPowerShell -ErrorAction Stop | Out-Null
        }
    }

    $connectSpoCommand = Get-Command -Name Connect-SPOService -ErrorAction Stop
    $connectParamKeys  = @($connectSpoCommand.Parameters.Keys)
    $hasAppParam       = ($connectParamKeys -contains 'ClientId' -or $connectParamKeys -contains 'AppId')
    $hasThumbprintParam= ($connectParamKeys -contains 'Thumbprint' -or $connectParamKeys -contains 'CertificateThumbprint')
    Write-Debug "Connect-SPOService source: $($connectSpoCommand.Source)"
    Write-Debug "Connect-SPOService params: $($connectParamKeys -join ', ')"

    $spoSupportsCertConnect = ($hasAppParam -and $hasThumbprintParam)
    if (-not $spoSupportsCertConnect -and $ps.Major -ge 6) {
        Write-Host -ForegroundColor $Colors.WarningMessage "Loaded SPO module context does not expose certificate auth parameters. Retrying with Windows PowerShell compatibility mode."
        Import-Module microsoft.online.sharepoint.powershell -DisableNameChecking -UseWindowsPowerShell -Force -ErrorAction Stop | Out-Null

        $connectSpoCommand = Get-Command -Name Connect-SPOService -ErrorAction Stop
        $connectParamKeys  = @($connectSpoCommand.Parameters.Keys)
        $hasAppParam       = ($connectParamKeys -contains 'ClientId' -or $connectParamKeys -contains 'AppId')
        $hasThumbprintParam= ($connectParamKeys -contains 'Thumbprint' -or $connectParamKeys -contains 'CertificateThumbprint')
        Write-Debug "Connect-SPOService source after compatibility reload: $($connectSpoCommand.Source)"
        Write-Debug "Connect-SPOService params after compatibility reload: $($connectParamKeys -join ', ')"

        $spoSupportsCertConnect = ($hasAppParam -and $hasThumbprintParam)
    }

    if (-not $spoSupportsCertConnect) {
        $moduleVersion = (Get-Module -Name microsoft.online.sharepoint.powershell | Sort-Object Version -Descending | Select-Object -First 1).Version
        throw "Connect-SPOService certificate auth parameters are unavailable in the loaded SharePoint Online module context (reported version: $moduleVersion). Update microsoft.online.sharepoint.powershell to a version that supports Connect-SPOService certificate authentication."
    }

    if ($GenerateLocalCertificate) {
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

        $certFriendlyName = if (-not [string]::IsNullOrWhiteSpace($provisionTenant)) { "$GeneratedCertSubject - $provisionTenant" } else { $GeneratedCertSubject }
        $exportPfxForProvisioning = $ExportGeneratedPfx -or $ProvisionEntraApp
        $generatedCertificate = New-SpoAdminLocalCertificate -SubjectName $GeneratedCertSubject -YearsValid $GeneratedCertYearsValid -OutputPath $GeneratedCertOutputPath -ExportPfx:$exportPfxForProvisioning -PfxPassword $GeneratedPfxPassword -NoPrompt:$noprompt -FriendlyName $certFriendlyName

        Write-Host -ForegroundColor $Colors.ProcessMessage "Generated certificate thumbprint: $($generatedCertificate.Thumbprint)"
        Write-Host -ForegroundColor $Colors.ProcessMessage "Public certificate exported to: $($generatedCertificate.CerPath)"
        if (-not [string]::IsNullOrWhiteSpace($generatedCertificate.PfxPath)) {
            Write-Host -ForegroundColor $Colors.ProcessMessage "PFX exported to: $($generatedCertificate.PfxPath)"
        }

        if ($ProvisionEntraApp) {
            $resolvedDisplayName = if ([string]::IsNullOrWhiteSpace($AppDisplayName)) { $GeneratedCertSubject } else { $AppDisplayName }

            Write-Host -ForegroundColor $Colors.ProcessMessage "`nStarting Graph authentication for app provisioning..."
            $graphToken = Get-DeviceCodeGraphToken -TenantId $provisionTenant -ClientId $SetupClientId -CopyCodeToClipboard:$CopyDeviceCodeToClipboard -Colors $Colors

            try {
                $provisionResult = Invoke-SpoAdminEntraAppProvisioning `
                    -AccessToken $graphToken `
                    -DisplayName $resolvedDisplayName `
                    -CerFilePath $generatedCertificate.CerPath `
                    -Thumbprint $generatedCertificate.Thumbprint `
                    -SpoResourceAppId $SpoResourceAppId `
                    -GraphResourceAppId $GraphResourceAppId `
                    -Colors $Colors
            }
            finally {
                ## FIX #10: Clear the plaintext Graph token from memory as soon as provisioning completes or fails.
                Remove-Variable -Name graphToken -ErrorAction SilentlyContinue
            }

            $adminUrlToStore = Resolve-SpoAdminUrl -ExplicitUrl $AdminUrl -TenantId $provisionTenant
            if ([string]::IsNullOrWhiteSpace($adminUrlToStore)) {
                Write-Host -ForegroundColor $Colors.WarningMessage "Could not derive SharePoint admin URL from tenant '$provisionTenant'. The adminUrl field in the profile map will be blank. Re-run with -AdminUrl to set it, or edit the JSON file manually."
            }
            else {
                Write-Host -ForegroundColor $Colors.ProcessMessage "Admin URL resolved: $adminUrlToStore"
            }

            $mapPath      = $CertificateMapPath
            $profileEntry = [PSCustomObject]@{
                name                  = $resolvedDisplayName
                tenant                = $provisionTenant
                adminUrl              = $adminUrlToStore
                appId                 = $provisionResult.AppId
                certificateThumbprint = $generatedCertificate.Thumbprint
                pfxPath               = $generatedCertificate.PfxPath  ## already normalised to "" if not exported
            }

            Set-SpoAdminProfileMapEntry -MapPath $mapPath -ProfileEntry $profileEntry -Colors $Colors

            Write-Host -ForegroundColor $Colors.SystemMessage "`n=== Provisioning complete ==="
            Write-Host -ForegroundColor $Colors.ProcessMessage "App ID:           $($provisionResult.AppId)"
            Write-Host -ForegroundColor $Colors.ProcessMessage "Cert Thumbprint:  $($generatedCertificate.Thumbprint)"
            Write-Host -ForegroundColor $Colors.ProcessMessage "Tenant:           $provisionTenant"
            Write-Host -ForegroundColor $Colors.WarningMessage "IMPORTANT: New app role grants can take 15-30 minutes to replicate across services."
            Write-Host -ForegroundColor $Colors.WarningMessage "Certificate-based Connect-SPOService attempts may fail during this window even when provisioning succeeded."
            Write-Host -ForegroundColor $Colors.ProcessMessage "`nConnect any time using:"
            Write-Host -ForegroundColor $Colors.ProcessMessage "  .\o365-connect-spo-cert.ps1 -UseCertificateAuth -Tenant '$provisionTenant'"
        }

        Write-Host -ForegroundColor $Colors.SystemMessage "`nCertificate generation finished`n"
        exit 0
    }

    # --- UseCertificateAuth mode ---
    Write-Debug "Resolving profile and certificate auth inputs."
    $resolvedProfile = Resolve-SpoAdminCertificateProfile -Path $CertificateMapPath -TenantFilter $Tenant -ProfileFilter $ProfileName -AdminUrlFilter $AdminUrl -NoPrompt:$noprompt -Colors $Colors
    if ($null -ne $resolvedProfile) {
        if ([string]::IsNullOrWhiteSpace($Tenant))                 { $Tenant                 = $resolvedProfile.tenant }
        if ([string]::IsNullOrWhiteSpace($AdminUrl))               { $AdminUrl               = $resolvedProfile.adminUrl }
        if ([string]::IsNullOrWhiteSpace($AppId))                  { $AppId                  = $resolvedProfile.appId }
        if ([string]::IsNullOrWhiteSpace($CertificateThumbprint))  { $CertificateThumbprint  = $resolvedProfile.certificateThumbprint }
    }


    ## Derive admin URL from tenant if still not resolved.
    if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
        $AdminUrl = Resolve-SpoAdminUrl -TenantId $Tenant
    }

    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
        $CertificateThumbprint = Resolve-SpoAdminCertificateThumbprint -PreferredSubject $GeneratedCertSubject -Colors $Colors
        if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
            Write-Host -ForegroundColor $Colors.WarningMessage "No local certificate thumbprint was supplied. Looking for a generated certificate in the local store..."
        }
    }

    if ([string]::IsNullOrWhiteSpace($Tenant) -or [string]::IsNullOrWhiteSpace($AdminUrl) -or [string]::IsNullOrWhiteSpace($AppId) -or [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
        $missingFields = @()
        if ([string]::IsNullOrWhiteSpace($Tenant))                { $missingFields += 'Tenant' }
        if ([string]::IsNullOrWhiteSpace($AdminUrl))              { $missingFields += 'AdminUrl' }
        if ([string]::IsNullOrWhiteSpace($AppId))                 { $missingFields += 'AppId' }
        if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $missingFields += 'CertificateThumbprint' }
        throw "UseCertificateAuth missing required value(s): $($missingFields -join ', '). Provide directly or via CertificateMapPath profile."
    }

    ## Verify the certificate is present in the local store before attempting to connect.
    $candidatePfxPaths = @()
    if ($null -ne $resolvedProfile -and -not [string]::IsNullOrWhiteSpace($resolvedProfile.pfxPath)) {
        $candidatePfxPaths += $resolvedProfile.pfxPath
    }
    $localCert = Resolve-SpoAdminCertificate -Thumbprint $CertificateThumbprint -Colors $Colors -CandidatePfxPaths $candidatePfxPaths

    ## FIX #8: Use actual datetime comparison for expiry, not day-count, to handle same-day expiry correctly.
    if ($localCert.NotAfter -lt (Get-Date)) {
        throw "Certificate '$CertificateThumbprint' expired on $($localCert.NotAfter.ToString('yyyy-MM-dd HH:mm:ss')). Provision a new certificate."
    }
    $daysUntilExpiry = ($localCert.NotAfter - (Get-Date)).Days
    if ($daysUntilExpiry -le 30) {
        Write-Host -ForegroundColor $Colors.WarningMessage "Warning: Certificate expires in $daysUntilExpiry day(s) on $($localCert.NotAfter.ToString('yyyy-MM-dd'))."
    }
    else {
        Write-Debug "Certificate valid for $daysUntilExpiry more day(s)."
    }

    ## Probe for an existing SPO connection and disconnect if one is active.
    $existingConnection = $false
    try { Get-SPOTenant -ErrorAction Stop | Out-Null; $existingConnection = $true } catch {}

    if ($existingConnection) {
        Write-Host -ForegroundColor $Colors.WarningMessage "An existing SharePoint Online session is active. Disconnecting before making a new connection."
        Disconnect-SPOService -ErrorAction SilentlyContinue
        Write-Host -ForegroundColor $Colors.ProcessMessage "Previous connection disconnected."
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Connecting to SharePoint admin center with certificate authentication"

    $connectSpoCommand = Get-Command -Name Connect-SPOService -ErrorAction Stop
    $connectParams = @{
        Url         = $AdminUrl
        ErrorAction = 'Stop'
    }

    if ($connectSpoCommand.Parameters.ContainsKey('ClientId')) {
        $connectParams.ClientId = $AppId
    }
    else {
        $connectParams.AppId = $AppId
    }

    if ($connectSpoCommand.Parameters.ContainsKey('Certificate')) {
        ## Prefer passing the in-memory certificate object when supported to avoid store lookup edge-cases.
        $connectParams.Certificate = $localCert
    }
    elseif ($connectSpoCommand.Parameters.ContainsKey('CertificateThumbprint')) {
        $connectParams.CertificateThumbprint = $CertificateThumbprint
    }
    else {
        $connectParams.Thumbprint = $CertificateThumbprint
    }

    if ($connectSpoCommand.Parameters.ContainsKey('Tenant')) {
        $connectParams.Tenant = $Tenant
    }
    elseif ($connectSpoCommand.Parameters.ContainsKey('TenantId')) {
        $connectParams.TenantId = $Tenant
    }
    else {
        Write-Debug "Connect-SPOService does not expose Tenant/TenantId parameter in this module version; attempting connect with URL-scoped tenant only."
    }

    Write-Debug "Connect-SPOService parameter keys: $($connectParams.Keys -join ', ')"
    Connect-SPOService @connectParams

    Write-SpoAdminConnectedCenter -RequestedAdminUrl $AdminUrl -Colors $Colors
    Write-SpoAdminCertConnectionDetails -TenantId $Tenant -AppId $AppId -LocalCert $localCert -Colors $Colors

    Write-Host -ForegroundColor $Colors.ProcessMessage "Connected to SharePoint admin center`n"
    Write-Host -ForegroundColor $Colors.ProcessMessage "SPO cmdlets are available, for example: Get-SPOTenant or Get-SPOSite"
    Write-Host -ForegroundColor $Colors.SystemMessage "SharePoint Online admin center certificate auth flow finished`n"
}
catch {
    Write-Host -ForegroundColor $Colors.ErrorMessage "Script failed: $($_.Exception.Message)"
    if ($UseCertificateAuth -and $_.Exception.Message -match '(?i)certificate with thumbprint|no certificate was found|private key|pfx') {
        Write-Host -ForegroundColor $Colors.WarningMessage "The certificate thumbprint could not be resolved for Connect-SPOService. Confirm the certificate exists in CurrentUser or LocalMachine\\My and contains a private key."
    }
    if ($UseCertificateAuth -and $_.Exception.Message -match '(?i)access denied|forbidden|unauthorized|insufficient privileges|aadsts|permission') {
        Write-Host -ForegroundColor $Colors.WarningMessage "If this app/certificate was just provisioned, wait 15-30 minutes and try again due to RBAC replication lag."
    }
    exit 1
}
finally {
    if ($enableLog) {
        Stop-Transcript | Out-Null
    }
}