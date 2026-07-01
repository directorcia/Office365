[CmdletBinding()]
param(
    [switch]$noprompt = $false,   ## if -noprompt used then user will not be asked for any input
    [switch]$noupdate = $false,   ## if -noupdate used then module will not be checked for more recent version
    [switch]$enableLog = $false,  ## if -enableLog create a transcript log file

    [switch]$GenerateLocalCertificate = $false,
    [switch]$UseCertificateAuth = $false,

    [string]$GeneratedCertSubject = "O365-TMS-AppAuth",
    [int]$GeneratedCertYearsValid = 2,
    [string]$GeneratedCertOutputPath = "",  ## defaults to parent of script directory at runtime
    [switch]$ExportGeneratedPfx = $false,
    [securestring]$GeneratedPfxPassword,

    [string]$Tenant,
    [string]$ProfileName,
    [string]$Organization,
    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$CertificateMapPath = "",  ## defaults to o365-teams-cert-auth.json in parent of script directory

    ## Well-known Microsoft Graph resource app ID used to resolve Graph app roles during provisioning.
    [string]$GraphResourceAppId = "00000003-0000-0000-c000-000000000000",

    ## When used with -GenerateLocalCertificate, also create the Entra app, upload the cert,
    ## and update the profile map automatically.
    [switch]$ProvisionEntraApp = $false,
    [string]$AppDisplayName = "o365-TMS-AppAuth",
    ## Client ID of a public client app registered in your tenant for device-code Graph auth.
    ## The default Azure PowerShell public client is often blocked by tenant preauthorization,
    ## so we fall back to an Azure CLI-style public client that is broadly accepted.
    [string]$SetupClientId = "04b07795-8ddb-461a-bbee-02f9e1bf7bfe",
    ## Microsoft Entra directory role to assign to the service principal for Teams app-based auth.
    [string]$EntraDirectoryRoleName = "Teams Administrator",
    ## Opt-in Graph device-code pre-check in UseCertificateAuth mode to verify/assign SP + role.
    [switch]$ValidateEntraOnConnect = $false,
    ## Skip Graph device-code pre-check in UseCertificateAuth mode and connect directly with cert.
    ## Backward compatibility switch; -ValidateEntraOnConnect is preferred.
    [switch]$SkipEntraValidationOnConnect = $false,
    ## Opt-in only. Clipboard copy can leak auth codes in shared/RDP sessions.
    [switch]$CopyDeviceCodeToClipboard = $false
)
<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.
Documentation - https://github.com/directorcia/Office365/wiki/Certificate-based-connection-for-Teams
Description - Simplified Microsoft Teams connect script with two modes:
1. GenerateLocalCertificate: create/export local cert files.
2. UseCertificateAuth: connect to Microsoft Teams with existing app/cert.
Usage - Example:
  .\o365-connect-tms-cert.ps1 -GenerateLocalCertificate -ProvisionEntraApp -Tenant 'contoso.onmicrosoft.com'
  .\o365-connect-tms-cert.ps1 -UseCertificateAuth -Tenant 'contoso.onmicrosoft.com'
#>

## Resolve paths relative to the script file itself, not the caller's working directory.
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptParentDir = Split-Path -Parent $scriptDir
if ([string]::IsNullOrWhiteSpace($GeneratedCertOutputPath)) { $GeneratedCertOutputPath = $scriptParentDir }

if ([string]::IsNullOrWhiteSpace($CertificateMapPath)) {
    ## Parent directory is searched first so reads are consistent with where writes land.
    $candidateCertificateMapPaths = @(
        (Join-Path $scriptParentDir 'o365-tms-cert-auth.json'),
        (Join-Path $scriptDir 'o365-tms-cert-auth.json')
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($candidatePath in $candidateCertificateMapPaths) {
        if (Test-Path -LiteralPath $candidatePath) {
            $CertificateMapPath = $candidatePath
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($CertificateMapPath)) {
        ## Default write location: parent directory, matching GeneratedCertOutputPath convention.
        $CertificateMapPath = Join-Path $scriptParentDir 'o365-tms-cert-auth.json'
    }
}

## Shared output colors passed explicitly to functions to avoid hidden script-scope coupling.
$Colors = @{
    SystemMessage  = "cyan"
    ProcessMessage = "green"
    ErrorMessage   = "red"
    WarningMessage = "yellow"
}

## Tracks whether this run opened a Teams session that should be closed on error.
$disconnectTeamsAuthOnError = $false

## Resolve the executable of the current host so elevated module installs target the same runtime.
$elevatedShellPath = (Get-Process -Id $PID).MainModule.FileName
if ([string]::IsNullOrWhiteSpace($elevatedShellPath) -or -not (Test-Path -LiteralPath $elevatedShellPath)) {
    throw "Unable to resolve current PowerShell host executable path for elevated module operations (PID $PID)."
}

function Invoke-ElevatedPowerShellCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CommandText,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ShellPath
    )

    # Use -EncodedCommand so quoting and special characters survive UAC elevation.
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($CommandText))
    $startParams = @{
        FilePath     = $ShellPath
        Verb         = 'RunAs'
        Wait         = $true
        WindowStyle  = 'Hidden'
        ArgumentList = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $encodedCommand)
    }
    $process = Start-Process @startParams -PassThru
    if ($null -eq $process -or $process.ExitCode -ne 0) {
        $exitCode = if ($null -eq $process) { "(unknown)" } else { $process.ExitCode }
        throw "Elevated command failed with exit code $exitCode. Command: $CommandText"
    }
}

function Resolve-TeamsCertificateProfile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Path,
        [Parameter(Mandatory = $false)]
        [string]$TenantFilter,
        [Parameter(Mandatory = $false)]
        [string]$ProfileFilter,
        [Parameter(Mandatory = $false)]
        [string]$OrganizationFilter,
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

    $candidateProfiles = $profileItems
    if (-not [string]::IsNullOrWhiteSpace($ProfileFilter)) {
        $candidateProfiles = @($candidateProfiles | Where-Object { $_.name -eq $ProfileFilter })
    }
    if (-not [string]::IsNullOrWhiteSpace($TenantFilter)) {
        $candidateProfiles = @($candidateProfiles | Where-Object { $_.tenant -eq $TenantFilter -or $_.organization -eq $TenantFilter })
    }
    if (-not [string]::IsNullOrWhiteSpace($OrganizationFilter)) {
        $candidateProfiles = @($candidateProfiles | Where-Object { $_.organization -eq $OrganizationFilter })
    }

    if ($candidateProfiles.Count -eq 0) {
        throw "No matching certificate profile found in '$Path'."
    }

    if ($candidateProfiles.Count -eq 1 -or $NoPrompt) {
        if ($candidateProfiles.Count -gt 1 -and $NoPrompt) {
            throw "Multiple matching profiles found in '$Path'. Specify -ProfileName, -Tenant, or -Organization."
        }
        return $candidateProfiles[0]
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Multiple matching certificate profiles found:"
    for ($index = 0; $index -lt $candidateProfiles.Count; $index++) {
        $displayName = if ([string]::IsNullOrWhiteSpace($candidateProfiles[$index].name)) { "(unnamed)" } else { $candidateProfiles[$index].name }
        Write-Host -ForegroundColor $Colors.ProcessMessage ("[{0}] {1} | Tenant={2} | Org={3} | AppId={4}" -f ($index + 1), $displayName, $candidateProfiles[$index].tenant, $candidateProfiles[$index].organization, $candidateProfiles[$index].appId)
    }

    do {
        $choice = Read-Host -Prompt "Select profile number"
        [int]$parsedChoice = 0
        $validSelection = [int]::TryParse($choice, [ref]$parsedChoice) -and $parsedChoice -ge 1 -and $parsedChoice -le $candidateProfiles.Count
    } until ($validSelection)

    return $candidateProfiles[$parsedChoice - 1]
}

function New-TeamsLocalCertificate {
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
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    $certificate = New-SelfSignedCertificate -Subject "CN=$SubjectName" -CertStoreLocation "Cert:\CurrentUser\My" -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears($YearsValid) -KeySpec Signature -ErrorAction Stop

    $resolvedFriendlyName = if ([string]::IsNullOrWhiteSpace($FriendlyName)) { $SubjectName } else { $FriendlyName }
    $certificate.FriendlyName = $resolvedFriendlyName

    $safeSubject = ($SubjectName -replace '[^A-Za-z0-9\-_.]', '-')
    $fileBase = "{0}-{1}" -f $safeSubject, $certificate.Thumbprint
    $cerPath = Join-Path -Path $OutputPath -ChildPath "$fileBase.cer"

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
        Subject = $certificate.Subject
        NotAfter = $certificate.NotAfter
        CerPath = $cerPath
        PfxPath = $pfxPath
    }
}

function Get-DeviceCodeGraphToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TenantId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ClientId,
        [Parameter(Mandatory = $false)][string]$Scope = "offline_access openid profile https://graph.microsoft.com/Application.ReadWrite.All https://graph.microsoft.com/AppRoleAssignment.ReadWrite.All https://graph.microsoft.com/RoleManagement.ReadWrite.Directory",
        [Parameter(Mandatory = $false)][switch]$CopyCodeToClipboard,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $normalizedTenantId = ($TenantId ?? '').Trim().Trim('/')
    if ([string]::IsNullOrWhiteSpace($normalizedTenantId)) {
        throw "TenantId cannot be empty for device-code Graph auth."
    }

    if ($normalizedTenantId -match '^https?://login\.microsoftonline\.com/([^/]+)/?.*$') {
        $normalizedTenantId = $Matches[1]
    }

    $clientCandidates = [System.Collections.Generic.List[string]]::new()
    $clientCandidates.Add($ClientId)
    foreach ($fallbackClientId in @(
            '14d82eec-204b-4c2f-b7e8-296a70dab67e', # Microsoft Graph PowerShell public client
            '1950a258-227b-4e31-a9cf-717495945fc2'  # Azure PowerShell legacy public client
        )) {
        if ($clientCandidates -notcontains $fallbackClientId) {
            $clientCandidates.Add($fallbackClientId)
        }
    }

    $deviceCodeResponse = $null
    $deviceCodeRequestErrors = [System.Collections.Generic.List[string]]::new()

    foreach ($candidateClientId in $clientCandidates) {
        Write-Debug "Requesting device code for tenant: $normalizedTenantId, client: $candidateClientId"
        try {
            $deviceCodeResponse = Invoke-RestMethod -Method Post `
                -Uri "https://login.microsoftonline.com/$normalizedTenantId/oauth2/v2.0/devicecode" `
                -Body @{ client_id = $candidateClientId; scope = $Scope } `
                -ContentType "application/x-www-form-urlencoded" `
                -ErrorAction Stop

            # If a fallback public client succeeds, continue token polling with that client ID.
            $ClientId = $candidateClientId
            if ($candidateClientId -ne $clientCandidates[0]) {
                Write-Host -ForegroundColor $Colors.WarningMessage "Primary setup client was rejected. Using fallback public client ID for this run: $candidateClientId"
            }
            break
        }
        catch {
            $errorCode = $null
            $errorDescription = $null
            try {
                $errorContent = ($_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction Stop)
                $errorCode = $errorContent.error
                $errorDescription = $errorContent.error_description
            }
            catch {
                $errorDescription = $_.Exception.Message
            }

            $flatError = if ([string]::IsNullOrWhiteSpace($errorCode)) {
                "client=${candidateClientId}: $errorDescription"
            }
            else {
                "client=${candidateClientId}: $errorCode - $errorDescription"
            }
            $deviceCodeRequestErrors.Add($flatError)

            if ($errorCode -eq 'invalid_scope') {
                throw "Device code request failed with invalid_scope. The public client app used for setup is not allowed to request the required Graph delegated scopes in this tenant. Create or use a tenant-approved public client app registration and pass it with -SetupClientId. Tenant='$normalizedTenantId'. Details: $errorDescription"
            }
        }
    }

    if ($null -eq $deviceCodeResponse) {
        $joinedErrors = ($deviceCodeRequestErrors -join ' | ')
        throw "Device code request failed for tenant '$normalizedTenantId'. Tried client IDs: $($clientCandidates -join ', '). Check -SetupClientId and -Tenant. Details: $joinedErrors"
    }

    if ($CopyCodeToClipboard) {
        if (Get-Command Set-Clipboard -ErrorAction SilentlyContinue) {
            Set-Clipboard -Value $deviceCodeResponse.user_code
        }
        else {
            Write-Host -ForegroundColor $Colors.WarningMessage "Set-Clipboard is unavailable in this host. Displaying the code instead."
            $CopyCodeToClipboard = $false
        }
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
    $pollInterval = [int]$deviceCodeResponse.interval

    :pollLoop while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $pollInterval
        try {
            $tokenResponse = Invoke-RestMethod -Method Post `
                -Uri "https://login.microsoftonline.com/$normalizedTenantId/oauth2/v2.0/token" `
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
                    "authorization_pending" { Write-Debug "Waiting for user to complete sign-in..."; continue pollLoop }
                    "slow_down" { $pollInterval += 5; continue pollLoop }
                    "authorization_declined" { throw "User declined the authorization request." }
                    "expired_token" { throw "Device code expired before authorization was completed." }
                    default { throw "Token exchange failed ($($errorContent.error)): $($errorContent.error_description)" }
                }
            }
            throw "Token exchange failed: $($_.Exception.Message)"
        }
    }
    throw "Device code authorization timed out."
}

function Invoke-TeamsGraphRequest {
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
    $params = @{ Method = $Method; Uri = $Uri; Headers = $headers; ErrorAction = "Stop" }
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
            $response = $null
            try { $response = $_.Exception.Response } catch {}
            if ($null -ne $response) {
                try { $statusCode = [int]$response.StatusCode } catch {}
            }
            $shouldRetry = ($attempt -le $MaxRetries) -and ($statusCode -in @(429, 500, 502, 503, 504) -or $msg -match '(?i)throttl|rate limit|temporar|timeout|try again')
            if (-not $shouldRetry) {
                throw "Graph call failed [$Method $Uri]: $msg"
            }
            $delaySeconds = [int][math]::Min(30, [math]::Pow(2, ($attempt - 1)) * $InitialRetryDelaySeconds)
            Write-Debug "Graph call retry $attempt/$MaxRetries after ${delaySeconds}s for [$Method $Uri]. Status=$statusCode"
            Start-Sleep -Seconds $delaySeconds
        }
    }
}

function Get-OrCreateTeamsEntraServicePrincipal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AppId,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $graphBase = "https://graph.microsoft.com/v1.0"

    Write-Host -ForegroundColor $Colors.ProcessMessage "Ensuring Teams service principal exists..."
    $spFilter = [uri]::EscapeDataString("appId eq '$AppId'")
    $existingSp = Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals?`$filter=$spFilter"

    if ($existingSp.value.Count -gt 0) {
        $spObject = $existingSp.value[0]
        Write-Host -ForegroundColor $Colors.ProcessMessage "Teams service principal already exists. Object ID: $($spObject.id)"
        return $spObject
    }

    $spObject = Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/servicePrincipals" -Body @{ appId = $AppId }
    Write-Host -ForegroundColor $Colors.ProcessMessage "Teams service principal created. Object ID: $($spObject.id)"
    return $spObject
}

function Get-TeamsProvisioningRoleTargets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$GraphResourceAppId,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $graphBase = "https://graph.microsoft.com/v1.0"

    Write-Host -ForegroundColor $Colors.ProcessMessage "Locating Microsoft Graph service principal in directory..."
    $graphSpFilter = [uri]::EscapeDataString("appId eq '$GraphResourceAppId'")
    $graphSpResult = Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals?`$filter=$graphSpFilter"
    if ($null -eq $graphSpResult.value -or $graphSpResult.value.Count -eq 0) {
        throw "Microsoft Graph service principal not found in this tenant."
    }

    $graphSp = $graphSpResult.value[0]

    $requiredGraphRoleValues = @(
        'Application.Read.All',
        'Organization.Read.All',
        'User.Read.All',
        'Group.Read.All',
        'Directory.Read.All',
        'Group.ReadWrite.All',
        'Directory.ReadWrite.All',
        'Team.ReadBasic.All',
        'TeamSettings.Read.All',
        'TeamSettings.ReadWrite.All',
        'AppCatalog.ReadWrite.All',
        'Channel.Delete.All',
        'ChannelSettings.ReadWrite.All',
        'ChannelMember.ReadWrite.All'
    )

    $requiredResourceAccess = @(
        @{
            resourceAppId  = $GraphResourceAppId
            resourceAccess = @()
        }
    )

    $graphRoleIds = New-Object System.Collections.Generic.List[object]
    foreach ($roleValue in $requiredGraphRoleValues) {
        $graphRole = $graphSp.appRoles | Where-Object { $_.value -eq $roleValue -and $_.allowedMemberTypes -contains 'Application' } | Select-Object -First 1
        if ($null -eq $graphRole) {
            throw "Microsoft Graph app role '$roleValue' was not found on the Microsoft Graph service principal."
        }

        $requiredResourceAccess[0].resourceAccess += @{ id = $graphRole.id; type = 'Role' }
        $graphRoleIds.Add($graphRole.id)
    }

    return [PSCustomObject]@{
        GraphServicePrincipal    = $graphSp
        GraphRoleIds             = $graphRoleIds.ToArray()
        RequiredResourceAccess   = $requiredResourceAccess
        RequiredGraphRoleValues   = $requiredGraphRoleValues
    }
}

function Set-TeamsEntraRequiredResourceAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AppObjectId,
        [Parameter(Mandatory = $true)][object[]]$RequiredResourceAccess,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $graphBase = "https://graph.microsoft.com/v1.0"
    $existingApp = Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/applications/${AppObjectId}?`$select=requiredResourceAccess"
    $existingRequiredAccess = @($existingApp.requiredResourceAccess)
    $missingRequiredAccess = @()

    foreach ($requiredEntry in @($RequiredResourceAccess)) {
        $existingEntry = $existingRequiredAccess | Where-Object { $_.resourceAppId -eq $requiredEntry.resourceAppId } | Select-Object -First 1
        if ($null -eq $existingEntry) {
            $missingRequiredAccess += $requiredEntry
            continue
        }

        $requiredRoleIds = @($requiredEntry.resourceAccess | ForEach-Object { $_.id })
        $existingRoleIds = @($existingEntry.resourceAccess | ForEach-Object { $_.id })
        foreach ($requiredRoleId in $requiredRoleIds) {
            if ($existingRoleIds -notcontains $requiredRoleId) {
                $missingRequiredAccess += @{
                    resourceAppId  = $requiredEntry.resourceAppId
                    resourceAccess = @(
                        @{ id = $requiredRoleId; type = 'Role' }
                    )
                }
            }
        }
    }

    if ($missingRequiredAccess.Count -gt 0) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Updating app registration required resource access entries..."
        $patchBody = @{ requiredResourceAccess = @($existingRequiredAccess) + $missingRequiredAccess }
        Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Patch -Uri "$graphBase/applications/$AppObjectId" -Body $patchBody | Out-Null
    }
    else {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Required resource access entries already present - skipping update."
    }
}

function Set-TeamsEntraAppRoleAssignments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ServicePrincipalId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$GraphServicePrincipalId,
        [Parameter(Mandatory = $true)][string[]]$GraphRoleIds,
        [Parameter(Mandatory = $true)][string[]]$GraphRoleValues,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $graphBase = "https://graph.microsoft.com/v1.0"
    Write-Host -ForegroundColor $Colors.ProcessMessage "Granting Microsoft Graph permissions (admin consent)..."

    $existingAssignments = Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/servicePrincipals/$ServicePrincipalId/appRoleAssignments"

    for ($i = 0; $i -lt $GraphRoleIds.Count; $i++) {
        $roleId = $GraphRoleIds[$i]
        $roleValue = $GraphRoleValues[$i]
        $alreadyGranted = $existingAssignments.value | Where-Object { $_.appRoleId -eq $roleId -and $_.resourceId -eq $GraphServicePrincipalId } | Select-Object -First 1
        if ($null -ne $alreadyGranted) {
            Write-Host -ForegroundColor $Colors.ProcessMessage "$roleValue already granted - skipping."
            continue
        }

        $roleBody = @{ principalId = $ServicePrincipalId; resourceId = $GraphServicePrincipalId; appRoleId = $roleId }
        Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/servicePrincipals/$ServicePrincipalId/appRoleAssignments" -Body $roleBody | Out-Null
        Write-Host -ForegroundColor $Colors.ProcessMessage "$roleValue granted."
    }
}

function Set-TeamsEntraDirectoryRoleAssignment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ServicePrincipalObjectId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DirectoryRoleName,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $graphBase = "https://graph.microsoft.com/v1.0"

    Write-Host -ForegroundColor $Colors.ProcessMessage "Assigning Microsoft Entra role '$DirectoryRoleName' to the Teams service principal..."
    $roleFilter = [uri]::EscapeDataString("displayName eq '$DirectoryRoleName'")
    $roleResult = Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/directoryRoles?`$filter=$roleFilter"

    if ($null -eq $roleResult.value -or $roleResult.value.Count -eq 0) {
        Write-Host -ForegroundColor $Colors.WarningMessage "Role '$DirectoryRoleName' is not currently activated in this tenant. Attempting to activate it from role templates..."

        # directoryRoleTemplates does not support server-side $filter in Graph v1.0.
        $roleTemplateResult = Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/directoryRoleTemplates"
        $matchingTemplates = @($roleTemplateResult.value | Where-Object { $_.displayName -eq $DirectoryRoleName })
        if ($null -eq $matchingTemplates -or $matchingTemplates.Count -eq 0) {
            throw "Microsoft Entra role template '$DirectoryRoleName' was not found. Choose a different role name with -EntraDirectoryRoleName."
        }

        $roleTemplate = $matchingTemplates[0]
        Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/directoryRoles" -Body @{ roleTemplateId = $roleTemplate.id } | Out-Null
        Write-Host -ForegroundColor $Colors.ProcessMessage "Activated Microsoft Entra role '$DirectoryRoleName' from template."

        $roleResult = Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/directoryRoles?`$filter=$roleFilter"
        if ($null -eq $roleResult.value -or $roleResult.value.Count -eq 0) {
            throw "Microsoft Entra directory role '$DirectoryRoleName' could not be resolved after activation. Verify role availability in this tenant and retry."
        }
    }

    $directoryRole = $roleResult.value[0]
    $membersResult = Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/directoryRoles/$($directoryRole.id)/members?`$select=id"
    $existingMember = $membersResult.value | Where-Object { $_.id -eq $ServicePrincipalObjectId } | Select-Object -First 1

    if ($null -ne $existingMember) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "The Teams service principal is already a member of role '$DirectoryRoleName' - skipping."
        return
    }

    $membershipBody = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$ServicePrincipalObjectId" }
    try {
        Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/directoryRoles/$($directoryRole.id)/members/`$ref" -Body $membershipBody | Out-Null
        Write-Host -ForegroundColor $Colors.ProcessMessage "Assigned Microsoft Entra role '$DirectoryRoleName' to the Teams service principal."
    }
    catch {
        # Graph may return a duplicate-reference error if membership already exists but was not visible in the previous members query page.
        if ($_.Exception.Message -match '(?i)already exist.*members|added object references already exist') {
            Write-Host -ForegroundColor $Colors.ProcessMessage "The Teams service principal is already a member of role '$DirectoryRoleName' - continuing."
            return
        }
        throw
    }
}

function Set-TeamsProfileMapEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$MapPath,
        [Parameter(Mandatory = $true)][object]$ProfileEntry,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $fullMapPath = [System.IO.Path]::GetFullPath($MapPath)
    $mapDirectory = Split-Path -Parent $fullMapPath
    if (-not [string]::IsNullOrWhiteSpace($mapDirectory) -and -not (Test-Path -LiteralPath $mapDirectory)) {
        New-Item -Path $mapDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    $hashBytes = [System.Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::UTF8.GetBytes($fullMapPath.ToLowerInvariant()))
    $hashHex = ([System.BitConverter]::ToString($hashBytes)).Replace('-', '')
    $mutexName = "Global\CIAOPS_TEAMS_PROFILEMAP_$hashHex"

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

        $existingIdx = $null
        for ($i = 0; $i -lt $profileList.Count; $i++) {
            $sameApp = (-not [string]::IsNullOrWhiteSpace($profileList[$i].appId) -and $profileList[$i].appId -eq $ProfileEntry.appId)
            $sameTenant = ((-not [string]::IsNullOrWhiteSpace($profileList[$i].tenant) -and $profileList[$i].tenant -eq $ProfileEntry.tenant) -or
                (-not [string]::IsNullOrWhiteSpace($profileList[$i].organization) -and $profileList[$i].organization -eq $ProfileEntry.organization))
            if ($sameApp -or $sameTenant) { $existingIdx = $i; break }
        }
        if ($null -ne $existingIdx) { $profileList[$existingIdx] = $ProfileEntry } else { $profileList.Add($ProfileEntry) }

        @{ profiles = $profileList.ToArray() } | ConvertTo-Json -Depth 5 | Set-Content -Path $tempPath -Encoding UTF8 -ErrorAction Stop
        Move-Item -Path $tempPath -Destination $fullMapPath -Force -ErrorAction Stop
        Write-Host -ForegroundColor $Colors.ProcessMessage "Profile map updated: $fullMapPath"
    }
    finally {
        if (Test-Path -Path $tempPath) { Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue }
        if ($hasHandle) { $mutex.ReleaseMutex() }
        $mutex.Dispose()
    }
}

function Get-OrCreateTeamsEntraApplication {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DisplayName,
        [Parameter(Mandatory = $true)][ValidatePattern('^[0-9A-Fa-f]{40}$')][string]$Thumbprint,
        [Parameter(Mandatory = $true)][hashtable]$Colors
    )

    $graphBase = "https://graph.microsoft.com/v1.0"
    $createdNewApp = $false

    Write-Host -ForegroundColor $Colors.ProcessMessage "Checking for existing app registration: $DisplayName..."
    $appFilter = [uri]::EscapeDataString("displayName eq '" + ($DisplayName -replace "'", "''") + "'")
    $existingApps = Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/applications?`$filter=$appFilter"

    if ($existingApps.value.Count -gt 0) {
        Write-Host -ForegroundColor $Colors.WarningMessage "App '$DisplayName' already exists - reusing existing registration."
        $appObject = $existingApps.value[0]
    }

    if ($null -eq $appObject) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Creating app registration: $DisplayName..."
        $newAppBody = @{ displayName = $DisplayName; signInAudience = "AzureADMyOrg" }
        $appObject = Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Post -Uri "$graphBase/applications" -Body $newAppBody
        $createdNewApp = $true
    }

    $certStoreObj = Get-Item "Cert:\CurrentUser\My\$Thumbprint" -ErrorAction Stop
    $cerBytes = $certStoreObj.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    $cerBase64 = [System.Convert]::ToBase64String($cerBytes)
    $customKeyIdBase64 = [System.Convert]::ToBase64String($certStoreObj.GetCertHash())

    $existingApp = Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Get -Uri "$graphBase/applications/$($appObject.id)?`$select=id,appId,displayName,keyCredentials"
    $existingKeys = @($existingApp.keyCredentials | Where-Object { $_.customKeyIdentifier -ne $customKeyIdBase64 })
    $hasMatchingKey = ($existingKeys.Count -ne @($existingApp.keyCredentials).Count)

    if (-not $hasMatchingKey) {
        $newKey = @{ type = "AsymmetricX509Cert"; usage = "Verify"; key = $cerBase64; displayName = "$($env:COMPUTERNAME) - Teams-Auth-$Thumbprint"; customKeyIdentifier = $customKeyIdBase64 }
        $certPatch = @{ keyCredentials = @($existingKeys) + @($newKey) }
        Invoke-TeamsGraphRequest -AccessToken $AccessToken -Method Patch -Uri "$graphBase/applications/$($appObject.id)" -Body $certPatch | Out-Null
        Write-Host -ForegroundColor $Colors.ProcessMessage "Certificate uploaded to app registration for thumbprint: $Thumbprint"
    }
    else {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Certificate thumbprint already present in app registration - no key update needed."
    }

    if ($createdNewApp) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "App created and certificate processed. Object ID: $($existingApp.id) | App ID: $($existingApp.appId)"
    }
    else {
        Write-Host -ForegroundColor $Colors.ProcessMessage "App reused and certificate processed. Object ID: $($existingApp.id) | App ID: $($existingApp.appId)"
    }
    return $existingApp
}

function Write-TeamsConnectedTenant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$RequestedOrganization,
        [Parameter(Mandatory = $true)]
        [hashtable]$Colors
    )

    $connectedTenant = $null
    try {
        if (Get-Command Get-CsTenant -ErrorAction SilentlyContinue) {
            $tenant = Get-CsTenant -ErrorAction Stop | Select-Object -First 1
            if ($null -ne $tenant) {
                $connectedTenant = $tenant.TenantId
                if ([string]::IsNullOrWhiteSpace($connectedTenant) -and -not [string]::IsNullOrWhiteSpace($tenant.Tenant)) {
                    $connectedTenant = $tenant.Tenant
                }
                if ([string]::IsNullOrWhiteSpace($connectedTenant) -and -not [string]::IsNullOrWhiteSpace($tenant.Id)) {
                    $connectedTenant = $tenant.Id
                }
            }
        }
    }
    catch {
        Write-Debug "Get-CsTenant unavailable or returned no data."
    }

    if ([string]::IsNullOrWhiteSpace($connectedTenant)) {
        try {
            if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
                $graphContext = Get-MgContext -ErrorAction Stop
                if ($null -ne $graphContext) {
                    $connectedTenant = $graphContext.TenantId
                    if ([string]::IsNullOrWhiteSpace($connectedTenant) -and $graphContext.PSObject.Properties.Name -contains 'Tenant') {
                        $connectedTenant = $graphContext.Tenant
                    }
                }
            }
        }
        catch {
            Write-Debug "Get-MgContext unavailable or returned no data."
        }
    }

    if ([string]::IsNullOrWhiteSpace($connectedTenant) -and -not [string]::IsNullOrWhiteSpace($RequestedOrganization)) {
        $connectedTenant = $RequestedOrganization
    }

    if ([string]::IsNullOrWhiteSpace($connectedTenant)) {
        $connectedTenant = "(unable to determine from current session)"
    }

    Write-Host -ForegroundColor $Colors.SystemMessage "Connected Teams tenant: $connectedTenant"
    if (-not [string]::IsNullOrWhiteSpace($RequestedOrganization) -and $connectedTenant -ne "(unable to determine from current session)") {
        $requestedNormalized = $RequestedOrganization.Trim().ToLowerInvariant()
        $connectedNormalized = $connectedTenant.Trim().ToLowerInvariant()

        $resolvedRequestedTenantId = Resolve-TeamsTenantId -TenantOrDomain $RequestedOrganization
        $resolvedConnectedTenantId = Resolve-TeamsTenantId -TenantOrDomain $connectedTenant

        $tenantMatches = ($requestedNormalized -eq $connectedNormalized) -or
            (-not [string]::IsNullOrWhiteSpace($resolvedRequestedTenantId) -and $resolvedRequestedTenantId -eq $connectedNormalized) -or
            (-not [string]::IsNullOrWhiteSpace($resolvedConnectedTenantId) -and $requestedNormalized -eq $resolvedConnectedTenantId) -or
            (-not [string]::IsNullOrWhiteSpace($resolvedRequestedTenantId) -and -not [string]::IsNullOrWhiteSpace($resolvedConnectedTenantId) -and $resolvedRequestedTenantId -eq $resolvedConnectedTenantId)

        if (-not $tenantMatches) {
            Write-Host -ForegroundColor $Colors.WarningMessage "Requested organization was '$RequestedOrganization' but active session reports '$connectedTenant'."
        }
    }
}

function Resolve-TeamsTenantId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$TenantOrDomain
    )

    $inputValue = ($TenantOrDomain ?? '').Trim().Trim('/').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($inputValue)) {
        return $null
    }

    if ($inputValue -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') {
        return $inputValue
    }

    if ($inputValue -match '^https?://login\.microsoftonline\.com/([^/]+)/?.*$') {
        $inputValue = $Matches[1].ToLowerInvariant()
        if ($inputValue -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') {
            return $inputValue
        }
    }

    try {
        $openidConfig = Invoke-RestMethod -Method Get -Uri "https://login.microsoftonline.com/$inputValue/v2.0/.well-known/openid-configuration" -ErrorAction Stop
        if ($null -ne $openidConfig -and -not [string]::IsNullOrWhiteSpace($openidConfig.issuer) -and $openidConfig.issuer -match '^https://login\.microsoftonline\.com/([0-9a-fA-F\-]{36})/v2\.0$') {
            return $Matches[1].ToLowerInvariant()
        }
    }
    catch {
        Write-Debug "Unable to resolve tenant identifier '$TenantOrDomain' via OpenID metadata."
    }

    return $null
}

function Get-TeamsCertClientAssertionToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TenantId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AppId,
        [Parameter(Mandatory = $true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory = $false)][string]$Scope = "https://graph.microsoft.com/.default"
    )

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

    $headerJson = '{"alg":"RS256","typ":"JWT","x5t":"' + ([System.Convert]::ToBase64String($Certificate.GetCertHash())).TrimEnd('=') + '"}'
    $payloadJson = '{"aud":"' + $tokenEndpoint + '","iss":"' + $AppId + '","sub":"' + $AppId + '","jti":"' + [System.Guid]::NewGuid().ToString() + '","nbf":' + $now + ',"exp":' + ($now + 600) + '}'

    $headerB64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($headerJson)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $payloadB64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payloadJson)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $signingInput = [System.Text.Encoding]::UTF8.GetBytes("$headerB64.$payloadB64")

    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    $sigBytes = $rsa.SignData($signingInput, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $sigB64 = [System.Convert]::ToBase64String($sigBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    $clientAssertion = "$headerB64.$payloadB64.$sigB64"

    try {
        $response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ErrorAction Stop `
            -ContentType "application/x-www-form-urlencoded" `
            -Body @{
                grant_type            = "client_credentials"
                client_id             = $AppId
                scope                 = $Scope
                client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
                client_assertion      = $clientAssertion
            }
        return $response.access_token
    }
    catch {
        $detail = $null
        try { $detail = ($_.ErrorDetails.Message | ConvertFrom-Json).error_description } catch {}
        throw "Client assertion token request failed: $(if ($detail) { $detail } else { $_.Exception.Message })"
    }
}

function Write-TeamsCertConnectionDetails {
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

    try {
        $graphToken = Get-TeamsCertClientAssertionToken -TenantId $TenantId -AppId $AppId -Certificate $LocalCert
        $graphBase = "https://graph.microsoft.com/v1.0"
        $appFilter = [uri]::EscapeDataString("appId eq '" + ($AppId -replace "'", "''") + "'")
        $appResult = Invoke-TeamsGraphRequest -AccessToken $graphToken -Method Get -Uri "$graphBase/applications?`$filter=$appFilter&`$select=displayName,keyCredentials"
        $appObj = $appResult.value | Select-Object -First 1

        if ($null -ne $appObj) {
            Write-Host -ForegroundColor $Colors.ProcessMessage "`n$sep"
            Write-Host -ForegroundColor $Colors.ProcessMessage "  ENTRA ID APP REGISTRATION"
            Write-Host -ForegroundColor $Colors.ProcessMessage $sep
            Write-Host -ForegroundColor $Colors.ProcessMessage ("  Display Name  : {0}" -f $appObj.displayName)
            Write-Host -ForegroundColor $Colors.ProcessMessage ("  App ID        : {0}" -f $AppId)

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

            $otherKeys = @($appObj.keyCredentials | Where-Object { $_.customKeyIdentifier -ne $thumbBase64 })
            if ($otherKeys.Count -gt 0) {
                Write-Host -ForegroundColor $Colors.ProcessMessage ("`n  Other registered certs on this app ({0}):" -f $otherKeys.Count)
                foreach ($key in $otherKeys) {
                    Write-Host -ForegroundColor $Colors.ProcessMessage ("    - {0}  [{1} -> {2}]" -f $key.displayName, ([datetime]$key.startDateTime).ToString('yyyy-MM-dd'), ([datetime]$key.endDateTime).ToString('yyyy-MM-dd'))
                }
            }
        }
    }
    catch {
        $entraDetailError = $_.Exception.Message
        if ($entraDetailError -match "Insufficient privileges") {
            Write-Host -ForegroundColor $Colors.WarningMessage "`n  Entra ID cert details skipped: this app lacks Graph read permission for applications."
            Write-Host -ForegroundColor $Colors.WarningMessage "  Local cert details above are valid; Teams connection is unaffected."
            Write-Host -ForegroundColor $Colors.WarningMessage "  To enable Entra matching details, grant Microsoft Graph Application.Read.All (Application) and admin consent."
        }
        elseif ($entraDetailError -match '(?i)AADSTS700027|key was not found|certificate with identifier used to sign the client assertion is not registered') {
            Write-Host -ForegroundColor $Colors.WarningMessage "`n  Entra ID cert details unavailable: current local certificate is not registered on this app ID."
            Write-Host -ForegroundColor $Colors.WarningMessage "  Re-run provisioning to upload the active cert to the app, then reconnect."
        }
        else {
            Write-Host -ForegroundColor $Colors.WarningMessage "`n  (Entra ID cert details unavailable: $entraDetailError)"
        }
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "$sep`n"
}

Clear-Host

## Enforce TLS 1.2 minimum.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ($enableLog) {
    $logPath = Join-Path $scriptParentDir 'o365-connect-teams.txt'
    Write-Host "Script activity logged at $logPath"
    Start-Transcript $logPath | Out-Null
}

try {
    Write-Host -ForegroundColor $Colors.SystemMessage "Microsoft Teams Connection script started`n"
    Write-Host -ForegroundColor $Colors.ProcessMessage "Prompt =", (-not $noprompt)

    if (($GenerateLocalCertificate -and $UseCertificateAuth) -or (-not $GenerateLocalCertificate -and -not $UseCertificateAuth)) {
        throw "Specify exactly one mode: -GenerateLocalCertificate or -UseCertificateAuth."
    }

    if (Get-Module -ListAvailable -Name MicrosoftTeams) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Microsoft Teams PowerShell module installed"
    }
    else {
        Write-Host -ForegroundColor $Colors.WarningMessage -BackgroundColor $Colors.ErrorMessage "[001] - Microsoft Teams PowerShell module not installed`n"
        if (-not $noprompt) {
            do {
                $response = Read-Host -Prompt "`nDo you wish to install the Microsoft Teams PowerShell module (Y/N)?"
            } until (-not [string]::IsNullOrWhiteSpace($response))
            if ($response -ne 'Y' -and $response -ne 'y') {
                throw "Microsoft Teams module is required."
            }
        }

        Write-Host -ForegroundColor $Colors.ProcessMessage "Installing Microsoft Teams PowerShell module - Administration escalation required"
        Invoke-ElevatedPowerShellCommand -ShellPath $elevatedShellPath -CommandText "Install-Module -Name MicrosoftTeams -Scope AllUsers -Force -Confirm:`$false -ErrorAction Stop"
        Write-Host -ForegroundColor $Colors.ProcessMessage "Microsoft Teams PowerShell module installed"
    }

    if (-not $noupdate) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Checking whether newer version of Microsoft Teams module is available"
        $version = Get-InstalledModule -Name MicrosoftTeams -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        $psgalleryVersion = Find-Module -Name MicrosoftTeams -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1

        $localVersion = if ($null -ne $version) { $version.Version -as [string] } else { $null }
        $onlineVersion = if ($null -ne $psgalleryVersion) { $psgalleryVersion.Version -as [string] } else { $null }

        if ($null -eq $localVersion -or $null -eq $onlineVersion) {
            Write-Host -ForegroundColor $Colors.WarningMessage "Unable to compare module versions - skipping update check."
        }
        elseif ([version]$localVersion -lt [version]$onlineVersion) {
            Write-Host -ForegroundColor $Colors.WarningMessage "Local module $localVersion is lower than Gallery module $onlineVersion"
            if (-not $noprompt) {
                do {
                    $updateResponse = Read-Host -Prompt "`nDo you wish to update the Microsoft Teams PowerShell module (Y/N)?"
                } until (-not [string]::IsNullOrWhiteSpace($updateResponse))
                if ($updateResponse -eq 'Y' -or $updateResponse -eq 'y') {
                    Write-Host -ForegroundColor $Colors.ProcessMessage "Updating Microsoft Teams module - Administration escalation required"
                    Invoke-ElevatedPowerShellCommand -ShellPath $elevatedShellPath -CommandText "Update-Module -Name MicrosoftTeams -Force -Confirm:`$false -ErrorAction Stop"
                }
            }
            else {
                Write-Host -ForegroundColor $Colors.ProcessMessage "Updating Microsoft Teams module - Administration escalation required"
                Invoke-ElevatedPowerShellCommand -ShellPath $elevatedShellPath -CommandText "Update-Module -Name MicrosoftTeams -Force -Confirm:`$false -ErrorAction Stop"
            }
        }
        else {
            Write-Host -ForegroundColor $Colors.ProcessMessage "Local module $localVersion is current"
        }
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Loading Microsoft Teams PowerShell module"
    Import-Module MicrosoftTeams -ErrorAction Stop | Out-Null

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
        $generatedCertificate = New-TeamsLocalCertificate -SubjectName $GeneratedCertSubject -YearsValid $GeneratedCertYearsValid -OutputPath $GeneratedCertOutputPath -ExportPfx:$ExportGeneratedPfx -PfxPassword $GeneratedPfxPassword -NoPrompt:$noprompt -FriendlyName $certFriendlyName

        Write-Host -ForegroundColor $Colors.ProcessMessage "Generated certificate thumbprint: $($generatedCertificate.Thumbprint)"
        Write-Host -ForegroundColor $Colors.ProcessMessage "Public certificate exported to: $($generatedCertificate.CerPath)"
        if (-not [string]::IsNullOrWhiteSpace($generatedCertificate.PfxPath)) {
            Write-Host -ForegroundColor $Colors.ProcessMessage "PFX exported to: $($generatedCertificate.PfxPath)"
        }

        if ($ProvisionEntraApp) {
            $resolvedDisplayName = if ([string]::IsNullOrWhiteSpace($AppDisplayName)) { $GeneratedCertSubject } else { $AppDisplayName }
            Write-Host -ForegroundColor $Colors.ProcessMessage "`nStarting Graph authentication for app provisioning..."
            $graphToken = Get-DeviceCodeGraphToken -TenantId $provisionTenant -ClientId $SetupClientId -CopyCodeToClipboard:$CopyDeviceCodeToClipboard -Colors $Colors

            $teamsProvisioningTargets = Get-TeamsProvisioningRoleTargets -AccessToken $graphToken -GraphResourceAppId $GraphResourceAppId -Colors $Colors
            $provisionResult = Get-OrCreateTeamsEntraApplication -AccessToken $graphToken -DisplayName $resolvedDisplayName -Thumbprint $generatedCertificate.Thumbprint -Colors $Colors

            Set-TeamsEntraRequiredResourceAccess -AccessToken $graphToken -AppObjectId $provisionResult.id -RequiredResourceAccess $teamsProvisioningTargets.RequiredResourceAccess -Colors $Colors

            $teamsServicePrincipal = Get-OrCreateTeamsEntraServicePrincipal -AccessToken $graphToken -AppId $provisionResult.appId -Colors $Colors

            Set-TeamsEntraAppRoleAssignments -AccessToken $graphToken -ServicePrincipalId $teamsServicePrincipal.id -GraphServicePrincipalId $teamsProvisioningTargets.GraphServicePrincipal.id -GraphRoleIds $teamsProvisioningTargets.GraphRoleIds -GraphRoleValues $teamsProvisioningTargets.RequiredGraphRoleValues -Colors $Colors
            Set-TeamsEntraDirectoryRoleAssignment -AccessToken $graphToken -ServicePrincipalObjectId $teamsServicePrincipal.id -DirectoryRoleName $EntraDirectoryRoleName -Colors $Colors

            $mapPath = $CertificateMapPath
            $profileEntry = [PSCustomObject]@{
                name = $resolvedDisplayName
                tenant = $provisionTenant
                organization = $provisionTenant
                appId = $provisionResult.appId
                certificateThumbprint = $generatedCertificate.Thumbprint
            }
            Set-TeamsProfileMapEntry -MapPath $mapPath -ProfileEntry $profileEntry -Colors $Colors

            Write-Host -ForegroundColor $Colors.SystemMessage "`n=== Provisioning complete ==="
            Write-Host -ForegroundColor $Colors.ProcessMessage "Entra App Summary:"
            Write-Host -ForegroundColor $Colors.ProcessMessage "  Display Name:   $resolvedDisplayName"
            Write-Host -ForegroundColor $Colors.ProcessMessage "  Object ID:      $($provisionResult.id)"
            Write-Host -ForegroundColor $Colors.ProcessMessage "App ID:          $($provisionResult.appId)"
            Write-Host -ForegroundColor $Colors.ProcessMessage "Cert Thumbprint: $($generatedCertificate.Thumbprint)"
            Write-Host -ForegroundColor $Colors.ProcessMessage "Tenant / Org:    $provisionTenant"
            Write-Host -ForegroundColor $Colors.ProcessMessage "Profile Map:     $CertificateMapPath"
            Write-Host -ForegroundColor $Colors.WarningMessage "If the app registration needs additional admin-consented permissions for your Teams workflow, grant them in Entra ID before using the certificate."
            Write-Host -ForegroundColor $Colors.ProcessMessage "`nConnect any time using:"
            Write-Host -ForegroundColor $Colors.ProcessMessage "  .\\o365-connect-tms-cert.ps1 -UseCertificateAuth -Tenant '$provisionTenant'"
        }

        Write-Host -ForegroundColor $Colors.SystemMessage "`nCertificate generation finished`n"
        exit 0
    }

    Write-Debug "Resolving profile and certificate auth inputs."
    $resolvedProfile = Resolve-TeamsCertificateProfile -Path $CertificateMapPath -TenantFilter $Tenant -ProfileFilter $ProfileName -OrganizationFilter $Organization -NoPrompt:$noprompt -Colors $Colors
    if ($null -ne $resolvedProfile) {
        if ([string]::IsNullOrWhiteSpace($Organization)) {
            if (-not [string]::IsNullOrWhiteSpace($resolvedProfile.organization)) {
                $Organization = $resolvedProfile.organization
            }
            elseif (-not [string]::IsNullOrWhiteSpace($resolvedProfile.tenant)) {
                $Organization = $resolvedProfile.tenant
            }
        }
        if ([string]::IsNullOrWhiteSpace($AppId)) {
            $AppId = $resolvedProfile.appId
        }
        if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
            $CertificateThumbprint = $resolvedProfile.certificateThumbprint
        }
    }

    # Treat -Tenant as an alias for -Organization in certificate-auth mode.
    if ([string]::IsNullOrWhiteSpace($Organization) -and -not [string]::IsNullOrWhiteSpace($Tenant)) {
        $Organization = $Tenant
    }

    $missingFields = [System.Collections.Generic.List[string]]::new()
    if ([string]::IsNullOrWhiteSpace($Organization)) { $missingFields.Add('Organization (or Tenant)') }
    if ([string]::IsNullOrWhiteSpace($AppId)) { $missingFields.Add('AppId') }
    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $missingFields.Add('CertificateThumbprint') }

    if ($missingFields.Count -gt 0) {
        throw "UseCertificateAuth is missing required value(s): $($missingFields -join ', '). Provide them directly, or complete provisioning to create/update profile map '$CertificateMapPath'."
    }

    $CertificateThumbprint = ($CertificateThumbprint -replace '\s+', '').Trim()
    if ($CertificateThumbprint -notmatch '^[0-9A-Fa-f]{40}$') {
        throw "CertificateThumbprint '$CertificateThumbprint' is not a valid SHA-1 thumbprint (40 hex characters)."
    }

    $localCert = Get-Item "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
    if ($null -eq $localCert) {
        throw "Certificate with thumbprint '$CertificateThumbprint' not found in Cert:\CurrentUser\My. Import the PFX or run -GenerateLocalCertificate -ProvisionEntraApp on this machine first."
    }

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

    $runEntraValidationOnConnect = $ValidateEntraOnConnect
    if ($SkipEntraValidationOnConnect) {
        $runEntraValidationOnConnect = $false
    }

    if ($runEntraValidationOnConnect) {
        Write-Host -ForegroundColor $Colors.ProcessMessage "Ensuring the Teams application service principal and directory role are present..."
        try {
            $graphTokenForTeamsRole = Get-DeviceCodeGraphToken -TenantId $Organization -ClientId $SetupClientId -CopyCodeToClipboard:$CopyDeviceCodeToClipboard -Colors $Colors
            $teamsAppServicePrincipal = Get-OrCreateTeamsEntraServicePrincipal -AccessToken $graphTokenForTeamsRole -AppId $AppId -Colors $Colors
            Set-TeamsEntraDirectoryRoleAssignment -AccessToken $graphTokenForTeamsRole -ServicePrincipalObjectId $teamsAppServicePrincipal.id -DirectoryRoleName $EntraDirectoryRoleName -Colors $Colors
            Write-Host -ForegroundColor $Colors.ProcessMessage "Waiting briefly for role assignment propagation..."
            Start-Sleep -Seconds 20
        }
        catch {
            Write-Host -ForegroundColor $Colors.WarningMessage "Unable to verify or assign the Teams directory role from Graph during connection setup: $($_.Exception.Message)"
            Write-Host -ForegroundColor $Colors.WarningMessage "If Teams cmdlets continue to fail with Authorization_IdentityNotFound, assign role '$EntraDirectoryRoleName' to the service principal for AppId '$AppId' in Entra ID and re-run the script."
            Write-Host -ForegroundColor $Colors.WarningMessage "The device-code flow may also require a tenant-approved public client; if prompted, use a custom -SetupClientId from your own Entra app registration."
        }
    }
    elseif ($SkipEntraValidationOnConnect -and $ValidateEntraOnConnect) {
        Write-Host -ForegroundColor $Colors.WarningMessage "Skipping Graph pre-check because -SkipEntraValidationOnConnect overrides -ValidateEntraOnConnect."
    }

    $existingSession = $null
    if (Get-Command Get-CsOnlineSession -ErrorAction SilentlyContinue) {
        $existingSession = Get-CsOnlineSession -ErrorAction SilentlyContinue
    }

    if ($null -ne $existingSession) {
        $tenantMatches = $false

        $sessionTenantCandidates = @()
        if ($existingSession.PSObject.Properties.Name -contains 'TenantId' -and -not [string]::IsNullOrWhiteSpace($existingSession.TenantId)) {
            $sessionTenantCandidates += $existingSession.TenantId
        }
        if ($existingSession.PSObject.Properties.Name -contains 'Tenant' -and -not [string]::IsNullOrWhiteSpace($existingSession.Tenant)) {
            $sessionTenantCandidates += $existingSession.Tenant
        }
        if ($existingSession.PSObject.Properties.Name -contains 'Account' -and -not [string]::IsNullOrWhiteSpace($existingSession.Account)) {
            $sessionTenantCandidates += $existingSession.Account
        }

        foreach ($candidate in ($sessionTenantCandidates | Select-Object -Unique)) {
            if ($candidate -ieq $Organization) {
                $tenantMatches = $true
                break
            }
        }

        if ($tenantMatches) {
            Write-Host -ForegroundColor $Colors.ProcessMessage "Already connected to $Organization - skipping reconnect."
            Write-TeamsConnectedTenant -RequestedOrganization $Organization -Colors $Colors
            Write-TeamsCertConnectionDetails -TenantId $Organization -AppId $AppId -LocalCert $localCert -Colors $Colors
            Write-Host -ForegroundColor $Colors.SystemMessage "Microsoft Teams certificate auth flow finished`n"
            exit 0
        }
    }

    Write-Host -ForegroundColor $Colors.ProcessMessage "Connecting to Microsoft Teams with certificate authentication"
    if (-not (Get-Command Connect-MicrosoftTeams -ErrorAction SilentlyContinue)) {
        throw "The Connect-MicrosoftTeams cmdlet is not available after importing the Microsoft Teams module."
    }

    Connect-MicrosoftTeams -TenantId $Organization -ApplicationId $AppId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop | Out-Null
    $disconnectTeamsAuthOnError = $true
    Write-TeamsConnectedTenant -RequestedOrganization $Organization -Colors $Colors
    Write-TeamsCertConnectionDetails -TenantId $Organization -AppId $AppId -LocalCert $localCert -Colors $Colors
    $disconnectTeamsAuthOnError = $false

    Write-Host -ForegroundColor $Colors.ProcessMessage "Connected to Microsoft Teams`n"
    Write-Host -ForegroundColor $Colors.SystemMessage "Microsoft Teams certificate auth flow finished`n"
}
catch {
    if ($disconnectTeamsAuthOnError) {
        Disconnect-MicrosoftTeams -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    Write-Host -ForegroundColor $Colors.ErrorMessage "Script failed: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($enableLog) {
        Stop-Transcript | Out-Null
    }
}
