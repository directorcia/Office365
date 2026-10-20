<#
.SYNOPSIS
    Read-only check of M365/Entra tenants for exposure to the October 2026
    ConnectWise Cloud Backup (formerly SkyKick) vendor-credential breach.
.DESCRIPTION
    Source: https://www.huntress.com/blog/vendor-credential-enables-unauthorized-access
    STRICTLY READ-ONLY. Every Graph call is an HTTP GET (enforced by Invoke-GraphGet).
    Checks per tenant:
      1. ConnectWise/SkyKick service principals (7 app IDs, display names, owner tenant
         3f9cd7a0-127a-4ed4-bf06-2f830d5616b7): enabled state, app role assignments,
         OAuth2 grants, directory roles.
      2. Service principal sign-ins since 1 Oct 2026 for those apps and IOC IPs (needs Entra ID P1).
            IOC IPs are loaded from a confidence-weighted JSON file; all sign-in types from those IPs
            are also checked for the Huntress 15-16 Oct audit window.
        3. Directory audit since 13 Oct 2026 (Add member to role / Update application /
            Add service principal credentials) plus a scan of all app registrations and
            service principals for credentials whose validity starts since 13 Oct. Graph
            credential objects do not expose their creation time.
      4. Conditional Access policies excluding those apps (needs Policy.Read.All).
    5. Verdict: Not present / Present-no-activity / Action-required / Review-credentials / Suspicious.
    A check that fails for licence/permission reasons is reported as 'Unavailable', not fatal.
    One tenant failing never aborts the run.

    Required Graph read permissions (delegated scopes or application permissions):
      Application.Read.All, Directory.Read.All, AuditLog.Read.All,
      RoleManagement.Read.Directory, Policy.Read.All
    Licensing: sign-in logs require Entra ID P1/P2 in the target tenant.
    Requires PowerShell 7.4+ and Microsoft.Graph.Authentication.
.PARAMETER TenantId
    Interactive mode: tenant ID or domain to check (delegated sign-in).
.PARAMETER UseDeviceCode
    Interactive mode: use device-code sign-in (headless consoles).
.PARAMETER TenantCsv
    Unattended mode: CSV with columns TenantId, ClientId, and CertificateThumbprint
    and/or CertificatePath. Optional: Name, CertificatePasswordEnvironmentVariable.
.PARAMETER Tenants
    Unattended mode: array of hashtables/objects with the same properties as the CSV.
    For GDAP/partner access, use a multi-tenant app consented in each customer tenant
    and list each customer TenantId with the partner ClientId/certificate.
.PARAMETER CertificatePasswordEnvironmentVariable
    Default env var holding the PFX password when CertificatePath is used. Never printed.
.PARAMETER OutputPath
    Report folder. Defaults to the current folder.
.PARAMETER SinceSignIns / SinceAudit
    UTC start dates. Defaults 2026-10-01 and 2026-10-13.
.PARAMETER IocConfigPath
    JSON file containing IOC IPs, confidence, and activity notes. Defaults to the sidecar JSON.
.PARAMETER IocAuditStart / IocAuditEnd
    UTC boundaries for the IOC sign-in audit window. End is exclusive; defaults to 15-17 Oct 2026.
.PARAMETER Headless
    Suppress the opening of the HTML report.
.PARAMETER DryRun
    Validate tenant and IOC inputs only. Do not connect to Graph or write reports.
.EXAMPLE
    ./cw-skykick-exposure-check.ps1 -TenantId contoso.onmicrosoft.com -UseDeviceCode
.EXAMPLE
    $env:SECMON_PFX_PASSWORD = '<from vault>'
    ./cw-skykick-exposure-check.ps1 -TenantCsv ./tenants.csv -CertificatePasswordEnvironmentVariable SECMON_PFX_PASSWORD -Headless -OutputPath ./reports
.LINK
    https://github.com/directorcia/Office365/wiki/ConnectWise-SkyKick-Exposure-Check
.NOTES
    Author: CIAOPS. Makes no changes to any tenant.
#>
[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(ParameterSetName = 'Interactive', Mandatory)][string]$TenantId,
    [Parameter(ParameterSetName = 'Interactive')][switch]$UseDeviceCode,
    [Parameter(ParameterSetName = 'Csv', Mandatory)][string]$TenantCsv,
    [Parameter(ParameterSetName = 'Array', Mandatory)][object[]]$Tenants,
    [string]$CertificatePasswordEnvironmentVariable = 'SECMON_PFX_PASSWORD',
    [string]$OutputPath = (Get-Location).Path,
    [datetime]$SinceSignIns = [datetime]::new(2026,10,1,0,0,0,[DateTimeKind]::Utc),
    [datetime]$SinceAudit   = [datetime]::new(2026,10,13,0,0,0,[DateTimeKind]::Utc),
    [string]$IocConfigPath = (Join-Path $PSScriptRoot 'cw-skykick-iocs.json'),
    [datetime]$IocAuditStart = [datetime]::new(2026,10,15,0,0,0,[DateTimeKind]::Utc),
    [datetime]$IocAuditEnd = [datetime]::new(2026,10,17,0,0,0,[DateTimeKind]::Utc),
    [switch]$Headless,
    [switch]$DryRun
)

Set-StrictMode -Version 3
$ErrorActionPreference = 'Stop'

#region IOCs
$CwAppIds = [ordered]@{
    '506e08f9-6e9a-4da4-942b-a099a03644ae' = 'ConnectWise Cloud Backup (SkyKick)'
    '2399d300-6fca-4e6a-a686-399fefe75b83' = 'ConnectWise Cloud Management'
    '24205314-2b89-4e07-a2a4-579e898a5640' = 'Cloud Management - MS Teams'
    'b78f6a77-d4e5-46fb-87e5-d1648bd02519' = 'Cloud Management - Exchange Online'
    '0de14c77-b9b0-4137-9360-34878a929827' = 'Cloud Management - MS Graph'
    '412dcf96-176d-4b3a-a1b0-4981ad66971e' = 'Cloud Management - Partner Center'
    'b2b55f5b-7758-4719-8d8f-88ae172e6bdb' = 'Cloud Management - SharePoint Online'
}
$CwOwnerTenant = '3f9cd7a0-127a-4ed4-bf06-2f830d5616b7'
$CwNamePattern = 'ConnectWise|SkyKick'
$iocConfig = Get-Content -LiteralPath $IocConfigPath -Raw | ConvertFrom-Json
if (-not $iocConfig.PSObject.Properties['indicators'] -or -not $iocConfig.indicators) {
    throw "IOC config '$IocConfigPath' must contain a non-empty 'indicators' array."
}
$IocByIp = @{}
foreach ($indicator in $iocConfig.indicators) {
    $parsedIp = $null
    if (-not $indicator.ip -or -not [System.Net.IPAddress]::TryParse([string]$indicator.ip, [ref]$parsedIp)) {
        throw "IOC config '$IocConfigPath' contains an invalid IP address."
    }
    if ($indicator.confidence -notin 'High','Medium','Low') {
        throw "IOC '$($indicator.ip)' must have confidence High, Medium, or Low."
    }
    if ($IocByIp.ContainsKey([string]$indicator.ip)) { throw "IOC config contains duplicate IP '$($indicator.ip)'." }
    $IocByIp[[string]$indicator.ip] = $indicator
}
$IocIps        = @($IocByIp.Keys)
$IocUaPattern  = 'Python-urllib|python-requests|curl/'
$PrivRoles     = @('62e90394-69f5-4237-9190-012177145e10','29232cdf-9323-42fd-ade2-1d097af3e4de','7be44c8a-adaf-4e2a-84d6-ab2649e08a13','9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3','158c047a-c907-4556-b7ef-446551a6b5f7') # GA, Exchange Admin, Priv Auth Admin, App Admin, Cloud App Admin
$PrivRoleNames = @('Global Administrator','Exchange Administrator','Privileged Authentication Administrator','Application Administrator','Cloud Application Administrator')
$SydneyTz      = try { [TimeZoneInfo]::FindSystemTimeZoneById('Australia/Sydney') } catch { [TimeZoneInfo]::Local }
$RequiredPerms = 'Application.Read.All','Directory.Read.All','AuditLog.Read.All','RoleManagement.Read.Directory','Policy.Read.All'
#endregion

#region helpers
function Write-Status([string]$Msg, [string]$Colour = 'Gray') { Write-Host $Msg -ForegroundColor $Colour }
function ConvertTo-Local($utc) {
    if (-not $utc) { return $null }
    $d = ([datetimeoffset]$utc).UtcDateTime
    $local = [TimeZoneInfo]::ConvertTimeFromUtc($d, $SydneyTz)
    $zone = if ($SydneyTz.IsDaylightSavingTime($local)) { 'AEDT' } else { 'AEST' }
    $local.ToString('yyyy-MM-dd HH:mm:ss') + " $zone"
}
function Invoke-GraphGet {
    # Read-only wrapper: GET only, follows @odata.nextLink, retries 429/503/504 honouring Retry-After.
    param([Parameter(Mandatory)][string]$Uri, [switch]$Eventual, [int]$MaxPages = 1000)
    $headers = @{}; if ($Eventual) { $headers['ConsistencyLevel'] = 'eventual' }
    $results = [System.Collections.Generic.List[object]]::new()
    $next = $Uri; $page = 0
    while ($next -and $page -lt $MaxPages) {
        $attempt = 0
        while ($true) {
            Write-Progress -Id 2 -Activity 'Microsoft Graph read-only request' -Status "Requesting page $($page + 1) (attempt $($attempt + 1))"
            try {
                $resp = Invoke-MgGraphRequest -Method GET -Uri $next -Headers $headers -OutputType PSObject -ErrorAction Stop
                break
            } catch {
                $code = $null; try { $code = [int]$_.Exception.Response.StatusCode } catch {}
                if ($code -notin 429,503,504) {
                    $details = if ($_.ErrorDetails) { $_.ErrorDetails.Message } else { '' }
                    if ($details -match 'TooManyRequests') { $code = 429 }
                    elseif ($details -match 'ServiceUnavailable') { $code = 503 }
                    elseif ($details -match 'GatewayTimeout') { $code = 504 }
                }
                if ($code -in 429,503,504 -and $attempt -lt 6) {
                    $wait = 5 * [math]::Pow(2, $attempt)
                    try { $ra = $_.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds; if ($ra) { $wait = [math]::Ceiling($ra) } } catch {}
                    Write-Status "    Throttled ($code); waiting $wait s" Yellow
                    Start-Sleep -Seconds $wait; $attempt++; continue
                }
                Write-Progress -Id 2 -Activity 'Microsoft Graph read-only request' -Completed
                throw
            }
        }
        if ($resp.PSObject.Properties['value']) { foreach ($v in $resp.value) { $results.Add($v) } } else { $results.Add($resp) }
        $next = if ($resp.PSObject.Properties['@odata.nextLink']) { $resp.'@odata.nextLink' } else { $null }
        $page++
    }
    Write-Progress -Id 2 -Activity 'Microsoft Graph read-only request' -Completed
    if ($next) { throw "Graph request exceeded the $MaxPages-page limit; results are incomplete." }
    , $results.ToArray()
}
function Get-ErrorReason($err) {
    $m = "$($err.Exception.Message) $(if ($err.ErrorDetails) { $err.ErrorDetails.Message })"
    if ($m -match 'Authentication_RequestFromNonPremiumTenantOrB2CTenant|premium|license') { return 'Unavailable: tenant lacks Entra ID P1 licence' }
    if ($m -match 'Forbidden|Authorization_RequestDenied|403|insufficient privileges|PermissionScopeNotGranted') { return 'Unavailable: missing Graph permission' }
    return "Error: $(($m -replace '\s+',' ').Substring(0,[math]::Min(250,$m.Length)))"
}
function Invoke-Check([string]$Name, [scriptblock]$Body, $Result) {
    try { $output = & $Body; $Result.Checks[$Name] = 'OK'; return $output }
    catch { $r = Get-ErrorReason $_; $Result.Checks[$Name] = $r; Write-Status "    $Name -> $r" Yellow; return $null }
}
function ConvertFrom-PfxFile([string]$Path, [string]$EnvVar) {
    $pw = [Environment]::GetEnvironmentVariable($EnvVar)
    if (-not $pw) { throw "Environment variable '$EnvVar' is empty; cannot open PFX." }
    [System.Security.Cryptography.X509Certificates.X509Certificate2]::new((Resolve-Path $Path).Path, $pw,
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet)
}
#endregion

function Test-Tenant {
    param($T)
    $r = [ordered]@{
        Tenant = $T.TenantId; Name = $T.Name; TenantDisplayName = $null; Verdict = 'Unknown'
        Checks = [ordered]@{}; GrantedPermissions = @(); MissingPermissions = @()
        ServicePrincipals = @(); SignIns = @(); AuditEvents = @(); NewCredentials = @(); CaExclusions = @()
        Recommendations = @(); Error = $null; CheckedAt = (ConvertTo-Local ([datetime]::UtcNow))
    }
    Write-Status "`n=== Tenant $($T.TenantId) ===" Cyan
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        if ($T.ClientId) {
            if ($T.CertificatePath) {
                $envVar = if ($T.CertificatePasswordEnvironmentVariable) { $T.CertificatePasswordEnvironmentVariable } else { $CertificatePasswordEnvironmentVariable }
                $cert = ConvertFrom-PfxFile $T.CertificatePath $envVar
                if ($T.CertificateThumbprint -and $cert.Thumbprint -ne $T.CertificateThumbprint) { throw "PFX thumbprint $($cert.Thumbprint) does not match expected $($T.CertificateThumbprint)" }
                Connect-MgGraph -TenantId $T.TenantId -ClientId $T.ClientId -Certificate $cert -NoWelcome
            } else {
                Connect-MgGraph -TenantId $T.TenantId -ClientId $T.ClientId -CertificateThumbprint $T.CertificateThumbprint -NoWelcome
            }
        } else {
            $p = @{ TenantId = $T.TenantId; Scopes = $RequiredPerms; NoWelcome = $true }
            if ($T.UseDeviceCode) { $p.UseDeviceCode = $true }
            Connect-MgGraph @p
        }
        $ctx = Get-MgContext
        $r.GrantedPermissions = @($ctx.Scopes)
        if ($ctx.PSObject.Properties['AuthType'] -and $ctx.AuthType -eq 'AppOnly') {
            $r.MissingPermissions = @()
            Write-Status '  App-only authentication; delegated-scope comparison skipped' DarkGray
        } else {
            $r.MissingPermissions = @($RequiredPerms | Where-Object { $_ -notin $ctx.Scopes -and ($_ -replace 'Read','ReadWrite') -notin $ctx.Scopes })
        }
        if ($r.MissingPermissions) { Write-Status "  Missing permissions: $($r.MissingPermissions -join ', ')" Yellow }
        try { $org = Invoke-GraphGet 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName'; $r.TenantDisplayName = $org[0].displayName } catch {}
    } catch {
        $r.Error = "Connect failed: $($_.Exception.Message)"; $r.Verdict = 'Error'
        Write-Status "  $($r.Error)" Red
        return [pscustomobject]$r
    }

    $suspicious = [System.Collections.Generic.List[string]]::new()
    $allSps = @(); $cwSps = @(); $cwSpIds = @()

    # ---- Check 1
    Write-Status '  [1] Service principals'
    $spData = Invoke-Check 'ServicePrincipals' {
        $all = Invoke-GraphGet 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=id,appId,displayName,accountEnabled,appOwnerOrganizationId,createdDateTime,passwordCredentials,keyCredentials&$top=999'
        $cwSps = @($all | Where-Object { $CwAppIds.Contains($_.appId) -or $_.displayName -match $CwNamePattern -or $_.appOwnerOrganizationId -eq $CwOwnerTenant })
        $spIndex = 0
        foreach ($sp in $cwSps) {
            $spIndex++
            Write-Status "      Inspecting matching service principal $spIndex/$($cwSps.Count): $($sp.displayName)" DarkGray
            $roles = @(); $grants = @(); $appRoles = @()
            try {
                $ra = Invoke-GraphGet "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$filter=principalId eq '$($sp.id)'&`$expand=roleDefinition"
                $roles = @($ra | ForEach-Object { [pscustomobject]@{ Role = $_.roleDefinition.displayName; RoleId = $_.roleDefinition.templateId; Scope = $_.directoryScopeId } })
            } catch { $roles = @([pscustomobject]@{ Role = (Get-ErrorReason $_); RoleId = $null; Scope = $null }) }
            try {
                $grants = @(Invoke-GraphGet "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$($sp.id)'" | ForEach-Object { "$($_.consentType): $($_.scope.Trim())" })
            } catch { $grants = @(Get-ErrorReason $_) }
            try {
                $resCache = @{}
                $appRoles = @(Invoke-GraphGet "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/appRoleAssignments" | ForEach-Object {
                    $rid = $_.resourceId
                    if (-not $resCache[$rid]) { $resCache[$rid] = Invoke-GraphGet "https://graph.microsoft.com/v1.0/servicePrincipals/$rid`?`$select=displayName,appRoles" | Select-Object -First 1 }
                    $res = $resCache[$rid]; $rv = ($res.appRoles | Where-Object id -eq $_.appRoleId | Select-Object -First 1).value
                    "$($res.displayName): $rv (granted $(ConvertTo-Local $_.createdDateTime))"
                })
            } catch { $appRoles = @(Get-ErrorReason $_) }
            $priv = @($roles | Where-Object { $_.RoleId -in $PrivRoles })
            if ($priv) { $suspicious.Add("$($sp.displayName) holds privileged role(s): $(($priv | ForEach-Object { $_.Role }) -join ', ')") }
            $r.ServicePrincipals += [pscustomobject]@{
                DisplayName = $sp.displayName; AppId = $sp.appId; ObjectId = $sp.id; Enabled = $sp.accountEnabled
                OwnerTenant = $sp.appOwnerOrganizationId; Created = (ConvertTo-Local $sp.createdDateTime)
                KnownIoCAppId = $CwAppIds.Contains($sp.appId)
                DirectoryRoles = (($roles | ForEach-Object { $_.Role }) -join '; '); AppRoleAssignments = ($appRoles -join '; '); OAuth2Grants = ($grants -join '; ')
            }
        }
        Write-Status "      Found $($cwSps.Count) ConnectWise/SkyKick SP(s)" $(if ($cwSps) {'Yellow'} else {'Green'})
        [pscustomobject]@{ All = $all; Matches = $cwSps }
    } $r
    if ($spData) {
        $allSps = @($spData.All)
        $cwSps = @($spData.Matches)
        $cwSpIds = @($cwSps | ForEach-Object { $_.id })
    }

    # ---- Check 2
    Write-Status '  [2] Service principal sign-ins'
    Invoke-Check 'SignIns' {
        $since = $SinceSignIns.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        $base = "https://graph.microsoft.com/beta/auditLogs/signIns?`$filter=createdDateTime ge $since and signInEventTypes/any(t: t eq 'servicePrincipal')"
        $found = @()
        $appIdsToQuery = @($CwAppIds.Keys) + @($cwSps | ForEach-Object { $_.appId }) | Select-Object -Unique
        Write-Host '      Sign-in query progress (one dot per query): ' -NoNewline -ForegroundColor Gray
        try {
            foreach ($a in $appIdsToQuery) {
                Write-Host '.' -NoNewline -ForegroundColor Cyan
                $found += Invoke-GraphGet "$base and appId eq '$a'&`$top=999"
            }
            foreach ($ip in $IocIps) {
                Write-Host '.' -NoNewline -ForegroundColor Cyan
                $found += Invoke-GraphGet "$base and ipAddress eq '$ip'&`$top=999"
            }
        } finally { Write-Host '' }
        $found = @($found | Sort-Object id -Unique)
        foreach ($s in $found) {
            $ipIoc = $IocByIp[[string]$s.ipAddress]
            $uaIoc = $s.userAgent -match $IocUaPattern
            $ioc = [bool]$ipIoc -or $uaIoc
            $successful = $null -ne $s.status.errorCode -and [int]$s.status.errorCode -eq 0
            $servicePrincipalSignIn = @($s.signInEventTypes) -contains 'servicePrincipal'
            $severity = if ($ipIoc -and $ipIoc.successfulAppSignInSeverity -and $successful -and $servicePrincipalSignIn) { $ipIoc.successfulAppSignInSeverity }
                elseif ($ioc) { 'Medium' } else { 'None' }
            $confidence = if ($ipIoc) { $ipIoc.confidence } elseif ($uaIoc) { 'Low' } else { 'None' }
            if ($ioc) { $suspicious.Add("$severity severity ($confidence confidence) sign-in by $($s.appDisplayName) from $($s.ipAddress) UA '$($s.userAgent)' at $(ConvertTo-Local $s.createdDateTime)") }
            $r.SignIns += [pscustomobject]@{
                SignInId = $s.id; Time = (ConvertTo-Local $s.createdDateTime); App = $s.appDisplayName; AppId = $s.appId; IP = $s.ipAddress
                SignInType = (@($s.signInEventTypes) -join ', '); UserAgent = $s.userAgent; Resource = $s.resourceDisplayName
                Status = $s.status.errorCode; IoC = $ioc; IocConfidence = $confidence; Severity = $severity
            }
        }
        Write-Status "      $($found.Count) sign-in(s)"
    } $r

    # ---- Check 2a
    Write-Status '  [2a] IOC sign-ins during the 15-16 Oct audit window'
    Invoke-Check 'IocAuditWindow' {
        if ($IocAuditStart.ToUniversalTime() -ge $IocAuditEnd.ToUniversalTime()) { throw 'IocAuditEnd must be later than IocAuditStart.' }
        $start = $IocAuditStart.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        $end = $IocAuditEnd.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        $windowFound = @()
        foreach ($ip in $IocIps) {
            Write-Host '.' -NoNewline -ForegroundColor Cyan
            $filter = "createdDateTime ge $start and createdDateTime lt $end and ipAddress eq '$ip'"
            $windowFound += Invoke-GraphGet "https://graph.microsoft.com/beta/auditLogs/signIns?`$filter=$filter&`$top=999"
        }
        Write-Host ''
        $existingIds = @($r.SignIns | ForEach-Object SignInId)
        $windowFound = @($windowFound | Sort-Object id -Unique)
        foreach ($s in $windowFound) {
            if ($s.id -in $existingIds) { continue }
            $indicator = $IocByIp[[string]$s.ipAddress]
            $successful = $null -ne $s.status.errorCode -and [int]$s.status.errorCode -eq 0
            $servicePrincipalSignIn = @($s.signInEventTypes) -contains 'servicePrincipal'
            $severity = if ($indicator.successfulAppSignInSeverity -and $successful -and $servicePrincipalSignIn) { $indicator.successfulAppSignInSeverity } else { 'Medium' }
            $suspicious.Add("$severity severity ($($indicator.confidence) confidence) audit-window sign-in by $($s.appDisplayName) from $($s.ipAddress) at $(ConvertTo-Local $s.createdDateTime)")
            $r.SignIns += [pscustomobject]@{
                SignInId = $s.id; Time = (ConvertTo-Local $s.createdDateTime); App = $s.appDisplayName; AppId = $s.appId; IP = $s.ipAddress
                SignInType = (@($s.signInEventTypes) -join ', '); UserAgent = $s.userAgent; Resource = $s.resourceDisplayName
                Status = $s.status.errorCode; IoC = $true; IocConfidence = $indicator.confidence; Severity = $severity
            }
        }
        Write-Status "      $($windowFound.Count) IOC sign-in(s) in audit window"
    } $r

    # ---- Check 3
    Write-Status '  [3] Audit log and new credentials'
    Invoke-Check 'AuditLog' {
        $since = $SinceAudit.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        foreach ($act in 'Add member to role','Update application','Add service principal credentials','Update application – Certificates and secrets management','Update application - Certificates and secrets management') {
            $ev = Invoke-GraphGet "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=activityDateTime ge $since and activityDisplayName eq '$act'&`$top=999"
            foreach ($e in $ev) {
                $initApp = $e.initiatedBy.app
                $ua = ($e.additionalDetails | Where-Object key -eq 'User-Agent' | Select-Object -First 1).value
                $uaIoc = $ua -match $IocUaPattern
                if ($act -ne 'Add member to role' -and -not $initApp -and -not $uaIoc) { continue }   # retain suspicious-UA events without an app initiator
                $byCw = $initApp -and ($initApp.servicePrincipalId -in $cwSpIds -or $CwAppIds.Contains("$($initApp.appId)") -or $initApp.displayName -match $CwNamePattern)
                if ($act -eq 'Add member to role' -and -not $byCw -and -not $initApp -and -not $uaIoc) { continue }
                $selfRoleAssignment = $false
                $assignedRole = $null
                $privilegedRole = $false
                if ($act -eq 'Add member to role' -and $initApp -and $initApp.servicePrincipalId -and $e.result -eq 'success') {
                    $targetServicePrincipalIds = @($e.targetResources | Where-Object type -eq 'ServicePrincipal' | ForEach-Object id)
                    $roleEvidence = @($e.targetResources | ForEach-Object {
                        $_.id
                        $_.displayName
                        $_.modifiedProperties | Where-Object displayName -match 'role|template' | ForEach-Object { $_.oldValue; $_.newValue }
                    }) -join ' '
                    $assignedRole = @($e.targetResources | Where-Object type -eq 'Role' | ForEach-Object displayName | Select-Object -First 1)
                    if (-not $assignedRole) {
                        $assignedRole = @($e.targetResources | ForEach-Object { $_.modifiedProperties | Where-Object displayName -match '^Role\.DisplayName$' | ForEach-Object newValue } | Select-Object -First 1)
                    }
                    $privilegedRoleName = @($PrivRoleNames | Where-Object { $roleEvidence -match "(?i)(?<![\w])$([regex]::Escape($_))(?![\w])" } | Select-Object -First 1)
                    $assignedPrivilegedRoleId = @($PrivRoles | Where-Object { $roleEvidence -match [regex]::Escape($_) } | Select-Object -First 1)
                    $selfRoleAssignment = $initApp.servicePrincipalId -in $targetServicePrincipalIds
                    $privilegedRole = $privilegedRoleName.Count -gt 0 -or $assignedPrivilegedRoleId.Count -gt 0
                    if (-not $assignedRole -and $privilegedRoleName) { $assignedRole = $privilegedRoleName }
                    if (-not $assignedRole -and $assignedPrivilegedRoleId) { $assignedRole = $assignedPrivilegedRoleId }
                }
                $flag = [bool]($byCw -or $selfRoleAssignment -or $uaIoc)
                $initiator = if ($initApp) { "app:$($initApp.displayName) ($($initApp.servicePrincipalId))" } elseif ($e.initiatedBy.user) { "user:$($e.initiatedBy.user.userPrincipalName)" } else { 'unknown' }
                if ($flag) {
                    $roleLabel = if ($assignedRole) { "role '$assignedRole'" } else { 'a directory role' }
                    $reason = if ($selfRoleAssignment) { " HIGH severity: app added itself to $roleLabel" } elseif ($uaIoc) { " suspicious user-agent '$ua'" } else { '' }
                    $suspicious.Add("Audit '$act' by $initiator$reason on $(($e.targetResources | ForEach-Object { $_.displayName }) -join ',') at $(ConvertTo-Local $e.activityDateTime)")
                }
                $r.AuditEvents += [pscustomobject]@{
                    Time = (ConvertTo-Local $e.activityDateTime); Activity = $act; Initiator = $initiator
                    Target = (($e.targetResources | ForEach-Object { $_.displayName }) -join '; '); UserAgent = $ua; Result = $e.result
                    SelfRoleAssignment = [bool]$selfRoleAssignment; PrivilegedRole = [bool]$privilegedRole
                    Severity = $(if ($selfRoleAssignment) {'High'} elseif ($flag) {'Medium'} else {'None'}); Flagged = [bool]$flag
                }
            }
        }
        Write-Status "      $($r.AuditEvents.Count) relevant audit event(s)"
    } $r
    Invoke-Check 'CredentialScan' {
        $apps = Invoke-GraphGet 'https://graph.microsoft.com/v1.0/applications?$select=id,appId,displayName,passwordCredentials,keyCredentials&$top=999'
        $sps = if ($r.Checks['ServicePrincipals'] -eq 'OK') { $allSps } else { Invoke-GraphGet 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=id,appId,displayName,appOwnerOrganizationId,passwordCredentials,keyCredentials&$top=999' }
        $totalObjects = @($apps).Count + @($sps).Count
        $scannedObjects = 0
        Write-Status "      Scanning credential metadata across $totalObjects app and service-principal objects" DarkGray
        foreach ($set in @(@{K='AppRegistration';L=$apps}, @{K='ServicePrincipal';L=$sps})) {
            foreach ($o in $set.L) {
                $scannedObjects++
                if ($scannedObjects -eq 1 -or $scannedObjects % 100 -eq 0 -or $scannedObjects -eq $totalObjects) {
                    Write-Status "      Credential scan progress: $scannedObjects/$totalObjects objects" DarkGray
                }
                $vendorApp = $CwAppIds.Contains([string]$o.appId) -or $o.displayName -match $CwNamePattern
                if ($set.K -eq 'ServicePrincipal' -and $o.PSObject.Properties['appOwnerOrganizationId']) {
                    $vendorApp = $vendorApp -or $o.appOwnerOrganizationId -eq $CwOwnerTenant
                }
                foreach ($c in @($o.passwordCredentials | Where-Object { $_ }) + @($o.keyCredentials | Where-Object { $_ })) {
                    if ($c.startDateTime -and ([datetimeoffset]$c.startDateTime).UtcDateTime -ge $SinceAudit.ToUniversalTime()) {
                        $r.NewCredentials += [pscustomobject]@{
                            ObjectType = $set.K; DisplayName = $o.displayName; AppId = $o.appId
                            VendorApp = [bool]$vendorApp
                            CredentialType = $(if ($c.PSObject.Properties['secretText'] -or $c.PSObject.Properties['hint']) {'Secret'} else {'Certificate'})
                            CredentialName = $c.displayName; KeyId = $c.keyId; Start = (ConvertTo-Local $c.startDateTime); End = (ConvertTo-Local $c.endDateTime)
                        }
                    }
                }
            }
        }
        Write-Status "      $($r.NewCredentials.Count) credential(s) with validity starting since $($SinceAudit.ToString('yyyy-MM-dd'))"
    } $r

    # ---- Check 4
    Write-Status '  [4] Conditional Access exclusions'
    Invoke-Check 'ConditionalAccess' {
        $pol = Invoke-GraphGet 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies'
        $ids = @($CwAppIds.Keys) + @($cwSps | ForEach-Object { $_.appId }) + @($cwSpIds) | Select-Object -Unique
        foreach ($p in $pol) {
            $exApps = @($p.conditions.applications.excludeApplications)
            $exSps  = @(); try { $exSps = @($p.conditions.clientApplications.excludeServicePrincipals) } catch {}
            $hit = @($exApps + $exSps | Where-Object { $_ -in $ids })
            if ($hit) { $r.CaExclusions += [pscustomobject]@{ Policy = $p.displayName; State = $p.state; ExcludedIds = ($hit -join '; ') } }
        }
        Write-Status "      $($pol.Count) policies, $($r.CaExclusions.Count) with ConnectWise exclusions"
    } $r

    # ---- Verdict
    $nonVendorNewCredentials = @($r.NewCredentials | Where-Object { -not $_.VendorApp })
    $enabledCwApps = @($r.ServicePrincipals | Where-Object { $_.Enabled -eq $true })
    $r.Verdict = if ($suspicious.Count) { 'Suspicious' }
        elseif ($nonVendorNewCredentials.Count) { 'Review-credentials' }
        elseif ($enabledCwApps.Count) { 'Action-required' }
        elseif ($r.ServicePrincipals.Count) { 'Present-no-activity' }
        elseif ($r.Checks['ServicePrincipals'] -eq 'OK') { 'Not present' }
        else { 'Unknown' }
    $rec = @()
    if ($r.ServicePrincipals) { $rec += 'Disable (do NOT delete) the ConnectWise/SkyKick enterprise apps until ConnectWise confirms credential reset.' }
    $highSeverity = @($r.SignIns | Where-Object Severity -eq 'High')
    if ($highSeverity) {
        $highIps = @($highSeverity | Select-Object -ExpandProperty IP -Unique) -join ', '
        $rec += "HIGH severity: investigate successful application sign-ins from $highIps immediately; found $($highSeverity.Count)."
    }
    if ($suspicious) { $rec += $suspicious | ForEach-Object { "Investigate: $_" }; $rec += 'Remove unexpected directory roles and secrets/certificates; revoke sessions; open an incident.' }
    if ($nonVendorNewCredentials) {
        $credentialApps = @($nonVendorNewCredentials | Select-Object -ExpandProperty DisplayName -Unique) -join ', '
        $rec += "Review credentials on non-vendor apps with validity starting since 13 Oct: $credentialApps. Graph does not expose credential creation time."
    }
    if (@($r.NewCredentials | Where-Object VendorApp)) { $rec += 'Confirm every vendor-app secret/certificate whose validity starts since 13 Oct was added by your team; Graph does not expose credential creation time.' }
    if ($r.CaExclusions) { $rec += 'Review Conditional Access policies that exclude the ConnectWise apps.' }
    $unavail = @($r.Checks.GetEnumerator() | Where-Object Value -ne 'OK')
    if ($unavail) { $rec += "Re-run once these checks can complete: $(($unavail | ForEach-Object { "$($_.Key) ($($_.Value))" }) -join '; ')" }
    if (-not $rec) { $rec = @('No action required. Keep monitoring ConnectWise/Huntress advisories.') }
    $r.Recommendations = $rec
    Write-Status "  Verdict: $($r.Verdict)" $(switch ($r.Verdict) { 'Suspicious' {'Red'} 'Not present' {'Green'} default {'Yellow'} })
    [pscustomobject]$r
}

function Write-Html($Results, $Path) {
    Add-Type -AssemblyName System.Web
    $enc = { param($s) [System.Web.HttpUtility]::HtmlEncode("$s") }
    $tbl = { param($rows) if (-not $rows) { return '<p><i>None</i></p>' }
        $cols = $rows[0].PSObject.Properties.Name
        '<table><tr>' + (($cols | ForEach-Object { "<th>$(& $enc $_)</th>" }) -join '') + '</tr>' +
        (($rows | ForEach-Object { $row = $_; '<tr>' + (($cols | ForEach-Object { "<td>$(& $enc $row.$_)</td>" }) -join '') + '</tr>' }) -join '') + '</table>' }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<html><head><meta charset='utf-8'><title>ConnectWise/SkyKick exposure check</title><style>body{font-family:Segoe UI,Arial;margin:20px}table{border-collapse:collapse;margin-bottom:12px}td,th{border:1px solid #ccc;padding:4px 8px;font-size:13px;vertical-align:top}th{background:#eee}.Suspicious{color:#b00}.Not{color:#070}.Present-no-activity,.Action-required,.Review-credentials,.Unknown,.Error{color:#b60}</style></head><body>")
    [void]$sb.Append("<h1>ConnectWise Cloud Backup (SkyKick) exposure check</h1><p>Generated $(ConvertTo-Local ([datetime]::UtcNow)). Read-only. Source: <a href='https://www.huntress.com/blog/vendor-credential-enables-unauthorized-access'>Huntress</a></p>")
    foreach ($r in $Results) {
        $cls = ($r.Verdict -split ' ')[0]
        [void]$sb.Append("<h2>$(& $enc ($r.TenantDisplayName ?? $r.Tenant)) ($(& $enc $r.Tenant)) - <span class='$cls'>$(& $enc $r.Verdict)</span></h2>")
        if ($r.Error) { [void]$sb.Append("<p class='Error'>$(& $enc $r.Error)</p>"); continue }
        [void]$sb.Append('<h3>Check status</h3>' + (& $tbl @($r.Checks.GetEnumerator() | ForEach-Object { [pscustomobject]@{ Check = $_.Key; Status = $_.Value } })))
        if ($r.MissingPermissions) { [void]$sb.Append("<p>Missing permissions: $(& $enc ($r.MissingPermissions -join ', '))</p>") }
        [void]$sb.Append('<h3>1. ConnectWise/SkyKick service principals</h3>' + (& $tbl $r.ServicePrincipals))
        [void]$sb.Append('<h3>2. Service principal sign-ins</h3>' + (& $tbl $r.SignIns))
        [void]$sb.Append('<h3>3a. Audit events</h3>' + (& $tbl $r.AuditEvents))
        [void]$sb.Append('<h3>3b. Secrets/certificates with validity starting since ' + $SinceAudit.ToString('yyyy-MM-dd') + '</h3><p>Graph credential objects expose validity dates, not creation time.</p>' + (& $tbl $r.NewCredentials))
        [void]$sb.Append('<h3>4. Conditional Access exclusions</h3>' + (& $tbl $r.CaExclusions))
        [void]$sb.Append('<h3>Recommended actions</h3><ul>' + (($r.Recommendations | ForEach-Object { "<li>$(& $enc $_)</li>" }) -join '') + '</ul>')
    }
    [void]$sb.Append('</body></html>')
    Set-Content -Path $Path -Value $sb.ToString() -Encoding utf8
}

# ---- main
$list = switch ($PSCmdlet.ParameterSetName) {
    'Interactive' { @([pscustomobject]@{ TenantId = $TenantId; UseDeviceCode = [bool]$UseDeviceCode; Name = $TenantId; ClientId = $null }) }
    'Csv'         { @(Import-Csv $TenantCsv) }
    'Array'       { @($Tenants | ForEach-Object { [pscustomobject]$_ }) }
}
foreach ($t in $list) { foreach ($p in 'Name','ClientId','CertificatePath','CertificateThumbprint','CertificatePasswordEnvironmentVariable','UseDeviceCode') { if (-not $t.PSObject.Properties[$p]) { $t | Add-Member $p $null } } }
if ($DryRun) {
    Write-Status "Dry run: validated $($list.Count) tenant target(s) and IOC config '$IocConfigPath'. No Graph calls or reports were made." Cyan
    return
}
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) { throw 'Microsoft.Graph.Authentication is required: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser' }
Import-Module Microsoft.Graph.Authentication
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }

$results = foreach ($t in $list) {
    try { Test-Tenant $t } catch { [pscustomobject]@{ Tenant = $t.TenantId; Name = $t.Name; TenantDisplayName = $null; Verdict = 'Error'; Error = $_.Exception.Message; Checks = @{}; MissingPermissions = @(); ServicePrincipals=@(); SignIns=@(); AuditEvents=@(); NewCredentials=@(); CaExclusions=@(); Recommendations=@() } }
}
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$base  = Join-Path $OutputPath "cw-skykick-exposure_$stamp"
$results | ForEach-Object { [pscustomobject]@{
    Tenant = $_.Tenant; TenantName = $_.TenantDisplayName; Verdict = $_.Verdict
    ServicePrincipals = @($_.ServicePrincipals).Count; EnabledConnectWiseApps = @($_.ServicePrincipals | Where-Object Enabled -eq $true).Count
    SignIns = @($_.SignIns).Count; HighSeveritySignIns = @($_.SignIns | Where-Object Severity -eq 'High').Count; FlaggedAudit = @($_.AuditEvents | Where-Object Flagged).Count
    ReviewCredentials = @($_.NewCredentials | Where-Object { -not $_.VendorApp }).Count
    NewCredentials = @($_.NewCredentials).Count; CaExclusions = @($_.CaExclusions).Count
    UnavailableChecks = (@($_.Checks.GetEnumerator() | Where-Object Value -ne 'OK' | ForEach-Object { "$($_.Key): $($_.Value)" }) -join ' | ')
    MissingPermissions = (@($_.MissingPermissions) -join ', '); Error = $_.Error } } | Export-Csv "$base.csv" -NoTypeInformation
ConvertTo-Json -InputObject @($results) -Depth 8 | Set-Content "$base.json" -Encoding utf8
Write-Html $results "$base.html"
Write-Status "`nReports:`n  $base.csv`n  $base.json`n  $base.html" Green
if (-not $Headless) {
    try { Start-Process -FilePath "$base.html" -ErrorAction Stop }
    catch { Write-Status "Could not open HTML report: $($_.Exception.Message)" Yellow }
}
$alertResults = @($results | Where-Object { $_.Verdict -in 'Suspicious','Error' })
if ($alertResults) {
    Write-Status "Alert: $($alertResults.Count) tenant result(s) are Suspicious or Error." Red
    exit 1
}
