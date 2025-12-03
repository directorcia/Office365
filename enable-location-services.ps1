param(
    [switch]$ForceReboot,
    [switch]$Silent
)

$exitCode = 0 # 0=success; non-zero indicates specific failure

# Enable Windows Location Services
# Requires: PowerShell running as Administrator
# Documentation: https://github.com/directorcia/Office365/wiki/Enable-Windows-Location-Services

# Check for admin rights
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator. Right-click PowerShell and choose 'Run as administrator'."
    exit 1
}

if ($Silent) {
    try {
        $logDir = 'C:\Temp'
        if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
        $logPath = Join-Path $logDir 'location-enable.log'
        Start-Transcript -Path $logPath -Append -ErrorAction SilentlyContinue | Out-Null
    } catch {}
}
if (-not $Silent) { Write-Host "Enabling Windows Location Services..." -ForegroundColor Cyan }

# 1) Ensure Geolocation Service (lfsvc) is set to start and is running
try {
    $svc = Get-Service -Name 'lfsvc' -ErrorAction Stop
    if ($svc.StartType -ne 'Automatic') {
        if (-not $Silent) { Write-Host "Setting 'lfsvc' start type to Automatic" -ForegroundColor Yellow }
        Set-Service -Name 'lfsvc' -StartupType Automatic
    }
    if ($svc.Status -ne 'Running') {
        if (-not $Silent) { Write-Host "Starting 'lfsvc' service" -ForegroundColor Yellow }
        Start-Service -Name 'lfsvc'
    }
} catch {
    if (-not $Silent) { Write-Warning "Could not configure or start 'lfsvc' (Geolocation Service): $($_.Exception.Message)" }
    # Service failure
    if ($exitCode -eq 0) { $exitCode = 10 }
}

# 2) Clear policy that disables location, if present
# Policy path: HKLM:\SOFTWARE\Policies\Microsoft\Windows\Location -> DisableLocation (DWORD)
try {
    $policyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Location'
    if (Test-Path $policyPath) {
        $disableVal = Get-ItemProperty -Path $policyPath -Name 'DisableLocation' -ErrorAction SilentlyContinue
        if ($disableVal) {
            if (-not $Silent) { Write-Host "Setting policy 'DisableLocation' to 0 (enabled)" -ForegroundColor Yellow }
            Set-ItemProperty -Path $policyPath -Name 'DisableLocation' -Value 0 -Type DWord
        }
    }
} catch {
    if (-not $Silent) { Write-Warning "Could not update policy 'DisableLocation': $($_.Exception.Message)" }
    # Policy update failure
    if ($exitCode -eq 0) { $exitCode = 20 }
}

# 3) Set system configuration flag for Location Services ON
# This is what the Settings toggle controls
# Path: HKLM:\SYSTEM\CurrentControlSet\Services\lfsvc\Service\Configuration -> Status (DWORD 1 = On, 0 = Off)
try {
    $configPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\lfsvc\Service\Configuration'
    if (-not (Test-Path $configPath)) {
        New-Item -Path $configPath -Force | Out-Null
    }
    if (-not $Silent) { Write-Host "Setting Location Services Status=1" -ForegroundColor Yellow }
    New-ItemProperty -Path $configPath -Name 'Status' -PropertyType DWord -Value 1 -Force | Out-Null
} catch {
    if (-not $Silent) { Write-Warning "Could not set Location Services status: $($_.Exception.Message)" }
    # Registry configuration status failure
    if ($exitCode -eq 0) { $exitCode = 30 }
}

# 4) Allow app-level location capability for current user (optional but helpful)
# Path: HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location -> Value = Allow
try {
    $userConsentPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location'
    if (-not (Test-Path $userConsentPath)) {
        New-Item -Path $userConsentPath -Force | Out-Null
    }
    if (-not $Silent) { Write-Host "Setting user location capability to 'Allow'" -ForegroundColor Yellow }
    New-ItemProperty -Path $userConsentPath -Name 'Value' -PropertyType String -Value 'Allow' -Force | Out-Null
} catch {
    if (-not $Silent) { Write-Warning "Could not set user location consent: $($_.Exception.Message)" }
    # User consent failure
    if ($exitCode -eq 0) { $exitCode = 40 }
}

if (-not $Silent) {
    Write-Host "Location Services should now be enabled." -ForegroundColor Green
    Write-Host "If the toggle still appears Off, sign out/in or reboot." -ForegroundColor Green
}

# Optional reboot
if ($ForceReboot) {
    if (-not $Silent) {
        Write-Host "System will reboot in 5 seconds..." -ForegroundColor Red
        Start-Sleep -Seconds 5
        Write-Host "Forcing system reboot to apply changes..." -ForegroundColor Yellow
    }
    if ($Silent) { try { Stop-Transcript | Out-Null } catch {} }
    Restart-Computer -Force
} else {
    if ($Silent) { try { Stop-Transcript | Out-Null } catch {} }
    exit $exitCode
}
