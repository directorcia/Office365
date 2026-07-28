## This script checks for devices registered to Azure AD and removes the local registration data so the device can be joined successfully.
# We recommend backing up the registry before running it. Use at your own risk.

# Source: https://www.inspiredtechs.com.au/fix-for-azure-ad-join-error-code-8018000a-this-device-is-already-enrolled

# This script is intended for Windows devices experiencing Azure AD/Entra ID join issues caused by stale local registration data.
# It removes the local enrollment metadata and related scheduled tasks for each device SID found under the Enterprise Resource Manager registry hive.

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    # Use -Force to skip the interactive confirmation prompt for each SID.
    [switch]$Force
)

# Enable strict mode and stop on terminating errors so unexpected behavior is easier to catch.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Check whether the current process is running with elevated privileges.
# The script modifies HKLM registry keys and scheduled tasks, so Administrator rights are required.
function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Exit early if the script is not being run as Administrator.
if (-not (Test-IsAdministrator)) {
    throw 'This script must be run as Administrator.'
}

# Read all SID values from the Enterprise Resource Manager tracked registrations container.
# These entries represent locally registered device identities that may need cleanup.
try {
    $sids = Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked' -ErrorAction Stop |
        Where-Object { $_.PSChildName.Length -gt 25 } |
        Select-Object -ExpandProperty PSChildName
}
catch {
    throw "Unable to read the Enterprise Resource Manager registry path. $($_.Exception.Message)"
}

# If no registrations are found, inform the user and stop.
if (-not $sids) {
    Write-Host 'No device registrations were found.' -ForegroundColor Green
    return
}

# Process each SID one at a time.
foreach ($sid in $sids) {
    # Build the registry paths that may contain the registration state for this SID.
    $enrollmentPath = "HKLM:\SOFTWARE\Microsoft\Enrollments\$sid"
    $entResourcePath = "HKLM:\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked\$sid"

    $removeDevice = $false

    # If -Force is supplied, skip the interactive prompt and remove the registration directly.
    if ($Force) {
        $removeDevice = $true
    }
    else {
        Write-Host "Found a registered device for SID: $sid" -ForegroundColor Yellow
        $response = Read-Host 'Remove the device registration settings for this SID? (y/n)'
        switch -Regex ($response) {
            '^(y|yes)$' { $removeDevice = $true }
            default { $removeDevice = $false }
        }
    }

    # Skip this SID if the user chose not to remove it.
    if (-not $removeDevice) {
        Write-Host "Removal cancelled for SID: $sid" -ForegroundColor Cyan
        continue
    }

    # Use ShouldProcess so the script supports -WhatIf and -Confirm behavior.
    if ($PSCmdlet.ShouldProcess($sid, 'Remove local Azure AD device registration')) {
        try {
            # Remove the enrollment registry subtree if it exists.
            if (Test-Path -Path $enrollmentPath) {
                Remove-Item -Path $enrollmentPath -Recurse -Force -ErrorAction Stop
            }

            # Remove the tracked resource registry subtree if it exists.
            if (Test-Path -Path $entResourcePath) {
                Remove-Item -Path $entResourcePath -Recurse -Force -ErrorAction Stop
            }

            # Remove any scheduled tasks related to this SID under the EnterpriseMgmt task path.
            $taskPath = "\Microsoft\Windows\EnterpriseMgmt\$sid"
            $tasks = Get-ScheduledTask -TaskPath $taskPath -ErrorAction SilentlyContinue
            foreach ($task in $tasks) {
                Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
            }

            # Remove the EnterpriseMgmt task folder for this SID using the Task Scheduler COM object.
            $scheduleObject = New-Object -ComObject Schedule.Service
            $scheduleObject.Connect()
            $rootFolder = $scheduleObject.GetFolder('\Microsoft\Windows\EnterpriseMgmt')
            if ($rootFolder) {
                $rootFolder.DeleteFolder($sid, $null)
            }

            Write-Host "Device registration cleaned up for SID: $sid" -ForegroundColor Green
        }
        catch {
            # Log any failure for a specific SID without stopping the rest of the script.
            Write-Warning "Failed to remove registration for SID $sid. $($_.Exception.Message)"
        }
    }
}

# Final message to remind the operator that Azure AD/Entra ID registration must also be removed from the cloud side.
Write-Host 'Cleanup of device registration is complete. Ensure the registration is removed from Azure AD before attempting to join the device again.' -ForegroundColor Green


