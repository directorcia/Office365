<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/psmodules-get.ps1

Description - Report installed PowerShell module versions and identify outdated copies

Prerequisites - PowerShell 5.1 or later, PowerShellGet

Examples
  .\psmodules-get.ps1
  .\psmodules-get.ps1 -ModuleName "Microsoft.Graph", "Az"
  .\psmodules-get.ps1 -ShowAll -Csv
  Get-OldPSModules -ModuleName "Microsoft.Online.SharePoint.PowerShell"

More scripts available by joining http://www.ciaopspatron.com

#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Module names to inspect. When omitted, all PowerShellGet-installed modules are scanned.")]
    [Alias('ModuleNames')]
    [string[]]$ModuleName = @(),

    [Parameter(HelpMessage = "Include modules that only have a single installed version.")]
    [switch]$ShowAll,

    [Parameter(HelpMessage = "Export results to CSV.")]
    [switch]$Csv,

    [Parameter(HelpMessage = "CSV output path. Defaults to a timestamped file in the parent directory.")]
    [string]$OutputFile,

    [Parameter(HelpMessage = "Return result objects to the pipeline.")]
    [switch]$PassThru,

    [Parameter(HelpMessage = "Write a transcript log.")]
    [switch]$CreateLog,

    [Parameter(HelpMessage = "Transcript directory. Defaults to the current directory.")]
    [string]$LogDirectory
)

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-ScriptMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('System', 'Process', 'Warning', 'Error')]
        [string]$Type = 'Process'
    )

    $color = switch ($Type) {
        'System' { $systemmessagecolor }
        'Warning' { $warningmessagecolor }
        'Error' { $errormessagecolor }
        default { $processmessagecolor }
    }

    Write-Host -ForegroundColor $color $Message
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        $InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        $Default = $null
    )

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $Default
    }

    return $property.Value
}

function ConvertTo-SortableVersion {
    param(
        [Parameter(Mandatory = $true)]
        $Value
    )

    if ($Value -is [version]) {
        return $Value
    }

    $numeric = ([string]$Value -split '-')[0]
    try {
        return [version]$numeric
    }
    catch {
        return [version]'0.0.0'
    }
}

function Test-PowerShellGetAvailable {
    return [bool](Get-Command -Name Get-InstalledModule -ErrorAction SilentlyContinue)
}

function Get-InstalledModuleInventory {
    param(
        [string[]]$ModuleName = @()
    )

    $requestedNames = @(
        $ModuleName |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.Trim() } |
            Select-Object -Unique
    )

    if ($requestedNames.Count -eq 0) {
        Write-ScriptMessage "[INFO] Querying all installed modules via PowerShellGet. This can take a while."
        return @(Get-InstalledModule -AllVersions -ErrorAction Stop)
    }

    Write-ScriptMessage "[INFO] Querying $($requestedNames.Count) specified module(s) via PowerShellGet."
    $installed = New-Object System.Collections.Generic.List[object]

    foreach ($name in $requestedNames) {
        try {
            $versions = @(Get-InstalledModule -Name $name -AllVersions -ErrorAction Stop)
            foreach ($version in $versions) {
                $installed.Add($version)
            }
        }
        catch {
            Write-ScriptMessage "[WARN] Module not found via PowerShellGet: $name" -Type Warning
        }
    }

    return $installed.ToArray()
}

function Get-OldPSModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [Alias('ModuleNames')]
        [string[]]$ModuleName = @(),

        [Parameter(Mandatory = $false)]
        [switch]$ShowAll
    )

    $results = New-Object System.Collections.Generic.List[object]

    try {
        $installed = Get-InstalledModuleInventory -ModuleName $ModuleName
    }
    catch {
        Write-ScriptMessage "Unable to read installed modules: $($_.Exception.Message)" -Type Error
        Write-ScriptMessage "Confirm PowerShellGet is installed and that modules were installed with Install-Module." -Type Warning
        return @()
    }

    if ($installed.Count -eq 0) {
        Write-ScriptMessage "No matching PowerShellGet-installed modules were found." -Type Warning
        return @()
    }

    $grouped = @($installed | Group-Object -Property Name | Sort-Object Name)
    Write-ScriptMessage "[INFO] Modules discovered: $($grouped.Count)"

    $moduleCounter = 0
    foreach ($group in $grouped) {
        $moduleCounter++
        Write-Progress -Activity "Inspecting PowerShell modules" -Status "Processing $($group.Name)" -PercentComplete (($moduleCounter / $grouped.Count) * 100)

        $sorted = @($group.Group | Sort-Object { ConvertTo-SortableVersion -Value $_.Version } -Descending)
        $latest = $sorted[0]
        $oldVersions = @()
        if ($sorted.Count -gt 1) {
            $oldVersions = @($sorted[1..($sorted.Count - 1)])
        }

        Write-ScriptMessage "`n[INFO] $($group.Name) - $($sorted.Count) version(s), latest $($latest.Version)"

        if ($oldVersions.Count -eq 0) {
            if ($ShowAll) {
                $results.Add([PSCustomObject]@{
                    ModuleName        = $group.Name
                    Version           = [string]$latest.Version
                    LatestVersion     = [string]$latest.Version
                    Status            = 'Current'
                    Repository        = Get-ObjectPropertyValue -InputObject $latest -Name 'Repository'
                    InstalledLocation = Get-ObjectPropertyValue -InputObject $latest -Name 'InstalledLocation'
                    PublishedDate     = Get-ObjectPropertyValue -InputObject $latest -Name 'PublishedDate'
                })
            }
            else {
                Write-Host "  No older versions"
            }

            continue
        }

        foreach ($oldModule in $oldVersions) {
            Write-Host "  - Old version: $($oldModule.Version)"
            $results.Add([PSCustomObject]@{
                ModuleName        = $group.Name
                Version           = [string]$oldModule.Version
                LatestVersion     = [string]$latest.Version
                Status            = 'Outdated'
                Repository        = Get-ObjectPropertyValue -InputObject $oldModule -Name 'Repository'
                InstalledLocation = Get-ObjectPropertyValue -InputObject $oldModule -Name 'InstalledLocation'
                PublishedDate     = Get-ObjectPropertyValue -InputObject $oldModule -Name 'PublishedDate'
            })
        }
    }

    Write-Progress -Activity "Inspecting PowerShell modules" -Completed
    return $results.ToArray()
}

$transcriptStarted = $false

try {
    if ($CreateLog) {
        if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
            $LogDirectory = (Get-Location).Path
        }

        if (-not (Test-Path -LiteralPath $LogDirectory -PathType Container)) {
            throw "Log directory '$LogDirectory' does not exist or is not a directory"
        }

        $transcriptFile = Join-Path $LogDirectory ("psmodules-get-{0:yyyyMMdd-HHmmss}.txt" -f (Get-Date))
        Start-Transcript -Path $transcriptFile | Out-Null
        $transcriptStarted = $true
        Write-ScriptMessage "Script activity logged at $transcriptFile" -Type Warning
    }

    Clear-Host
    Write-ScriptMessage "Script started - Report PowerShell module versions`n" -Type System
    Write-ScriptMessage "[INFO] Checking PowerShell version"
    $ps = $PSVersionTable.PSVersion
    Write-ScriptMessage "- Detected PowerShell version: $($ps.Major).$($ps.Minor)`n"

    if (-not (Test-PowerShellGetAvailable)) {
        throw "Get-InstalledModule is not available. Install the PowerShellGet module and try again."
    }

    $moduleReport = @(Get-OldPSModules -ModuleName $ModuleName -ShowAll:$ShowAll)

    if ($moduleReport.Count -gt 0) {
        $moduleReport | Format-Table -AutoSize ModuleName, Version, LatestVersion, Status, Repository | Out-Host
        $outdatedCount = @($moduleReport | Where-Object { $_.Status -eq 'Outdated' }).Count
        $listedModuleCount = @($moduleReport | Select-Object -ExpandProperty ModuleName -Unique).Count
        Write-ScriptMessage "`nModules listed: $listedModuleCount"
        Write-ScriptMessage "Old module versions found: $outdatedCount"
    }
    else {
        Write-ScriptMessage "`nNo old module versions found."
    }

    if ($Csv) {
        if ([string]::IsNullOrWhiteSpace($OutputFile)) {
            $parentPath = Split-Path -Parent $PSScriptRoot
            if ([string]::IsNullOrWhiteSpace($parentPath)) {
                $parentPath = (Get-Location).Path
            }

            $OutputFile = Join-Path $parentPath ("psmodules-get-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))
        }

        $exportDir = Split-Path -Parent $OutputFile
        if (-not [string]::IsNullOrWhiteSpace($exportDir) -and -not (Test-Path -LiteralPath $exportDir)) {
            New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
        }

        $moduleReport | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
        Write-ScriptMessage "Results exported to $OutputFile"
    }

    if ($PassThru) {
        $moduleReport
    }

    Write-ScriptMessage "`nScript completed`n" -Type System
}
catch {
    Write-ScriptMessage "Script failed: $($_.Exception.Message)" -Type Error
    exit 1
}
finally {
    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }
}