<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/psmodules-get.ps1

Description - List versions of PowerShell modules installed

Prerequisites - PowerShell 5.1 or later

Get-OldPSModules -ModuleNames "Microsoft.Graph", "Microsoft.Online.SharePoint.PowerShell"

More scripts available by joining http://www.ciaopspatron.com

#>

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"

function Get-OldPSModules {
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$ModuleNames = @()
    )

    # If no modules specified, get all installed modules
    if ($ModuleNames.Count -eq 0) {
        write-host -ForegroundColor $processmessagecolor "[INFO] Getting all installed modules"
        $ModuleNames = Get-InstalledModule | Select-Object -ExpandProperty Name
        write-host -foregroundcolor $processmessagecolor "Number installed modules is $($ModuleNames.Count)"
    }

    # Create an array to store results
    $OldModuleVersions = @()

    foreach ($ModuleName in $ModuleNames) {
        write-host -ForegroundColor $processmessagecolor "`n[INFO] Processing module $ModuleName"
        try {
            # Get all installed versions of the module
            $InstalledModules = Get-InstalledModule -Name $ModuleName -AllVersions
            write-host -ForegroundColor $processmessagecolor "Number of installed versions is $($InstalledModules.Count)"

            # If more than one version exists
            if ($InstalledModules.Count -gt 1) {
                # Sort modules by version in descending order
                $SortedModules = $InstalledModules | Sort-Object Version -Descending

                # Identify old versions (all except the first/latest)
                $OldVersions = $SortedModules[1..($SortedModules.Count - 1)]

                # Add to results
                foreach ($OldModule in $OldVersions) {
                    write-host "- Old version found: $($OldModule.Version)"
                    $OldModuleVersions += [PSCustomObject]@{
                        ModuleName    = $ModuleName
                        Version       = $OldModule.Version
                        LatestVersion = $SortedModules[0].Version
                    }
                }
            }
        }
        catch {
            Write-Host -ForegroundColor $errormessagecolor "`nError processing $ModuleName : $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Display results
    if ($OldModuleVersions.Count -gt 0) {
        $OldModuleVersions | Format-Table -AutoSize
        Write-Host -ForegroundColor $processmessagecolor "`nTotal old module versions found: $($OldModuleVersions.Count)" -ForegroundColor Yellow
    }
    else {
        Write-Host -ForegroundColor $processmessagecolor "No old module versions found." -ForegroundColor Green
    }
}

# Main script
Clear-Host
write-host -foregroundcolor $systemmessagecolor "Script started - Report PowerShell Module versions`n"

write-host -foregroundcolor $processmessagecolor "[Info] = Checking PowerShell version"
$ps = $PSVersionTable.PSVersion
Write-host -foregroundcolor $processmessagecolor "- Detected supported PowerShell version: $($ps.Major).$($ps.Minor)`n"

Get-OldPSModules 

Write-Host "`nScript completed`n" -ForegroundColor $systemmessagecolor