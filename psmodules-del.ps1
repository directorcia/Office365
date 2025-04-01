<# CIAOPS
Script provided as is. Use at own risk. No guarantees or warranty provided.

Source - https://github.com/directorcia/Office365/blob/master/psmodules-del.ps1

Description - Remove old versions of PowerShell modules

Prerequisites - PowerShell 5.1 or later

Get-OldPSModules -ModuleNames "Microsoft.Graph", "Microsoft.Online.SharePoint.PowerShell"

More scripts available by joining http://www.ciaopspatron.com

#>

function Remove-OldPSModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string[]]$ModuleNames = @(),

        [Parameter(Mandatory=$false)]
        [switch]$ConfirmUninstall,

        [Parameter(Mandatory=$false)]
        [switch]$WhatIf,
        
        [Parameter(Mandatory=$false)]
        [switch]$SummaryOnly
    )

    # Create summary tracking variables
    $script:TotalModulesProcessed = 0
    $script:TotalVersionsRemoved = 0
    $script:TotalVersionsSkipped = 0
    $script:TotalErrors = 0
    $script:ModuleSummary = @()

    # Display script header
    Write-Host "`n========================================================" -ForegroundColor Cyan
    Write-Host "  PowerShell Module Version Cleanup Tool" -ForegroundColor Cyan
    Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan 
    Write-Host "========================================================`n" -ForegroundColor Cyan

    # If WhatIf is specified, notify the user
    if ($WhatIf) {
        Write-Host "RUNNING IN SIMULATION MODE (WhatIf). No modules will be removed." -ForegroundColor Yellow
        Write-Host "Use without -WhatIf to perform actual removal.`n" -ForegroundColor Yellow
    }

    # If no modules specified, get all installed module names
    if (-not $ModuleNames) {
        Write-Host "No specific modules specified. Scanning all installed modules..." -ForegroundColor Cyan
        $AllModules = Get-InstalledModule -ErrorAction SilentlyContinue
        if ($null -eq $AllModules) {
            Write-Host "No modules found installed via PowerShellGet." -ForegroundColor Red
            return
        }
        $ModuleNames = $AllModules | Select-Object -ExpandProperty Name | Get-Unique
        Write-Host "Found $($ModuleNames.Count) installed modules:" -ForegroundColor Green
        
        # List all modules one per line
        Write-Host "`nINSTALLED MODULES:" -ForegroundColor Cyan
        foreach ($module in $ModuleNames) {
            Write-Host "  - $module" -ForegroundColor Gray
        }
        Write-Host ""
    }
    else {
        Write-Host "Processing specified modules:" -ForegroundColor Cyan
        foreach ($module in $ModuleNames) {
            Write-Host "  - $module" -ForegroundColor Gray
        }
        Write-Host ""
    }

    # Create a progress bar for module processing
    $progressParams = @{
        Activity = "Processing PowerShell Modules"
        Status = "Starting module cleanup"
        PercentComplete = 0
    }
    Write-Progress @progressParams

    $moduleCounter = 0
    foreach ($ModuleName in $ModuleNames) {
        $moduleCounter++
        $progressParams.PercentComplete = ($moduleCounter / $ModuleNames.Count) * 100
        $progressParams.Status = "Processing module $moduleCounter of $($ModuleNames.Count): $ModuleName"
        Write-Progress @progressParams

        $moduleSummary = [PSCustomObject]@{
            Name = $ModuleName
            LatestVersion = "N/A"
            RemovedVersions = 0
            SkippedVersions = 0
            Error = $null
        }

        try {
            if (-not $SummaryOnly) {
                Write-Host "`n== Processing module: $ModuleName ==" -ForegroundColor Magenta
            }
            
            # Get all installed versions of the module, ordered by version descending
            $InstalledModules = Get-InstalledModule -Name $ModuleName -AllVersions -ErrorAction Stop | 
                                Sort-Object Version -Descending

            # If more than one version exists
            if ($InstalledModules.Count -gt 1) {
                # The first module in the sorted list is the latest version
                $LatestVersion = $InstalledModules[0]
                $moduleSummary.LatestVersion = $LatestVersion.Version
                
                if (-not $SummaryOnly) {
                    Write-Host "Found $($InstalledModules.Count) versions of '$ModuleName':" -ForegroundColor White
                    Write-Host "  - Latest: v$($LatestVersion.Version) (will be kept)" -ForegroundColor Green
                    Write-Host "  - Older versions:" -ForegroundColor Yellow
                    for ($i = 1; $i -lt $InstalledModules.Count; $i++) {
                        Write-Host "    $(if($WhatIf){"[SIMULATION] "})[v$($InstalledModules[$i].Version)] Published: $($InstalledModules[$i].PublishedDate)" -ForegroundColor Yellow
                    }
                }

                # Iterate through the older versions (starting from the second element)
                for ($i = 1; $i -lt $InstalledModules.Count; $i++) {
                    $OlderVersion = $InstalledModules[$i]
                    
                    if ($ConfirmUninstall) {
                        if (-not $SummaryOnly) {
                            Write-Host "`nOlder version: v$($OlderVersion.Version)" -NoNewline
                            Write-Host " [Published: $($OlderVersion.PublishedDate)]" -ForegroundColor Gray
                        }
                        
                        $ConfirmationResult = Read-Host "Confirm removal? (Y/N)"
                        if ($ConfirmationResult -ceq "Y") {
                            Uninstall-Module -Name $ModuleName -RequiredVersion $OlderVersion.Version -Force -WhatIf:$WhatIf
                            if (-not $WhatIf) {
                                $script:TotalVersionsRemoved++
                                $moduleSummary.RemovedVersions++
                                if (-not $SummaryOnly) {
                                    Write-Host "  ✓ Removed v$($OlderVersion.Version)" -ForegroundColor Green
                                }
                            }
                            else {
                                if (-not $SummaryOnly) {
                                    Write-Host "  [SIMULATION] Would remove v$($OlderVersion.Version)" -ForegroundColor Cyan
                                }
                            }
                        } 
                        else {
                            $script:TotalVersionsSkipped++
                            $moduleSummary.SkippedVersions++
                            if (-not $SummaryOnly) {
                                Write-Host "  ⊘ Skipped removal of v$($OlderVersion.Version)" -ForegroundColor Yellow
                            }
                        }
                    } 
                    else {
                        # Automatic removal without confirmation
                        Uninstall-Module -Name $ModuleName -RequiredVersion $OlderVersion.Version -Force -WhatIf:$WhatIf
                        if (-not $WhatIf) {
                            $script:TotalVersionsRemoved++
                            $moduleSummary.RemovedVersions++
                            if (-not $SummaryOnly) {
                                Write-Host "  ✓ Removed v$($OlderVersion.Version)" -ForegroundColor Green
                            }
                        }
                        else {
                            if (-not $SummaryOnly) {
                                Write-Host "  [SIMULATION] Would remove v$($OlderVersion.Version)" -ForegroundColor Cyan
                            }
                        }
                    }
                }

                if (-not $SummaryOnly) {
                    Write-Host "`n✓ Kept latest version of $ModuleName : v$($LatestVersion.Version)" -ForegroundColor Green
                }
                $script:TotalModulesProcessed++
            }
            elseif ($InstalledModules.Count -eq 1) {
                $moduleSummary.LatestVersion = $InstalledModules[0].Version
                if (-not $SummaryOnly) {
                    Write-Host "Only one version found for '$ModuleName': v$($InstalledModules[0].Version)" -ForegroundColor Yellow
                    Write-Host "No cleanup needed." -ForegroundColor Yellow
                }
                $script:TotalModulesProcessed++
            }
            else {
                if (-not $SummaryOnly) {
                    Write-Host "No versions of '$ModuleName' found installed via PowerShellGet." -ForegroundColor Yellow
                }
            }
        }
        catch {
            $script:TotalErrors++
            $errorMessage = $_.Exception.Message
            $moduleSummary.Error = $errorMessage
            
            if (-not $SummaryOnly) {
                Write-Host "❌ Error processing module '$ModuleName': $errorMessage" -ForegroundColor Red
            }
        }
        
        # Add to summary collection
        $script:ModuleSummary += $moduleSummary
    }

    # Complete the progress bar
    Write-Progress -Activity "Processing PowerShell Modules" -Completed

    # Display summary report
    Write-Host "`n========================================================" -ForegroundColor Cyan
    Write-Host "  SUMMARY REPORT" -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "Total modules processed: $script:TotalModulesProcessed" -ForegroundColor White
    Write-Host "Total versions removed: $(if($WhatIf){"[SIMULATION] "})$script:TotalVersionsRemoved" -ForegroundColor $(if($script:TotalVersionsRemoved -gt 0){"Green"}else{"White"})
    Write-Host "Total versions skipped: $script:TotalVersionsSkipped" -ForegroundColor $(if($script:TotalVersionsSkipped -gt 0){"Yellow"}else{"White"})
    Write-Host "Total errors encountered: $script:TotalErrors" -ForegroundColor $(if($script:TotalErrors -gt 0){"Red"}else{"White"})
    
    if ($script:ModuleSummary.Count -gt 0) {
        Write-Host "`nDETAILS BY MODULE:" -ForegroundColor Cyan
        foreach ($summary in $script:ModuleSummary) {
            $statusColor = if ($summary.Error) { "Red" } elseif ($summary.RemovedVersions -gt 0) { "Green" } else { "White" }
            Write-Host "  - $($summary.Name): " -NoNewline
            Write-Host "Latest v$($summary.LatestVersion)" -ForegroundColor Green -NoNewline
            
            if ($summary.RemovedVersions -gt 0) {
                if ($WhatIf) {
                    Write-Host ", Would remove $($summary.RemovedVersions) older version(s)" -ForegroundColor Cyan -NoNewline
                } else {
                    Write-Host ", Removed $($summary.RemovedVersions) older version(s)" -ForegroundColor Green -NoNewline
                }
            }
            
            if ($summary.SkippedVersions -gt 0) {
                Write-Host ", Skipped $($summary.SkippedVersions) version(s)" -ForegroundColor Yellow -NoNewline
            }
            
            if ($summary.Error) {
                Write-Host ", ERROR: $($summary.Error)" -ForegroundColor Red -NoNewline
            }
            
            Write-Host ""
        }
    }

    Write-Host "`nCleanup operation completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
    if ($WhatIf) {
        Write-Host "NOTE: This was a simulation run. Use without -WhatIf to perform actual removal." -ForegroundColor Yellow
    }
    Write-Host "========================================================`n" -ForegroundColor Cyan
}


Clear-Host

# Example usage:
# Remove-OldPSModules                                                  # Process all modules without confirmation
# Remove-OldPSModules -SummaryOnly                                     # Only show summary info, minimize verbose output
# Remove-OldPSModules -ModuleNames "Microsoft.Graph", "Az"            # Process specific modules
# Remove-OldPSModules -ConfirmUninstall                               # Ask before each removal
# Remove-OldPSModules -WhatIf                                         # Simulation mode, no actual changes
# Remove-OldPSModules -ModuleNames "Az" -ConfirmUninstall -WhatIf     # Combine parameters as needed

# Run the script with your chosen parameters
Remove-OldPSModules -ConfirmUninstall