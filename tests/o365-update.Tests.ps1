#requires -Version 7.0
<#
.SYNOPSIS
    Offline Pester tests for o365-update.ps1 helper functions.
.DESCRIPTION
    Function definitions are extracted from the script's AST so the main body (which checks
    elevation, talks to the gallery and calls exit) never runs. Cmdlets that would touch the
    system (Install-Module, Uninstall-Module, Get-InstalledModule, ...) are replaced by local
    stub functions, so nothing is installed or removed.

    Assertions use plain throw statements so the file runs on Pester 3.4 and Pester 5.
.EXAMPLE
    Invoke-Pester -Script ./tests/o365-update.Tests.ps1
#>

Describe 'o365-update.ps1' {
    BeforeAll {
        $scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'o365-update.ps1'
        $tokens = $null
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
        if ($parseErrors.Count) { throw ($parseErrors | Out-String) }

        $functionAsts = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
        }, $false)
        foreach ($definition in $functionAsts) {
            . ([scriptblock]::Create($definition.Extent.Text))
        }

        $Script:Colors = @{ System = 'Cyan'; Process = 'Green'; Warning = 'Yellow'; Error = 'Red'; Info = 'White'; Success = 'Green' }
        $Script:BootstrapModuleNames = @('PowerShellGet', 'PackageManagement')
        $Script:LastInstallationError = $null
        $Repository = 'PSGallery'

        # Silence UI output from the script's helpers during tests
        function Write-ColorOutput { param($Message, $Type, [switch]$NoNewline) }

        function Assert-Equal($Actual, $Expected, [string]$Because) {
            if ($Actual -ne $Expected) { throw "$Because. Expected '$Expected' but got '$Actual'." }
        }
    }

    Context 'ConvertTo-ModuleVersion' {
        It 'strips a prerelease label' {
            Assert-Equal (ConvertTo-ModuleVersion '3.1.0-nightly') ([version]'3.1.0') 'prerelease suffix'
        }
        It 'strips build metadata' {
            Assert-Equal (ConvertTo-ModuleVersion '1.2.3-beta+abc') ([version]'1.2.3') 'build metadata'
        }
        It 'accepts a single-component version' {
            Assert-Equal (ConvertTo-ModuleVersion '5') ([version]'5.0') 'single integer'
        }
        It 'passes a [version] straight through and maps $null to 0.0' {
            Assert-Equal (ConvertTo-ModuleVersion ([version]'2.0.1')) ([version]'2.0.1') 'version object'
            Assert-Equal (ConvertTo-ModuleVersion $null) ([version]'0.0') 'null'
        }
        It 'throws on text that is not a version' {
            $threw = $false
            try { $null = ConvertTo-ModuleVersion 'latest' } catch { $threw = $true }
            if (-not $threw) { throw 'Expected an exception for a non-version string' }
        }
    }

    Context 'Get-FilteredModuleList' {
        BeforeAll {
            $Script:ModuleList = @(
                @{ Name = 'ExchangeOnlineManagement'; Category = 'Exchange'; Priority = 2 },
                @{ Name = 'Microsoft.Graph'; Category = 'Graph'; Priority = 1 },
                @{ Name = 'Microsoft.Graph.Beta'; Category = 'Graph'; Priority = 3 }
            )
        }

        It 'yields one module for a single-module scope (previously reported the hashtable key count)' {
            $result = @(Get-FilteredModuleList -Scope 'Exchange')
            Assert-Equal $result.Count 1 'module count'
            Assert-Equal $result[0].Name 'ExchangeOnlineManagement' 'module name'
        }
        It 'de-duplicates and orders by priority across scopes' {
            $result = @(Get-FilteredModuleList -Scope 'Graph', 'Exchange', 'Graph')
            Assert-Equal $result.Count 3 'module count'
            Assert-Equal ($result.Name -join ',') 'Microsoft.Graph,ExchangeOnlineManagement,Microsoft.Graph.Beta' 'order'
        }
        It 'returns every module for All' {
            Assert-Equal @(Get-FilteredModuleList -Scope 'All').Count 3 'module count'
        }
    }

    Context 'Test-ModuleConflictError' {
        It 'detects command clobber errors' {
            if (-not (Test-ModuleConflictError -ModuleName 'Az' -ErrorMessage 'The following commands are already available on this system')) { throw 'not detected' }
        }
        It 'detects publisher errors only for PnP.PowerShell' {
            if (-not (Test-ModuleConflictError -ModuleName 'PnP.PowerShell' -ErrorMessage 'publisher mismatch')) { throw 'PnP not detected' }
            if (Test-ModuleConflictError -ModuleName 'Az' -ErrorMessage 'publisher mismatch') { throw 'false positive for Az' }
        }
        It 'ignores unrelated errors' {
            if (Test-ModuleConflictError -ModuleName 'Az' -ErrorMessage 'Unable to resolve package source') { throw 'false positive' }
        }
    }

    Context 'Test-AllUsersModuleWriteAccess' {
        It 'creates and removes a probe in a writable target directory' {
            $modulePath = Join-Path $env:TEMP ([guid]::NewGuid().ToString('N'))
            $null = New-Item -Path $modulePath -ItemType Directory
            try {
                $result = Test-AllUsersModuleWriteAccess -ModulePath $modulePath
                Assert-Equal $result.CanWrite $true 'write access'
                Assert-Equal @(Get-ChildItem -LiteralPath $modulePath).Count 0 'temporary probe cleanup'
            }
            finally {
                Remove-Item -LiteralPath $modulePath -Recurse -Force
            }
        }
        It 'fails closed when the target directory is missing' {
            $result = Test-AllUsersModuleWriteAccess -ModulePath (Join-Path $env:TEMP ([guid]::NewGuid().ToString('N')))
            Assert-Equal $result.CanWrite $false 'write access'
            if ($result.ErrorMessage -notmatch 'does not exist') { throw 'Expected an actionable missing-directory error' }
        }
        It 'does not create a probe under -WhatIf' {
            $modulePath = Join-Path $env:TEMP ([guid]::NewGuid().ToString('N'))
            $null = New-Item -Path $modulePath -ItemType Directory
            try {
                $result = Test-AllUsersModuleWriteAccess -ModulePath $modulePath -WhatIf
                Assert-Equal $result.CanWrite $null 'write check result'
                Assert-Equal @(Get-ChildItem -LiteralPath $modulePath).Count 0 'probe count'
            }
            finally {
                Remove-Item -LiteralPath $modulePath -Recurse -Force
            }
        }
    }

    Context 'Get-ModuleInstallWorkerScript' {
        It 'reads a request file and installs with its module name and parameters' {
            $command = Get-ModuleInstallWorkerScript
            if ($command -notmatch 'Get-Content -LiteralPath \$RequestPath -Raw') { throw 'Expected the worker to read its request file' }
            if ($command -notmatch 'Install-Module -Name \$request\.ModuleName @workerParams') { throw 'Expected the worker to splat the serialized parameters' }
        }
    }

    Context 'Invoke-PowerShellCommandWithTimeout' {
        It 'captures a successful child process result' {
            $command = "param([string]`$RequestPath)`n'worker-ok' | Write-Output; exit 0"
            $result = Invoke-PowerShellCommandWithTimeout -CommandText $command -TimeoutSeconds 15 -Activity 'Worker test'
            Assert-Equal $result.Success $true 'worker success'
            if ($result.StandardOutput -notmatch 'worker-ok') { throw 'Expected child process output to be captured' }
        }
        It 'passes its JSON request to the child process' {
            $command = @'
param([string]$RequestPath)
$request = Get-Content -LiteralPath $RequestPath -Raw | ConvertFrom-Json -AsHashtable
if ($request.ModuleName -eq 'PnP.PowerShell' -and $request.InstallParams.Scope -eq 'AllUsers') { exit 0 }
exit 1
'@
            $requestJson = @{ ModuleName = 'PnP.PowerShell'; InstallParams = @{ Scope = 'AllUsers' } } | ConvertTo-Json -Compress
            $result = Invoke-PowerShellCommandWithTimeout -CommandText $command -RequestJson $requestJson -TimeoutSeconds 15 -Activity 'Request test'
            Assert-Equal $result.Success $true 'request transfer'
        }
        It 'terminates a child process that exceeds its timeout' {
            $command = "param([string]`$RequestPath)`nStart-Sleep -Seconds 30"
            $result = Invoke-PowerShellCommandWithTimeout -CommandText $command -TimeoutSeconds 1 -Activity 'Timeout test'
            Assert-Equal $result.TimedOut $true 'timeout status'
            Assert-Equal $result.Success $false 'worker success'
            if ($result.ErrorMessage -notmatch 'timeout') { throw 'Expected a timeout error message' }
        }
    }

    Context 'Invoke-ModuleInstallDirect' {
        It 'retries with -AllowClobber on a conflict without mutating the caller''s parameters' {
            $Script:workerRequests = [System.Collections.Generic.List[string]]::new()
            Mock Invoke-PowerShellCommandWithTimeout {
                $Script:workerRequests.Add($RequestJson)
                if ($Script:workerRequests.Count -eq 1) {
                    [pscustomobject]@{ Success = $false; TimedOut = $false; ErrorMessage = 'The following commands are already available on this system' }
                } else {
                    [pscustomobject]@{ Success = $true; TimedOut = $false; ErrorMessage = $null }
                }
            }

            $params = @{ Force = $true; Scope = 'AllUsers' }
            $ok = Invoke-ModuleInstallDirect -ModuleName 'PnP.PowerShell' -InstallParams $params -Operation 'Install' -StartTime (Get-Date) -TimeoutMinutes 1

            Assert-Equal $ok $true 'result'
            Assert-Equal $Script:workerRequests.Count 2 'worker process calls'
            $retryRequest = ConvertFrom-Json -InputObject $Script:workerRequests[1] -AsHashtable
            Assert-Equal $retryRequest.InstallParams.AllowClobber $true 'retry AllowClobber'
            Assert-Equal $retryRequest.InstallParams.SkipPublisherCheck $true 'retry SkipPublisherCheck for PnP'
            if ($params.ContainsKey('AllowClobber')) { throw 'Caller hashtable was mutated' }
        }
        It 'returns $false and records the error for a non-conflict failure' {
            Mock Invoke-PowerShellCommandWithTimeout {
                [pscustomobject]@{ Success = $false; TimedOut = $false; ErrorMessage = 'Network unreachable' }
            }
            $ok = Invoke-ModuleInstallDirect -ModuleName 'Az' -InstallParams @{ Force = $true } -Operation 'Update' -StartTime (Get-Date) -TimeoutMinutes 1
            Assert-Equal $ok $false 'result'
            Assert-Equal $Script:LastInstallationError 'Network unreachable' 'recorded error'
        }
    }

    Context 'Test-ModulePresentOnDisk' {
        It 'caches each name, matches case-insensitively and rescans on -Refresh' {
            $Script:ModulePresenceCache = @{}
            $Script:scanCount = 0
            function Get-Module { [CmdletBinding()] param([string]$Name, [switch]$ListAvailable) $Script:scanCount++; if ($Name -ieq 'MicrosoftTeams') { [pscustomobject]@{ Name = $Name } } }

            if (-not (Test-ModulePresentOnDisk -Name 'microsoftteams')) { throw 'Expected present' }
            if (Test-ModulePresentOnDisk -Name 'AzureAD') { throw 'Expected absent' }
            Assert-Equal $Script:scanCount 2 'scans before refresh'
            $null = Test-ModulePresentOnDisk -Name 'AzureAD' -Refresh
            Assert-Equal $Script:scanCount 3 'scans after refresh'
        }
    }

    Context 'Get-ModuleVersionInfo' {
        It 'skips Get-InstalledModule for a module that is not on disk and recommends install' {
            $Script:getInstalledCalls = 0
            function Test-ModulePresentOnDisk { [CmdletBinding()] param($Name, [switch]$Refresh) $false }
            function Get-InstalledModule { [CmdletBinding()] param($Name, [switch]$AllVersions) $Script:getInstalledCalls++ }
            function Find-Module { [CmdletBinding()] param($Name, $Repository) [pscustomobject]@{ Name = $Name; Version = '2.0.0' } }

            $info = Get-ModuleVersionInfo -ModuleName 'MicrosoftTeams'
            Assert-Equal $Script:getInstalledCalls 0 'Get-InstalledModule calls'
            Assert-Equal $info.InstallRecommended $true 'InstallRecommended'
            Assert-Equal $info.ErrorMessage $null 'ErrorMessage'
        }
        It 'compares prerelease versions without failing' {
            function Test-ModulePresentOnDisk { [CmdletBinding()] param($Name, [switch]$Refresh) $true }
            function Get-InstalledModule { [CmdletBinding()] param($Name, [switch]$AllVersions) [pscustomobject]@{ Name = $Name; Version = '2.0.0-preview' } }
            function Find-Module { [CmdletBinding()] param($Name, $Repository) [pscustomobject]@{ Name = $Name; Version = '2.1.0' } }

            $info = Get-ModuleVersionInfo -ModuleName 'PnP.PowerShell'
            Assert-Equal $info.ErrorMessage $null 'ErrorMessage'
            Assert-Equal $info.UpdateAvailable $true 'UpdateAvailable'
        }
    }

    Context 'Remove-OldModuleVersions' {
        BeforeAll {
            function Get-ModuleVersionCleanupEntries {
                [CmdletBinding()] param($ModuleName, [switch]$IncludeAvailable)
                foreach ($v in '3.0.0', '2.0.0', '1.0.0') {
                    [pscustomobject]@{ Version = [version]$v; ModuleBase = "C:\Modules\$ModuleName\$v"; ModuleRoot = "C:\Modules\$ModuleName" }
                }
            }
            function Uninstall-Module { [CmdletBinding(SupportsShouldProcess)] param($Name, $RequiredVersion, [switch]$Force) $Script:uninstalled.Add([string]$RequiredVersion) }
        }

        It 'removes all but the newest version' {
            $Script:uninstalled = [System.Collections.Generic.List[string]]::new()
            $result = Remove-OldModuleVersions -ModuleName 'Az' -KeepLatestOnly
            Assert-Equal ($Script:uninstalled -join ',') '2.0.0,1.0.0' 'removed versions'
            Assert-Equal ($result.KeptVersions -join ',') '3.0.0' 'kept versions'
            Assert-Equal $result.Success $true 'Success'
        }
        It 'removes nothing under -WhatIf' {
            $Script:uninstalled = [System.Collections.Generic.List[string]]::new()
            $result = Remove-OldModuleVersions -ModuleName 'Az' -KeepLatestOnly -WhatIf
            Assert-Equal $Script:uninstalled.Count 0 'Uninstall-Module calls'
            Assert-Equal $result.RemovedCount 0 'RemovedCount'
        }
    }
}
