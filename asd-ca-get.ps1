<#
.SYNOPSIS
    Evaluate Conditional Access policies against ASD Blueprint recommendations

.DESCRIPTION
    This script evaluates the current Conditional Access policies in an Entra ID tenant
    against the recommendations from the ASD (Australian Signals Directorate) Blueprint
    for Secure Cloud and generates an HTML report showing compliance gaps and improvements.
    
    Reference: https://blueprint.asd.gov.au/configuration/entra-id/protection/conditional-access/

.EXAMPLE
    .\asd-ca-get.ps1
    
    Connects to Microsoft Graph, evaluates all CA policies, and generates an HTML report

.NOTES
    Author: CIAOPS
    Date: 2025-11-25
    Version: 1.0
    Requires: Microsoft.Graph.Authentication PowerShell module (auto-installed if missing)
    
    COMPLIANCE THRESHOLDS:
    This script uses the following thresholds for compliance status determination:
    - Compliant:     ≥80% (policy meets most/all requirements with minor gaps)
    - Partial:       50-79% (policy meets core requirements but has significant gaps)
    - Non-Compliant: <50% (policy fails to meet fundamental requirements)
    
    RATIONALE:
    These thresholds are balanced to:
    1. Allow reasonable implementation flexibility while maintaining security posture
    2. Avoid false negatives (marking partial implementations as non-compliant)
    3. Identify policies that are "close enough" for remediation vs complete rebuild
    4. Account for organizational constraints (exclusions, phased rollouts)
    
    NOTE: The ASD Blueprint does not prescribe specific compliance thresholds.
    Organizations may adjust these values in lines 1344-1346 based on their risk
    tolerance and compliance requirements. Some frameworks use 90% for "compliant";
    however, this can result in excessive false negatives in real-world deployments
    where minor exclusions or conditional variations are legitimate.
    
.LINK
    Code - https://github.com/directorcia/office365-tools/tree/main/scripts/asd-ca-get.ps1
    Documentation - https://github.com/directorcia/Office365/wiki/ASD-Conditional-Access-Policy-Evaluation-Script
#>

[CmdletBinding()]
param()

# Check for required Microsoft.Graph modules
$requiredModules = @('Microsoft.Graph.Authentication')
$missingModules = @()

Write-Host "`nChecking for required PowerShell modules..." -ForegroundColor Cyan

foreach ($moduleName in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        $missingModules += $moduleName
        Write-Host "  ✗ $moduleName not found" -ForegroundColor Yellow
    }
    else {
        Write-Host "  ✓ $moduleName found" -ForegroundColor Green
    }
}

# Install missing modules
if ($missingModules.Count -gt 0) {
    Write-Host "`nInstalling missing modules..." -ForegroundColor Yellow
    foreach ($moduleName in $missingModules) {
        try {
            Write-Host "  Installing $moduleName..." -ForegroundColor Cyan
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Host "  ✓ Successfully installed $moduleName" -ForegroundColor Green
        }
        catch {
            Write-Host "  ✗ Failed to install $moduleName. Error: $_" -ForegroundColor Red
            Write-Host "`nPlease install the module manually using:" -ForegroundColor Yellow
            Write-Host "  Install-Module -Name $moduleName -Scope CurrentUser -Force" -ForegroundColor White
            exit 1
        }
    }
}
else {
    Write-Host "All required modules are installed" -ForegroundColor Green
}

# Connect to Microsoft Graph
Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan

# Check if already connected
$context = Get-MgContext -ErrorAction SilentlyContinue
if ($context -and 
    $context.Scopes -contains "Policy.Read.All" -and 
    $context.Scopes -contains "Directory.Read.All") {
    Write-Host "Already connected to Microsoft Graph" -ForegroundColor Green
    Write-Host "Account: $($context.Account)" -ForegroundColor Gray
}
else {
    try {
        # Try interactive browser authentication first
        # Note: No Out-Default needed - browser auth produces no console output
        Connect-MgGraph -Scopes "Policy.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
        Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
    }
    catch {
        # Check if it's a localhost binding / HTTP listener issue
        # This can be an HttpListenerException or contain related keywords in the error chain
        $isHttpListenerIssue = $false
        $currentException = $_.Exception
        
        # Walk the exception chain to find HttpListenerException or related errors
        while ($currentException) {
            if ($currentException -is [System.Net.HttpListenerException]) {
                $isHttpListenerIssue = $true
                break
            }
            if ($currentException.Message -match "HttpListener|localhost.*listen|unable to listen|127\.0\.0\.1.*listen") {
                $isHttpListenerIssue = $true
                break
            }
            $currentException = $currentException.InnerException
        }
        
        if ($isHttpListenerIssue) {
            Write-Host "`nInteractive browser authentication failed (localhost binding issue)." -ForegroundColor Yellow
            Write-Host "Switching to device code authentication..." -ForegroundColor Cyan
            Write-Host "`n============================================================" -ForegroundColor Cyan
            Write-Host "  DEVICE CODE AUTHENTICATION REQUIRED" -ForegroundColor Cyan
            Write-Host "============================================================" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "Opening browser to: " -NoNewline
            Write-Host "https://microsoft.com/devicelogin" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "The device code will appear below - copy it and paste into the browser." -ForegroundColor Cyan
            Write-Host "============================================================`n" -ForegroundColor Cyan
            
            # Open the device login page in default browser
            try {
                Start-Process "https://microsoft.com/devicelogin"
                Start-Sleep -Seconds 2  # Give browser time to open
            } catch {
                Write-Host "Could not open browser automatically." -ForegroundColor Yellow
                Write-Host "If on a headless system, manually navigate to https://microsoft.com/devicelogin" -ForegroundColor Cyan
            }
            
            # Connect with device code - output goes directly to console
            # Note: Out-Default forces immediate display of device code and instructions
            Write-Host ""
            Connect-MgGraph -Scopes "Policy.Read.All", "Directory.Read.All" -UseDeviceAuthentication -NoWelcome -ErrorAction Stop | Out-Default
            Write-Host ""
            Write-Host "Successfully connected to Microsoft Graph via device code." -ForegroundColor Green
        }
        else {
            throw
        }
    }
}

# Get tenant information
Write-Host "`nRetrieving tenant information..." -ForegroundColor Cyan
try {
    $orgResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -OutputType PSObject
    $tenantDetails = $orgResponse.value[0]
    $tenantName = $tenantDetails.displayName
    $tenantId = $tenantDetails.id
    $tenantDomain = $tenantDetails.verifiedDomains | Where-Object { $_.isDefault -eq $true } | Select-Object -ExpandProperty name
    if (-not $tenantDomain) {
        $tenantDomain = $tenantDetails.verifiedDomains[0].name
    }
    Write-Host "Tenant: $tenantName" -ForegroundColor Green
}
catch {
    Write-Host "Failed to retrieve tenant information. Error: $_" -ForegroundColor Red
    $tenantName = "Unknown"
    $tenantId = "Unknown"
    $tenantDomain = "Unknown"
}

Write-Host "`nEvaluating Conditional Access policies for tenant: $tenantName" -ForegroundColor Cyan

# Get all Conditional Access policies
try {
    Write-Host "Retrieving Conditional Access policies..." -ForegroundColor Cyan
    $policiesResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -OutputType PSObject
    $policies = $policiesResponse.value
    Write-Host "Found $($policies.Count) Conditional Access policies" -ForegroundColor Green
}
catch {
    Write-Host "Failed to retrieve Conditional Access policies. Error: $_" -ForegroundColor Red
    exit 1
}

# Check for available authentication contexts
Write-Host "Checking for authentication contexts..." -ForegroundColor Cyan
$authContexts = @()
try {
    $authContextResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/authenticationContextClassReferences" -OutputType PSObject -ErrorAction SilentlyContinue
    $authContexts = $authContextResponse.value
    if ($authContexts.Count -gt 0) {
        Write-Host "Found $($authContexts.Count) authentication context(s)" -ForegroundColor Green
    }
    else {
        Write-Host "No authentication contexts configured" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Authentication contexts not available (requires appropriate licensing)" -ForegroundColor Yellow
    $authContexts = @()
}

# Check for Insider Risk Management availability
Write-Host "Checking for Insider Risk Management features..." -ForegroundColor Cyan
$insiderRiskAvailable = $false
try {
    # Try to access a policy with insider risk conditions to verify availability
    # Note: This is a best-effort check as the API might not expose this directly
    $hasInsiderRiskPolicy = $policies | Where-Object { 
        $null -ne $_.conditions.insiderRiskLevels 
    }
    if ($hasInsiderRiskPolicy) {
        $insiderRiskAvailable = $true
        Write-Host "Insider Risk Management integration detected" -ForegroundColor Green
    }
    else {
        Write-Host "No policies with Insider Risk conditions found" -ForegroundColor Yellow
        Write-Host "  Note: Insider Risk Management requires Microsoft Purview and additional licensing" -ForegroundColor Gray
    }
}
catch {
    Write-Host "Insider Risk Management features not available" -ForegroundColor Yellow
    $insiderRiskAvailable = $false
}

# Fetch authentication strength policies for validation
Write-Host "Retrieving authentication strength policies..." -ForegroundColor Cyan
$authStrengthPolicies = @{}
$authStrengthAvailable = $false
try {
    $authStrengthResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identity/conditionalAccess/authenticationStrength/policies" -OutputType PSObject -ErrorAction SilentlyContinue
    if ($authStrengthResponse.value) {
        foreach ($policy in $authStrengthResponse.value) {
            $authStrengthPolicies[$policy.id] = @{
                DisplayName = $policy.displayName
                Description = $policy.description
                PolicyType = $policy.policyType
                AllowedCombinations = $policy.allowedCombinations
            }
        }
        $authStrengthAvailable = $true
        Write-Host "Found $($authStrengthPolicies.Count) authentication strength policy/policies" -ForegroundColor Green
    }
    else {
        Write-Host "No authentication strength policies configured" -ForegroundColor Yellow
        Write-Host "  Note: Authentication strength requires Microsoft Entra ID P1 or P2 licensing" -ForegroundColor Gray
    }
}
catch {
    Write-Host "Unable to retrieve authentication strength policies (requires appropriate licensing)" -ForegroundColor Yellow
    $authStrengthPolicies = @{}
    $authStrengthAvailable = $false
}

# Check for Terms of Use availability
Write-Host "Checking for Terms of Use..." -ForegroundColor Cyan
$termsOfUseAvailable = $false
$termsOfUsePolicies = @()
try {
    $touResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identityGovernance/termsOfUse/agreements" -OutputType PSObject -ErrorAction SilentlyContinue
    if ($touResponse.value) {
        $termsOfUsePolicies = $touResponse.value
        $termsOfUseAvailable = $true
        Write-Host "Found $($termsOfUsePolicies.Count) Terms of Use agreement(s)" -ForegroundColor Green
    }
    else {
        Write-Host "No Terms of Use agreements configured" -ForegroundColor Yellow
        Write-Host "  Note: Terms of Use requires Microsoft Entra ID P1 or P2 licensing" -ForegroundColor Gray
    }
}
catch {
    Write-Host "Unable to retrieve Terms of Use agreements (requires appropriate licensing)" -ForegroundColor Yellow
    $termsOfUseAvailable = $false
    $termsOfUsePolicies = @()
}

# Define phishing-resistant MFA methods per Microsoft documentation
# Reference: https://learn.microsoft.com/en-us/entra/identity/authentication/concept-authentication-strengths
$phishingResistantMethods = @(
    "windowsHelloForBusiness",
    "fido2",
    "x509CertificateMultiFactor",
    "deviceBasedPush"  # Authenticator with GPS location on compliant device
)

# Helper function to check if authentication strength policy is phishing-resistant
function Test-IsPhishingResistantAuthStrength {
    param(
        [string]$AuthStrengthId,
        [hashtable]$AuthStrengthPolicies
    )
    
    if (-not $AuthStrengthPolicies.ContainsKey($AuthStrengthId)) {
        return @{
            IsPhishingResistant = $false
            Reason = "Authentication strength policy not found (ID: $AuthStrengthId)"
            Methods = @()
        }
    }
    
    $policy = $AuthStrengthPolicies[$AuthStrengthId]
    $allowedMethods = $policy.AllowedCombinations
    
    # Check if policy uses built-in phishing-resistant strength
    if ($policy.PolicyType -eq "builtIn" -and $policy.DisplayName -like "*phishing*resistant*") {
        return @{
            IsPhishingResistant = $true
            Reason = "Built-in phishing-resistant authentication strength"
            Methods = $allowedMethods
            PolicyName = $policy.DisplayName
        }
    }
    
    # For custom policies, check if all allowed methods are phishing-resistant
    if ($allowedMethods -and $allowedMethods.Count -gt 0) {
        $hasNonPhishingResistant = $false
        $nonPhishingResistantMethods = @()
        
        foreach ($method in $allowedMethods) {
            # Each combination can be a single method or comma-separated methods
            $methodsInCombination = $method -split ','
            foreach ($individualMethod in $methodsInCombination) {
                $individualMethod = $individualMethod.Trim()
                if ($individualMethod -notin $phishingResistantMethods) {
                    $hasNonPhishingResistant = $true
                    if ($individualMethod -notin $nonPhishingResistantMethods) {
                        $nonPhishingResistantMethods += $individualMethod
                    }
                }
            }
        }
        
        if (-not $hasNonPhishingResistant) {
            return @{
                IsPhishingResistant = $true
                Reason = "All authentication methods are phishing-resistant"
                Methods = $allowedMethods
                PolicyName = $policy.DisplayName
            }
        }
        else {
            return @{
                IsPhishingResistant = $false
                Reason = "Policy allows non-phishing-resistant methods: $($nonPhishingResistantMethods -join ', ')"
                Methods = $allowedMethods
                PolicyName = $policy.DisplayName
                NonPhishingResistantMethods = $nonPhishingResistantMethods
            }
        }
    }
    
    return @{
        IsPhishingResistant = $false
        Reason = "No authentication methods configured"
        Methods = @()
        PolicyName = $policy.DisplayName
    }
}

# Define ASD Blueprint recommended policies
$asdRecommendations = @(
    @{
        Category = "Admin Protection"
        Name = "ADM - S - Limit admin sessions"
        Description = "Limits administrative session durations"
        Key = "limit-admin-sessions"
        CheckCriteria = @{
            Users = "Directory roles (admin roles)"
            GrantControls = "Grant access"
            SessionControls = "Sign-in frequency configured"
        }
        Priority = "High"
    },
    @{
        Category = "Device Protection"
        Name = "DEV - B - Block access from unapproved devices"
        Description = "Blocks access from devices not approved"
        Key = "block-unapproved-devices"
        CheckCriteria = @{
            Users = "All users"
            GrantControls = "Block access OR Require compliant device"
            DevicePlatforms = "Specified platforms"
        }
        Priority = "High"
    },
    @{
        Category = "Device Protection"
        Name = "DEV - G - Compliant devices"
        Description = "Requires devices to be marked as compliant"
        Key = "compliant-devices"
        CheckCriteria = @{
            Users = "All users"
            Applications = "All resources"
            GrantControls = "Require device to be marked as compliant"
        }
        Priority = "High"
    },
    @{
        Category = "Device Protection"
        Name = "DEV - G - Intune enrolment with strong auth"
        Description = "Requires strong authentication for Intune enrollment"
        Key = "intune-enrolment"
        CheckCriteria = @{
            Applications = "Microsoft Intune Enrollment"
            GrantControls = "Require authentication strength (Phishing-resistant MFA)"
        }
        Priority = "Medium"
    },
    @{
        Category = "Guest Protection"
        Name = "GST - B - Block guests"
        Description = "Blocks guest access where not required"
        Key = "block-guests"
        CheckCriteria = @{
            Users = "Guest or external users"
            GrantControls = "Block access"
        }
        Priority = "Medium"
    },
    @{
        Category = "Guest Protection"
        Name = "GST - G - Guest application access with strong auth"
        Description = "Requires strong authentication for guest access to specific applications"
        Key = "guest-app-access"
        CheckCriteria = @{
            Users = "Guest or external users"
            GrantControls = "Require authentication strength"
        }
        Priority = "Medium"
    },
    @{
        Category = "Location Protection"
        Name = "LOC - B - Block access from unapproved countries"
        Description = "Blocks access from countries not on approved list"
        Key = "block-unapproved-countries"
        CheckCriteria = @{
            Users = "All users"
            Locations = "Configured with excluded locations"
            GrantControls = "Block access"
        }
        Priority = "High"
    },
    @{
        Category = "User Protection"
        Name = "USR - B - Block access to PROTECTED information"
        Description = "Blocks access to information classified as PROTECTED"
        Key = "block-protected-info"
        CheckCriteria = @{
            AuthenticationContexts = "PROTECTED information"
            GrantControls = "Block access"
        }
        Priority = "Medium"
    },
    @{
        Category = "User Protection"
        Name = "USR - B - Block access via legacy auth"
        Description = "Blocks legacy authentication protocols"
        Key = "block-legacy-auth"
        CheckCriteria = @{
            Users = "All users"
            ClientApps = "Exchange ActiveSync clients, Other clients"
            GrantControls = "Block access"
        }
        Priority = "Critical"
    },
    @{
        Category = "User Protection"
        Name = "USR - B - Block high-risk sign-ins"
        Description = "Blocks sign-ins identified as high risk"
        Key = "block-high-risk-signins"
        CheckCriteria = @{
            Users = "All users"
            SignInRiskLevels = "High"
            GrantControls = "Block access"
        }
        Priority = "Critical"
    },
    @{
        Category = "User Protection"
        Name = "USR - B - Block high-risk users"
        Description = "Blocks users identified as high risk"
        Key = "block-high-risk-users"
        CheckCriteria = @{
            Users = "All users"
            UserRiskLevels = "High"
            GrantControls = "Block access"
        }
        Priority = "Critical"
    },
    @{
        Category = "User Protection"
        Name = "USR - B - Block users with elevated insider risk"
        Description = "Blocks users with elevated insider risk scores"
        Key = "block-insider-risk"
        CheckCriteria = @{
            InsiderRisk = "Elevated"
            GrantControls = "Block access"
        }
        Priority = "High"
    },
    @{
        Category = "User Protection"
        Name = "USR - G - Agreement to terms of use"
        Description = "Requires agreement to terms of use"
        Key = "terms-of-use"
        CheckCriteria = @{
            Users = "All users"
            GrantControls = "Require terms of use"
        }
        Priority = "Low"
    },
    @{
        Category = "User Protection"
        Name = "USR - G - Register security info with strong auth"
        Description = "Requires strong authentication when registering security information"
        Key = "register-security-info"
        CheckCriteria = @{
            Users = "All users"
            UserActions = "Register security information"
            GrantControls = "Require authentication strength"
        }
        Priority = "High"
    },
    @{
        Category = "User Protection"
        Name = "USR - G - Require strong auth"
        Description = "Requires phishing-resistant MFA for all users"
        Key = "strong-auth"
        CheckCriteria = @{
            Users = "All users"
            Applications = "All resources"
            GrantControls = "Require authentication strength (Phishing-resistant MFA)"
        }
        Priority = "Critical"
    },
    @{
        Category = "User Protection"
        Name = "USR - G - Risky sign-ins with strong auth"
        Description = "Requires strong authentication for medium/high risk sign-ins"
        Key = "risky-signins-auth"
        CheckCriteria = @{
            Users = "All users"
            SignInRiskLevels = "Medium, High"
            GrantControls = "Require MFA"
        }
        Priority = "High"
    },
    @{
        Category = "User Protection"
        Name = "USR - S - Limit user sessions"
        Description = "Limits user session durations"
        Key = "limit-user-sessions"
        CheckCriteria = @{
            Users = "All users"
            SessionControls = "Sign-in frequency configured"
        }
        Priority = "Medium"
    }
)

# Function to check if a policy matches ASD recommendation
# Uses normalized weighted scoring where each criterion type (Users, Applications, GrantControls, etc.)
# receives equal weight (100 points) when present in the recommendation's CheckCriteria.
# This ensures all policies are compared on an equal 0-100% scale regardless of which criteria are checked.
#
# Critical Validation Features:
# - Policy State: Non-enabled policies (disabled/reportOnly) have scores reduced by 50%
# - Exclusions: Validates that inclusions aren't negated by exclusions (users, apps, roles)
# - Grant Controls Operator: Checks AND vs OR operator when multiple controls are configured
# - Authentication Strength: Reports auth strength policy IDs for manual verification
# - Detailed Findings: Provides actionable warnings for partial compliance scenarios
function Test-PolicyCompliance {
    param(
        [Parameter(Mandatory)]
        $Policy,
        [Parameter(Mandatory)]
        $Recommendation
    )
    
    $findings = @()
    $complianceScore = 0
    $maxScore = 0
    
    # Helper function to check for significant exclusions
    function Test-HasSignificantExclusions {
        param($Policy, $InclusionType)
        
        # Check if policy excludes all users or large groups
        if ($Policy.conditions.users.excludeUsers -contains "All") {
            return $true
        }
        
        # If including "All users", check for exclusions that might negate the policy
        if ($InclusionType -eq "All" -and 
            ($Policy.conditions.users.excludeUsers.Count -gt 0 -or
             $Policy.conditions.users.excludeGroups.Count -gt 5 -or
             $Policy.conditions.users.excludeRoles.Count -gt 3)) {
            return $true
        }
        
        return $false
    }
    
    # Helper function to validate grant controls operator
    function Test-GrantControlsOperator {
        param($Policy, $RequiredControls)
        
        $operator = $Policy.grantControls._Operator
        
        # If multiple controls are required, AND operator is typically needed
        if ($RequiredControls -gt 1 -and $operator -ne "AND") {
            return @{
                Valid = $false
                Message = "Uses OR operator (any control satisfies) instead of AND (all controls required)"
            }
        }
        
        return @{
            Valid = $true
            Message = ""
        }
    }
    
    # Standardized scoring weights (normalized to equal importance per criterion)
    # Each criterion type gets equal weight when present
    $criteriaWeights = @{
        Users = 100
        Applications = 100
        GrantControls = 100
        ClientApps = 100
        SignInRiskLevels = 100
        UserRiskLevels = 100
        Locations = 100
        SessionControls = 100
        UserActions = 100
        DevicePlatforms = 100
    }
    
    # Critical validation: Check policy state first
    if ($Policy.state -ne "enabled") {
        $findings += "⚠ Policy is in '$($Policy.state)' state (not enforcing)"
        # Reduce all scores by 50% for non-enabled policies
        $stateMultiplier = 0.5
    }
    else {
        $stateMultiplier = 1.0
    }
    
    # Check Users assignment
    if ($Recommendation.CheckCriteria.Users) {
        $maxScore += $criteriaWeights.Users
        $expectedUsers = $Recommendation.CheckCriteria.Users
        
        if ($expectedUsers -match "All users" -and $Policy.conditions.users.includeUsers -contains "All") {
            # Check for significant exclusions that might negate the policy
            if (Test-HasSignificantExclusions -Policy $Policy -InclusionType "All") {
                $complianceScore += ($criteriaWeights.Users * 0.5)
                $findings += "⚠ Includes all users but has significant exclusions"
                if ($Policy.conditions.users.excludeUsers -contains "All") {
                    $findings += "  - Excludes: All users (policy effectively disabled)"
                }
                elseif ($Policy.conditions.users.excludeGroups.Count -gt 0) {
                    $findings += "  - Excludes: $($Policy.conditions.users.excludeGroups.Count) group(s)"
                }
                if ($Policy.conditions.users.excludeRoles.Count -gt 0) {
                    $findings += "  - Excludes: $($Policy.conditions.users.excludeRoles.Count) role(s)"
                }
            }
            else {
                $complianceScore += $criteriaWeights.Users
                $findings += "✓ Applies to all users"
            }
        }
        elseif ($expectedUsers -match "Directory roles" -and $Policy.conditions.users.includeRoles) {
            $complianceScore += $criteriaWeights.Users
            $findings += "✓ Applies to directory roles ($($Policy.conditions.users.includeRoles.Count) role(s))"
            if ($Policy.conditions.users.excludeRoles.Count -gt 0) {
                $findings += "  ⚠ Excludes $($Policy.conditions.users.excludeRoles.Count) role(s)"
            }
        }
        elseif ($expectedUsers -match "Guest or external users" -and 
                ($Policy.conditions.users.includeGuestsOrExternalUsers -or 
                 $Policy.conditions.users.guestOrExternalUserTypes)) {
            $complianceScore += $criteriaWeights.Users
            $findings += "✓ Applies to guest/external users"
        }
        else {
            $findings += "✗ User assignment doesn't match recommendation"
        }
    }
    
    # Check Applications
    if ($Recommendation.CheckCriteria.Applications) {
        $maxScore += $criteriaWeights.Applications
        $expectedApps = $Recommendation.CheckCriteria.Applications
        
        if ($expectedApps -match "All resources" -and 
            ($Policy.conditions.applications.includeApplications -contains "All" -or
             $Policy.conditions.applications.includeApplications -contains "Office365")) {
            # Check for application exclusions
            if ($Policy.conditions.applications.excludeApplications -and 
                $Policy.conditions.applications.excludeApplications.Count -gt 5) {
                $complianceScore += ($criteriaWeights.Applications * 0.7)
                $findings += "⚠ Applies to all resources but excludes $($Policy.conditions.applications.excludeApplications.Count) app(s)"
            }
            else {
                $complianceScore += $criteriaWeights.Applications
                $findings += "✓ Applies to all resources"
            }
        }
        elseif ($expectedApps -match "Intune" -and 
                ($Policy.conditions.applications.includeApplications -match "Intune" -or
                 $Policy.conditions.applications.includeApplications -contains "d4ebce55-015a-49b5-a083-c84d1797ae8c")) {
            $complianceScore += $criteriaWeights.Applications
            $findings += "✓ Applies to Microsoft Intune"
        }
        else {
            $findings += "✗ Application scope doesn't match recommendation"
        }
    }
    
    # Check Grant Controls
    if ($Recommendation.CheckCriteria.GrantControls) {
        $maxScore += $criteriaWeights.GrantControls
        $expectedGrant = $Recommendation.CheckCriteria.GrantControls
        
        # Validate grant controls are actually configured
        $hasGrantControls = $Policy.grantControls -and 
                           ($Policy.grantControls.builtInControls -or 
                            $Policy.grantControls.authenticationStrength -or 
                            $Policy.grantControls.termsOfUse)
        
        if ($expectedGrant -eq "Grant access") {
            # "Grant access" means: allow access WITHOUT requiring MFA, device compliance, or blocking
            # This is typically used with session controls (like sign-in frequency limits)
            # Must have grant controls defined and NOT block or require additional auth
            
            if (-not $hasGrantControls) {
                $findings += "✗ No grant controls configured"
            }
            elseif ($Policy.grantControls.builtInControls -contains "block") {
                $findings += "✗ Policy blocks access (expected to grant access)"
            }
            elseif ($Policy.grantControls.builtInControls -contains "mfa" -or
                    $Policy.grantControls.authenticationStrength -or
                    $Policy.grantControls.builtInControls -contains "compliantDevice") {
                # Policy requires additional controls beyond simple access grant
                $findings += "⚠ Policy grants access but requires additional controls (may be acceptable)"
                $complianceScore += ($criteriaWeights.GrantControls * 0.7)
            }
            elseif ($Policy.grantControls._Operator -in @("AND", "OR")) {
                # Valid grant access - allows access without blocking or requiring MFA/compliance
                $complianceScore += $criteriaWeights.GrantControls
                $findings += "✓ Grants access as recommended"
            }
            else {
                $findings += "✗ Grant controls configuration unclear"
            }
        }
        elseif ($expectedGrant -match "Block access" -and $hasGrantControls -and 
                $Policy.grantControls.builtInControls -contains "block") {
            $complianceScore += $criteriaWeights.GrantControls
            $findings += "✓ Blocks access as recommended"
        }
        elseif ($expectedGrant -match "Require device to be marked as compliant") {
            if (-not $hasGrantControls) {
                $findings += "✗ No grant controls configured"
            }
            elseif ($Policy.grantControls.builtInControls -contains "compliantDevice") {
                # Check operator if multiple controls exist
                $controlCount = ($Policy.grantControls.builtInControls | Measure-Object).Count
                if ($controlCount -gt 1) {
                    $operatorCheck = Test-GrantControlsOperator -Policy $Policy -RequiredControls $controlCount
                    if (-not $operatorCheck.Valid) {
                        $complianceScore += ($criteriaWeights.GrantControls * 0.8)
                        $findings += "⚠ Requires compliant device (but $($operatorCheck.Message))"
                    }
                    else {
                        $complianceScore += $criteriaWeights.GrantControls
                        $findings += "✓ Requires compliant device"
                    }
                }
                else {
                    $complianceScore += $criteriaWeights.GrantControls
                    $findings += "✓ Requires compliant device"
                }
            }
            else {
                $findings += "✗ Does not require compliant device"
            }
        }
        elseif ($expectedGrant -match "Phishing-resistant MFA") {
            if (-not $hasGrantControls) {
                $findings += "✗ No grant controls configured"
            }
            elseif ($Policy.grantControls.authenticationStrength) {
                # Validate that the authentication strength is actually phishing-resistant
                $authStrengthId = $Policy.grantControls.authenticationStrength.id
                $authStrengthCheck = Test-IsPhishingResistantAuthStrength -AuthStrengthId $authStrengthId -AuthStrengthPolicies $authStrengthPolicies
                
                if ($authStrengthCheck.IsPhishingResistant) {
                    $complianceScore += $criteriaWeights.GrantControls
                    $findings += "✓ Requires phishing-resistant authentication strength"
                    $findings += "  Policy: $($authStrengthCheck.PolicyName) (ID: $authStrengthId)"
                    if ($authStrengthCheck.Methods -and $authStrengthCheck.Methods.Count -gt 0) {
                        $methodsList = ($authStrengthCheck.Methods | Select-Object -First 3) -join ", "
                        if ($authStrengthCheck.Methods.Count -gt 3) {
                            $methodsList += " (and $($authStrengthCheck.Methods.Count - 3) more)"
                        }
                        $findings += "  Allowed methods: $methodsList"
                    }
                }
                else {
                    $complianceScore += ($criteriaWeights.GrantControls * 0.5)
                    $findings += "⚠ Authentication strength configured but NOT phishing-resistant"
                    $findings += "  Policy: $($authStrengthCheck.PolicyName) (ID: $authStrengthId)"
                    $findings += "  Issue: $($authStrengthCheck.Reason)"
                    if ($authStrengthCheck.NonPhishingResistantMethods) {
                        $findings += "  Non-phishing-resistant methods: $($authStrengthCheck.NonPhishingResistantMethods -join ', ')"
                    }
                }
            }
            else {
                $findings += "✗ Does not require authentication strength (phishing-resistant MFA)"
                $findings += "  Note: Built-in MFA control is not phishing-resistant"
            }
        }
        elseif ($expectedGrant -match "Require MFA") {
            if (-not $hasGrantControls) {
                $findings += "✗ No grant controls configured"
            }
            elseif ($Policy.grantControls.builtInControls -contains "mfa" -or
                    $Policy.grantControls.authenticationStrength) {
                if ($Policy.grantControls.authenticationStrength) {
                    $authStrengthId = $Policy.grantControls.authenticationStrength.id
                    $authStrengthCheck = Test-IsPhishingResistantAuthStrength -AuthStrengthId $authStrengthId -AuthStrengthPolicies $authStrengthPolicies
                    
                    $complianceScore += $criteriaWeights.GrantControls
                    $findings += "✓ Requires MFA via authentication strength"
                    $findings += "  Policy: $($authStrengthCheck.PolicyName) (ID: $authStrengthId)"
                    
                    if ($authStrengthCheck.IsPhishingResistant) {
                        $findings += "  ⓘ This authentication strength is phishing-resistant (exceeds requirement)"
                    }
                    else {
                        $findings += "  ⓘ Note: $($authStrengthCheck.Reason)"
                    }
                }
                else {
                    $complianceScore += $criteriaWeights.GrantControls
                    $findings += "✓ Requires MFA (built-in)"
                    $findings += "  ⓘ Consider using authentication strength for more granular control"
                }
            }
            else {
                $findings += "✗ Does not require MFA"
            }
        }
        elseif ($expectedGrant -match "Require terms of use") {
            if (-not $hasGrantControls) {
                $findings += "✗ No grant controls configured"
            }
            elseif ($Policy.grantControls.termsOfUse) {
                $complianceScore += $criteriaWeights.GrantControls
                $findings += "✓ Requires terms of use ($($Policy.grantControls.termsOfUse.Count) ToU)"
            }
            else {
                $findings += "✗ Does not require terms of use"
            }
        }
        else {
            # Provide detailed information about what's expected vs what's configured
            if (-not $hasGrantControls) {
                $findings += "✗ Grant controls don't match recommendation"
                $findings += "  Expected: $expectedGrant"
                $findings += "  Current: None configured"
            }
            else {
                $actualControls = @()
                if ($Policy.grantControls.builtInControls) {
                    $actualControls += $Policy.grantControls.builtInControls | ForEach-Object { 
                        switch ($_) {
                            "block" { "Block access" }
                            "mfa" { "Require MFA" }
                            "compliantDevice" { "Require compliant device" }
                            "domainJoinedDevice" { "Require domain joined device" }
                            "approvedApplication" { "Require approved client app" }
                            "compliantApplication" { "Require app protection policy" }
                            default { $_ }
                        }
                    }
                }
                if ($Policy.grantControls.authenticationStrength) {
                    $actualControls += "Require authentication strength"
                }
                if ($Policy.grantControls.termsOfUse) {
                    $actualControls += "Require terms of use"
                }
                
                $actualControlsText = if ($actualControls.Count -gt 0) { 
                    $controlsText = ($actualControls -join ", ")
                    # Add operator info if multiple controls
                    if ($actualControls.Count -gt 1) {
                        $operator = $Policy.grantControls._Operator
                        "$controlsText (operator: $operator)"
                    }
                    else {
                        $controlsText
                    }
                } else { 
                    "Grant access (no additional controls)" 
                }
                
                $findings += "✗ Grant controls don't match recommendation"
                $findings += "  Expected: $expectedGrant"
                $findings += "  Current: $actualControlsText"
            }
        }
    }
    
    # Check Client Apps
    if ($Recommendation.CheckCriteria.ClientApps) {
        $maxScore += $criteriaWeights.ClientApps
        if ($Policy.conditions.clientAppTypes -and 
            ($Policy.conditions.clientAppTypes -contains "exchangeActiveSync" -or
             $Policy.conditions.clientAppTypes -contains "other")) {
            $complianceScore += $criteriaWeights.ClientApps
            $findings += "✓ Targets legacy authentication client apps"
        }
        else {
            $findings += "✗ Client app types not configured for legacy auth"
        }
    }
    
    # Check Sign-in Risk
    if ($Recommendation.CheckCriteria.SignInRiskLevels) {
        $maxScore += $criteriaWeights.SignInRiskLevels
        $expectedRisk = $Recommendation.CheckCriteria.SignInRiskLevels
        
        if (-not $Policy.conditions.signInRiskLevels -or $Policy.conditions.signInRiskLevels.Count -eq 0) {
            $findings += "✗ Sign-in risk not configured"
        }
        else {
            # Define what risk levels are configured
            $hasHigh = $Policy.conditions.signInRiskLevels -contains "high"
            $hasMedium = $Policy.conditions.signInRiskLevels -contains "medium"
            $hasLow = $Policy.conditions.signInRiskLevels -contains "low"
            $configuredLevels = @($Policy.conditions.signInRiskLevels) -join ", "
            
            if ($expectedRisk -match "High") {
                # Expecting high risk only
                if ($hasHigh -and -not $hasMedium -and -not $hasLow) {
                    $complianceScore += $criteriaWeights.SignInRiskLevels
                    $findings += "✓ Configured for high sign-in risk"
                }
                elseif ($hasHigh) {
                    # Has high but also includes medium/low (too broad)
                    $complianceScore += ($criteriaWeights.SignInRiskLevels * 0.7)
                    $findings += "⚠ Configured for high sign-in risk but also includes: $configuredLevels"
                    $findings += "  (May trigger on lower-risk events than intended)"
                }
                else {
                    # Missing high risk
                    $findings += "✗ Does not include high sign-in risk (configured: $configuredLevels)"
                }
            }
            elseif ($expectedRisk -match "Medium") {
                # Expecting medium or high risk
                if ($hasMedium -or $hasHigh) {
                    if ($hasLow) {
                        # Includes low risk (too broad)
                        $complianceScore += ($criteriaWeights.SignInRiskLevels * 0.7)
                        $findings += "⚠ Configured for medium/high sign-in risk but also includes low"
                        $findings += "  Configured levels: $configuredLevels"
                    }
                    else {
                        $complianceScore += $criteriaWeights.SignInRiskLevels
                        $findings += "✓ Configured for medium/high sign-in risk ($configuredLevels)"
                    }
                }
                else {
                    # Only has low risk
                    $complianceScore += ($criteriaWeights.SignInRiskLevels * 0.3)
                    $findings += "⚠ Only configured for low sign-in risk (expected: medium/high)"
                    $findings += "  Configured levels: $configuredLevels"
                }
            }
            else {
                # Unknown risk level expectation - show what's configured
                if ($hasHigh -and $hasMedium) {
                    $complianceScore += $criteriaWeights.SignInRiskLevels
                    $findings += "✓ Sign-in risk configured ($configuredLevels)"
                }
                else {
                    $complianceScore += ($criteriaWeights.SignInRiskLevels * 0.5)
                    $findings += "⚠ Sign-in risk partially configured ($configuredLevels)"
                }
            }
        }
    }
    
    # Check User Risk
    if ($Recommendation.CheckCriteria.UserRiskLevels) {
        $maxScore += $criteriaWeights.UserRiskLevels
        $expectedUserRisk = $Recommendation.CheckCriteria.UserRiskLevels
        
        if (-not $Policy.conditions.userRiskLevels -or $Policy.conditions.userRiskLevels.Count -eq 0) {
            $findings += "✗ User risk not configured"
        }
        else {
            # Define what risk levels are configured
            $hasHigh = $Policy.conditions.userRiskLevels -contains "high"
            $hasMedium = $Policy.conditions.userRiskLevels -contains "medium"
            $hasLow = $Policy.conditions.userRiskLevels -contains "low"
            $configuredLevels = @($Policy.conditions.userRiskLevels) -join ", "
            
            if ($expectedUserRisk -match "High") {
                # Expecting high risk only
                if ($hasHigh -and -not $hasMedium -and -not $hasLow) {
                    $complianceScore += $criteriaWeights.UserRiskLevels
                    $findings += "✓ Configured for high user risk"
                }
                elseif ($hasHigh) {
                    # Has high but also includes medium/low (too broad)
                    $complianceScore += ($criteriaWeights.UserRiskLevels * 0.7)
                    $findings += "⚠ Configured for high user risk but also includes: $configuredLevels"
                    $findings += "  (May trigger on lower-risk users than intended)"
                }
                else {
                    # Missing high risk
                    $findings += "✗ Does not include high user risk (configured: $configuredLevels)"
                }
            }
            elseif ($expectedUserRisk -match "Medium") {
                # Expecting medium or high risk
                if ($hasMedium -or $hasHigh) {
                    if ($hasLow) {
                        # Includes low risk (too broad)
                        $complianceScore += ($criteriaWeights.UserRiskLevels * 0.7)
                        $findings += "⚠ Configured for medium/high user risk but also includes low"
                        $findings += "  Configured levels: $configuredLevels"
                    }
                    else {
                        $complianceScore += $criteriaWeights.UserRiskLevels
                        $findings += "✓ Configured for medium/high user risk ($configuredLevels)"
                    }
                }
                else {
                    # Only has low risk
                    $complianceScore += ($criteriaWeights.UserRiskLevels * 0.3)
                    $findings += "⚠ Only configured for low user risk (expected: medium/high)"
                    $findings += "  Configured levels: $configuredLevels"
                }
            }
            else {
                # Unknown risk level expectation - show what's configured
                if ($hasHigh -and $hasMedium) {
                    $complianceScore += $criteriaWeights.UserRiskLevels
                    $findings += "✓ User risk configured ($configuredLevels)"
                }
                else {
                    $complianceScore += ($criteriaWeights.UserRiskLevels * 0.5)
                    $findings += "⚠ User risk partially configured ($configuredLevels)"
                }
            }
        }
    }
    
    # Check Locations
    if ($Recommendation.CheckCriteria.Locations) {
        $maxScore += $criteriaWeights.Locations
        if ($Policy.conditions.locations) {
            $complianceScore += $criteriaWeights.Locations
            $findings += "✓ Location-based conditions configured"
        }
        else {
            $findings += "✗ Location conditions not configured"
        }
    }
    
    # Check Session Controls
    if ($Recommendation.CheckCriteria.SessionControls) {
        $maxScore += $criteriaWeights.SessionControls
        if ($Policy.sessionControls.signInFrequency) {
            $complianceScore += $criteriaWeights.SessionControls
            $findings += "✓ Sign-in frequency configured"
        }
        else {
            $findings += "✗ Session sign-in frequency not configured"
        }
    }
    
    # Check User Actions
    if ($Recommendation.CheckCriteria.UserActions) {
        $maxScore += $criteriaWeights.UserActions
        $expectedActions = $Recommendation.CheckCriteria.UserActions
        if ($expectedActions -match "Register security information" -and
            $Policy.conditions.userActions -contains "registerSecurityInfo") {
            $complianceScore += $criteriaWeights.UserActions
            $findings += "✓ Applies to security info registration"
        }
        else {
            $findings += "✗ User actions don't match recommendation"
        }
    }
    
    # Check Device Platforms
    if ($Recommendation.CheckCriteria.DevicePlatforms) {
        $maxScore += $criteriaWeights.DevicePlatforms
        if ($Policy.conditions.platforms -and $Policy.conditions.platforms.includePlatforms) {
            $complianceScore += $criteriaWeights.DevicePlatforms
            $findings += "✓ Device platforms configured"
        }
        else {
            $findings += "✗ Device platforms not configured"
        }
    }
    
    # Check if policy is enabled
    $policyEnabled = $Policy.state -eq "enabled"
    
    # Apply state multiplier to compliance score (reduces score for non-enabled policies)
    $complianceScore = [math]::Round($complianceScore * $stateMultiplier, 0)
    
    # Calculate final percentage
    if ($maxScore -eq 0) { $maxScore = 100 }
    $compliancePercentage = [math]::Round(($complianceScore / $maxScore) * 100, 0)
    
    # Add policy state summary to findings
    if (-not $policyEnabled) {
        $findings = @("⚠ CRITICAL: Policy state is '$($Policy.state)' - not actively enforcing (score reduced by 50%)") + $findings
    }
    
    return @{
        ComplianceScore = $complianceScore
        MaxScore = $maxScore
        CompliancePercentage = $compliancePercentage
        Findings = $findings
        PolicyEnabled = $policyEnabled
    }
}

# Evaluate policies
$evaluationResults = @()

foreach ($recommendation in $asdRecommendations) {
    Write-Host "`nEvaluating: $($recommendation.Name)" -ForegroundColor Yellow
    
    # Try to find matching policies based primarily on conditions
    # 
    # Condition Matching Strategy:
    # - Each recommendation type gets a standardized score (100 points) for matching its key condition
    # - This identifies which policies are RELEVANT to the recommendation
    # - Grant controls and other detailed requirements are evaluated separately in compliance scoring
    # - This separation prevents double-counting and ensures fair comparison across policy types
    # 
    # Matching Threshold:
    # - Policies must score >= 50% on condition matching to be considered
    # - This ensures only genuinely relevant policies are evaluated for compliance
    # - The 50% threshold allows for partial matches (e.g., guest policies with some but not all criteria)
    #
    $matchingPolicies = @()
    
    foreach ($policy in $policies) {
        # Calculate condition-based match score (0-100)
        $conditionScore = 0
        $maxConditionScore = 0
        
        # Legacy auth check - identifies policies targeting legacy authentication clients
        if ($recommendation.Key -eq "block-legacy-auth") {
            $maxConditionScore += 100
            if ($policy.conditions.clientAppTypes -and
                ($policy.conditions.clientAppTypes -contains "exchangeActiveSync" -or
                 $policy.conditions.clientAppTypes -contains "other")) {
                $conditionScore += 100
            }
        }
        
        # High-risk sign-in blocking - identifies policies targeting high sign-in risk
        if ($recommendation.Key -eq "block-high-risk-signins") {
            $maxConditionScore += 100
            if ($policy.conditions.signInRiskLevels -and 
                $policy.conditions.signInRiskLevels -contains "high") {
                $conditionScore += 100
            }
        }
        
        # High-risk user blocking - identifies policies targeting high user risk
        if ($recommendation.Key -eq "block-high-risk-users") {
            $maxConditionScore += 100
            if ($policy.conditions.userRiskLevels -and 
                $policy.conditions.userRiskLevels -contains "high") {
                $conditionScore += 100
            }
        }
        
        # Device compliance check - identifies policies requiring compliant devices
        if ($recommendation.Key -eq "compliant-devices") {
            $maxConditionScore += 100
            if ($policy.grantControls.builtInControls -contains "compliantDevice") {
                $conditionScore += 100
            }
        }
        
        # Intune enrollment check - identifies policies targeting Intune enrollment
        if ($recommendation.Key -eq "intune-enrolment") {
            $maxConditionScore += 100
            if ($policy.conditions.applications.includeApplications -and
                ($policy.conditions.applications.includeApplications -match "Intune" -or
                 $policy.conditions.applications.includeApplications -contains "d4ebce55-015a-49b5-a083-c84d1797ae8c")) {
                $conditionScore += 100
            }
        }
        
        # Guest blocking or authentication - identifies policies targeting guest users
        if ($recommendation.Key -match "guest") {
            $maxConditionScore += 100
            if ($policy.conditions.users.includeGuestsOrExternalUsers -or
                $policy.conditions.users.includeUsers -contains "GuestsOrExternalUsers") {
                if ($recommendation.Key -eq "block-guests") {
                    # Blocking guests - check that it actually blocks
                    if ($policy.grantControls.builtInControls -contains "block") {
                        $conditionScore += 100
                    }
                }
                elseif ($recommendation.Key -eq "guest-app-access") {
                    # Guest app access - check for specific app targeting (not All)
                    $hasSpecificApp = $policy.conditions.applications.includeApplications -and
                                      $policy.conditions.applications.includeApplications -notcontains "All" -and
                                      $policy.conditions.applications.includeApplications.Count -gt 0
                    if ($hasSpecificApp) {
                        $conditionScore += 100
                    }
                    else {
                        # Partial match - targets guests but not specific apps
                        $conditionScore += 50
                    }
                }
            }
        }
        
        # Location-based checks - identifies policies with location conditions
        if ($recommendation.Key -match "location" -or $recommendation.Key -match "countr") {
            $maxConditionScore += 100
            if ($policy.conditions.locations) {
                $conditionScore += 100
            }
        }
        
        # Admin session limiting - identifies policies targeting admin roles
        if ($recommendation.Key -eq "limit-admin-sessions") {
            $maxConditionScore += 100
            if ($policy.conditions.users.includeRoles -and $policy.conditions.users.includeRoles.Count -gt 0) {
                $conditionScore += 100
            }
        }
        
        # User session limiting - identifies policies with session controls
        if ($recommendation.Key -eq "limit-user-sessions") {
            $maxConditionScore += 100
            if ($policy.sessionControls.signInFrequency) {
                $conditionScore += 100
            }
        }
        
        # Strong auth for all users - identifies policies requiring strong authentication
        if ($recommendation.Key -eq "strong-auth") {
            $maxConditionScore += 100
            if ($policy.grantControls.authenticationStrength -or
                $policy.grantControls.builtInControls -contains "mfa") {
                $conditionScore += 100
            }
        }
        
        # Risky sign-ins with strong auth - identifies policies targeting risky sign-ins
        if ($recommendation.Key -eq "risky-signins-auth") {
            $maxConditionScore += 100
            if ($policy.conditions.signInRiskLevels -and 
                ($policy.conditions.signInRiskLevels -contains "medium" -or
                 $policy.conditions.signInRiskLevels -contains "high")) {
                $conditionScore += 100
            }
        }
        
        # Register security info - identifies policies for security info registration
        if ($recommendation.Key -eq "register-security-info") {
            $maxConditionScore += 100
            if ($policy.conditions.userActions -and 
                $policy.conditions.userActions -contains "registerSecurityInfo") {
                $conditionScore += 100
            }
        }
        
        # Terms of use - identifies policies requiring terms of use acceptance
        if ($recommendation.Key -eq "terms-of-use") {
            $maxConditionScore += 100
            if ($policy.grantControls.termsOfUse -and $policy.grantControls.termsOfUse.Count -gt 0) {
                $conditionScore += 100
            }
        }
        
        # Block unapproved devices - identifies policies with device platform conditions
        if ($recommendation.Key -eq "block-unapproved-devices") {
            $maxConditionScore += 100
            if ($policy.conditions.platforms -and $policy.conditions.platforms.includePlatforms) {
                $conditionScore += 100
            }
        }
        
        # Insider risk - identifies policies targeting insider risk levels
        # Note: Requires Microsoft Purview Insider Risk Management
        if ($recommendation.Key -eq "block-insider-risk") {
            if ($insiderRiskAvailable) {
                $maxConditionScore += 100
                if ($policy.conditions.insiderRiskLevels -and
                    $policy.conditions.insiderRiskLevels -contains "elevated") {
                    $conditionScore += 100
                }
            }
        }
        
        # Protected information - identifies policies with authentication context conditions
        # Note: Requires authentication contexts to be configured
        if ($recommendation.Key -eq "block-protected-info") {
            if ($authContexts.Count -gt 0) {
                $maxConditionScore += 100
                if ($policy.conditions.authenticationContextClassReferences) {
                    $conditionScore += 100
                }
            }
        }
        
        # Calculate match percentage based on conditions
        $matchPercentage = 0
        if ($maxConditionScore -gt 0) {
            $matchPercentage = [math]::Round(($conditionScore / $maxConditionScore) * 100, 0)
        }
        
        # Include policy if it matches conditions with at least 50% confidence
        if ($matchPercentage -ge 50) {
            $matchingPolicies += @{
                Policy = $policy
                MatchScore = $matchPercentage
            }
        }
    }
    
    # Sort matching policies by match score (highest first)
    $matchingPolicies = $matchingPolicies | Sort-Object -Property MatchScore -Descending
    
    if ($matchingPolicies.Count -eq 0) {
        # No matching policy found
        $missingReason = "✗ No policy found matching this recommendation based on condition analysis"
        $featureUnavailable = $false
        
        # Add context for feature availability issues
        if ($recommendation.Key -eq "block-insider-risk" -and -not $insiderRiskAvailable) {
            $missingReason = "⚠ Insider Risk Management not detected. This recommendation requires Microsoft Purview Insider Risk Management with appropriate licensing."
            $featureUnavailable = $true
        }
        elseif ($recommendation.Key -eq "block-protected-info" -and $authContexts.Count -eq 0) {
            $missingReason = "⚠ No authentication contexts configured. This recommendation requires authentication contexts to be created and configured in your tenant."
            $featureUnavailable = $true
        }
        elseif ($recommendation.Key -match "strong-auth|register-security-info" -and -not $authStrengthAvailable) {
            $missingReason = "⚠ Authentication strength policies not available or configured. This recommendation requires authentication strength policies with Microsoft Entra ID P1/P2 licensing."
            $featureUnavailable = $true
        }
        elseif ($recommendation.Key -eq "terms-of-use" -and -not $termsOfUseAvailable) {
            $missingReason = "⚠ Terms of Use not configured. This recommendation requires Terms of Use agreements to be created in your tenant (requires Microsoft Entra ID P1/P2 licensing)."
            $featureUnavailable = $true
        }
        
        $evaluationResults += @{
            Recommendation = $recommendation
            Status = "Missing"
            CompliancePercentage = 0
            MatchingPolicies = @()
            BestMatch = $null
            Findings = @($missingReason)
            FeatureUnavailable = $featureUnavailable
        }
        Write-Host "  Status: Missing" -ForegroundColor $(if ($featureUnavailable) { "Yellow" } else { "Red" })
        if ($featureUnavailable) {
            Write-Host "  Note: Feature not available or configured" -ForegroundColor Gray
        }
    }
    else {
        # Evaluate each matching policy and consider both condition match score and compliance score
        $bestMatch = $null
        $bestCombinedScore = 0
        
        foreach ($matchResult in $matchingPolicies) {
            $policy = $matchResult.Policy
            $conditionMatchScore = $matchResult.MatchScore
            
            $compliance = Test-PolicyCompliance -Policy $policy -Recommendation $recommendation
            
            # Combined score: 40% condition match + 60% compliance score
            $combinedScore = ($conditionMatchScore * 0.4) + ($compliance.CompliancePercentage * 0.6)
            
            if ($combinedScore -gt $bestCombinedScore) {
                $bestCombinedScore = $combinedScore
                $bestMatch = @{
                    Policy = $policy
                    Compliance = $compliance
                    ConditionMatchScore = $conditionMatchScore
                    CombinedScore = [math]::Round($combinedScore, 0)
                }
            }
        }
        
        # Use the combined score for status determination (considers both condition match and compliance)
        $bestScore = $bestMatch.CombinedScore
        
        # COMPLIANCE THRESHOLDS - Configurable based on organizational risk tolerance
        # Current thresholds are balanced for practical deployment scenarios:
        # - 80%+ = Compliant: Policy substantially meets requirements (allows minor gaps)
        # - 50-79% = Partial: Core requirements met but significant gaps exist
        # - <50% = Non-Compliant: Fundamental requirements not met
        #
        # ADJUSTMENT GUIDANCE:
        # - Increase to 90% if strict compliance required (may increase false negatives)
        # - Decrease to 70% if more lenient assessment needed (may miss gaps)
        # - Consider organizational factors: exclusion policies, phased rollouts, compensating controls
        $complianceThreshold = 80  # Minimum score for "Compliant" status
        $partialThreshold = 50     # Minimum score for "Partial" status
        
        $status = if ($bestScore -ge $complianceThreshold) { "Compliant" }
                  elseif ($bestScore -ge $partialThreshold) { "Partial" }
                  else { "Non-Compliant" }
        
        $evaluationResults += @{
            Recommendation = $recommendation
            Status = $status
            CompliancePercentage = $bestScore
            ConditionMatchScore = $bestMatch.ConditionMatchScore
            ActualComplianceScore = $bestMatch.Compliance.CompliancePercentage
            MatchingPolicies = $matchingPolicies
            BestMatch = $bestMatch
            Findings = $bestMatch.Compliance.Findings
        }
        
        Write-Host "  Status: $status ($bestScore%)" -ForegroundColor $(
            if ($status -eq "Compliant") { "Green" }
            elseif ($status -eq "Partial") { "Yellow" }
            else { "Red" }
        )
        Write-Host "  Best matching policy: $($bestMatch.Policy.displayName)" -ForegroundColor Cyan
        Write-Host "  - Combined Score: $($bestMatch.CombinedScore)% (Condition: $($bestMatch.ConditionMatchScore)%, Compliance: $($bestMatch.Compliance.CompliancePercentage)%)" -ForegroundColor Gray
    }
}

# Generate HTML Report
$reportDate = Get-Date -Format "yyyyMMdd-HHmmss"
$reportPath = Join-Path (Split-Path $PSScriptRoot -Parent) "ASD-CA-Evaluation-Report-$reportDate.html"

# Calculate overall statistics
$totalRecommendations = $asdRecommendations.Count
$compliantCount = ($evaluationResults | Where-Object { $_.Status -eq "Compliant" }).Count
$partialCount = ($evaluationResults | Where-Object { $_.Status -eq "Partial" }).Count
# Combine Non-Compliant and Missing into a single "Non-Compliant" metric
# Both represent recommendations that are not adequately met
$nonCompliantCount = ($evaluationResults | Where-Object { $_.Status -eq "Non-Compliant" -or $_.Status -eq "Missing" }).Count
$missingCount = ($evaluationResults | Where-Object { $_.Status -eq "Missing" }).Count
$overallCompliance = [math]::Round((($evaluationResults | Measure-Object -Property CompliancePercentage -Average).Average), 0)

# Count by priority
$criticalIssues = ($evaluationResults | Where-Object { $_.Recommendation.Priority -eq "Critical" -and $_.Status -ne "Compliant" }).Count
$highIssues = ($evaluationResults | Where-Object { $_.Recommendation.Priority -eq "High" -and $_.Status -ne "Compliant" }).Count
$mediumIssues = ($evaluationResults | Where-Object { $_.Recommendation.Priority -eq "Medium" -and $_.Status -ne "Compliant" }).Count
$lowIssues = ($evaluationResults | Where-Object { $_.Recommendation.Priority -eq "Low" -and $_.Status -ne "Compliant" }).Count

# Group by category
$categories = $asdRecommendations | Select-Object -ExpandProperty Category -Unique

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ASD Conditional Access Evaluation Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            background: #f5f5f5;
            padding: 20px;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
        }
        
        .header p {
            font-size: 1.1em;
            opacity: 0.9;
        }
        
        .tenant-info {
            background: #f8f9fa;
            padding: 20px 30px;
            border-bottom: 2px solid #e9ecef;
        }
        
        .tenant-info h2 {
            color: #495057;
            margin-bottom: 10px;
        }
        
        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }
        
        .info-item {
            display: flex;
            align-items: center;
        }
        
        .info-label {
            font-weight: 600;
            color: #6c757d;
            margin-right: 10px;
        }
        
        .info-value {
            color: #212529;
        }
        
        .summary {
            padding: 30px;
            background: linear-gradient(to bottom, #ffffff 0%, #f8f9fa 100%);
        }
        
        .summary h2 {
            color: #495057;
            margin-bottom: 20px;
            font-size: 1.8em;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        
        .stat-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            text-align: center;
            border-left: 4px solid;
        }
        
        .stat-card.compliant {
            border-color: #28a745;
        }
        
        .stat-card.partial {
            border-color: #ffc107;
        }
        
        .stat-card.non-compliant {
            border-color: #dc3545;
        }
        
        .stat-card.missing {
            border-color: #6c757d;
        }
        
        .stat-card.overall {
            border-color: #667eea;
        }
        
        .stat-number {
            font-size: 3em;
            font-weight: bold;
            margin: 10px 0;
        }
        
        .stat-label {
            color: #6c757d;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .stat-sublabel {
            color: #9ca3af;
            font-size: 0.75em;
            font-style: italic;
            margin-top: 5px;
        }
        
        .priority-alerts {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }
        
        .priority-card {
            background: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            text-align: center;
        }
        
        .priority-card.critical {
            background: #fff5f5;
            border: 2px solid #dc3545;
        }
        
        .priority-card.high {
            background: #fff8e1;
            border: 2px solid #ff9800;
        }
        
        .priority-card.medium {
            background: #fffbf0;
            border: 2px solid #ffc107;
        }
        
        .priority-card.low {
            background: #f0f8ff;
            border: 2px solid #17a2b8;
        }
        
        .priority-number {
            font-size: 2em;
            font-weight: bold;
            margin: 5px 0;
        }
        
        .priority-label {
            font-size: 0.85em;
            font-weight: 600;
        }
        
        .content {
            padding: 30px;
        }
        
        .category-section {
            margin-bottom: 40px;
        }
        
        .category-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            font-size: 1.4em;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            justify-content: space-between;
            align-items: center;
            transition: opacity 0.3s;
        }
        
        .category-header:hover {
            opacity: 0.9;
        }
        
        .category-header::after {
            content: '▼';
            font-size: 0.8em;
            transition: transform 0.3s;
        }
        
        .category-header.collapsed::after {
            transform: rotate(-90deg);
        }
        
        .category-content {
            max-height: 10000px;
            overflow: hidden;
            transition: max-height 0.5s ease-out;
        }
        
        .category-content.collapsed {
            max-height: 0;
        }
        
        .recommendation-card {
            background: white;
            border: 1px solid #e9ecef;
            border-radius: 8px;
            margin-bottom: 20px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            transition: box-shadow 0.3s;
        }
        
        .recommendation-card:hover {
            box-shadow: 0 4px 16px rgba(0,0,0,0.1);
        }
        
        .recommendation-header {
            padding: 20px;
            background: #f8f9fa;
            border-bottom: 2px solid #e9ecef;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
            cursor: pointer;
            transition: background 0.3s;
        }
        
        .recommendation-header:hover {
            background: #e9ecef;
        }
        
        .recommendation-header::after {
            content: '▼';
            font-size: 0.7em;
            color: #6c757d;
            margin-left: 10px;
            transition: transform 0.3s;
        }
        
        .recommendation-header.collapsed::after {
            transform: rotate(-180deg);
        }
        
        .recommendation-title {
            font-size: 1.2em;
            font-weight: 600;
            color: #212529;
            flex: 1;
        }
        
        .status-badge {
            padding: 8px 16px;
            border-radius: 20px;
            font-weight: 600;
            font-size: 0.85em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .status-badge.compliant {
            background: #d4edda;
            color: #155724;
        }
        
        .status-badge.partial {
            background: #fff3cd;
            color: #856404;
        }
        
        .status-badge.non-compliant {
            background: #f8d7da;
            color: #721c24;
        }
        
        .status-badge.missing {
            background: #e2e3e5;
            color: #383d41;
        }
        
        .priority-badge {
            padding: 6px 12px;
            border-radius: 4px;
            font-weight: 600;
            font-size: 0.8em;
        }
        
        .priority-badge.critical {
            background: #dc3545;
            color: white;
        }
        
        .priority-badge.high {
            background: #ff9800;
            color: white;
        }
        
        .priority-badge.medium {
            background: #ffc107;
            color: #333;
        }
        
        .priority-badge.low {
            background: #17a2b8;
            color: white;
        }
        
        .compliance-bar {
            height: 8px;
            background: #e9ecef;
            border-radius: 4px;
            overflow: hidden;
            margin-top: 10px;
        }
        
        .compliance-fill {
            height: 100%;
            background: linear-gradient(90deg, #28a745 0%, #20c997 100%);
            transition: width 0.5s;
        }
        
        .recommendation-body {
            padding: 20px;
            max-height: 1000px;
            overflow: hidden;
            transition: max-height 0.4s ease-out, padding 0.4s ease-out;
        }
        
        .recommendation-body.collapsed {
            max-height: 0;
            padding-top: 0;
            padding-bottom: 0;
        }
        
        .description {
            color: #6c757d;
            margin-bottom: 15px;
            font-style: italic;
        }
        
        .findings {
            margin-top: 15px;
        }
        
        .findings h4 {
            color: #495057;
            margin-bottom: 10px;
            font-size: 1em;
        }
        
        .findings ul {
            list-style: none;
            padding-left: 0;
        }
        
        .findings li {
            padding: 8px 12px;
            margin-bottom: 5px;
            border-radius: 4px;
            background: #f8f9fa;
        }
        
        .findings li:before {
            margin-right: 8px;
        }
        
        .matching-policy {
            background: #e7f3ff;
            border-left: 4px solid #0066cc;
            padding: 15px;
            margin-top: 15px;
            border-radius: 4px;
        }
        
        .matching-policy h4 {
            color: #0066cc;
            margin-bottom: 8px;
        }
        
        .policy-details {
            display: grid;
            grid-template-columns: auto 1fr;
            gap: 10px;
            margin-top: 10px;
            font-size: 0.9em;
        }
        
        .policy-label {
            font-weight: 600;
            color: #495057;
        }
        
        .policy-value {
            color: #6c757d;
        }
        
        .footer {
            background: #343a40;
            color: white;
            padding: 20px;
            text-align: center;
        }
        
        .footer a {
            color: #80bdff;
            text-decoration: none;
        }
        
        .footer a:hover {
            text-decoration: underline;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            
            .container {
                box-shadow: none;
            }
            
            .recommendation-card {
                page-break-inside: avoid;
            }
            
            .category-header::after,
            .recommendation-header::after {
                display: none;
            }
            
            .category-content,
            .recommendation-body {
                max-height: none !important;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🛡️ ASD Conditional Access Evaluation Report</h1>
            <p>Australian Signals Directorate Blueprint for Secure Cloud - Compliance Assessment</p>
        </div>
        
        <div class="tenant-info">
            <h2>Tenant Information</h2>
            <div class="info-grid">
                <div class="info-item">
                    <span class="info-label">Tenant Name:</span>
                    <span class="info-value">$tenantName</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Tenant Domain:</span>
                    <span class="info-value">$tenantDomain</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Tenant ID:</span>
                    <span class="info-value">$tenantId</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Report Date:</span>
                    <span class="info-value">$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Total Policies:</span>
                    <span class="info-value">$($policies.Count)</span>
                </div>
            </div>
            
            <h3 style="margin-top: 20px; margin-bottom: 10px; color: #495057; font-size: 1.1em;">Feature Availability</h3>
            <div class="info-grid">
                <div class="info-item">
                    <span class="info-label">Authentication Contexts:</span>
                    <span class="info-value" style="color: $(if ($authContexts.Count -gt 0) { '#28a745' } else { '#ffc107' }); font-weight: 600;">
                        $(if ($authContexts.Count -gt 0) { "$($authContexts.Count) configured" } else { "Not configured" })
                    </span>
                </div>
                <div class="info-item">
                    <span class="info-label">Insider Risk Management:</span>
                    <span class="info-value" style="color: $(if ($insiderRiskAvailable) { '#28a745' } else { '#ffc107' }); font-weight: 600;">
                        $(if ($insiderRiskAvailable) { "Available" } else { "Not detected" })
                    </span>
                </div>
                <div class="info-item">
                    <span class="info-label">Authentication Strength:</span>
                    <span class="info-value" style="color: $(if ($authStrengthAvailable) { '#28a745' } else { '#ffc107' }); font-weight: 600;">
                        $(if ($authStrengthAvailable) { "$($authStrengthPolicies.Count) policies" } else { "Not available" })
                    </span>
                </div>
                <div class="info-item">
                    <span class="info-label">Terms of Use:</span>
                    <span class="info-value" style="color: $(if ($termsOfUseAvailable) { '#28a745' } else { '#ffc107' }); font-weight: 600;">
                        $(if ($termsOfUseAvailable) { "$($termsOfUsePolicies.Count) agreements" } else { "Not configured" })
                    </span>
                </div>
            </div>
        </div>
        
        <div class="summary">
            <h2>Executive Summary</h2>
            
            <div class="stats-grid">
                <div class="stat-card overall">
                    <div class="stat-label">Overall Compliance</div>
                    <div class="stat-number">$overallCompliance%</div>
                </div>
                <div class="stat-card compliant">
                    <div class="stat-label">Compliant</div>
                    <div class="stat-number">$compliantCount</div>
                    <div class="stat-label">of $totalRecommendations</div>
                </div>
                <div class="stat-card partial">
                    <div class="stat-label">Partial</div>
                    <div class="stat-number">$partialCount</div>
                    <div class="stat-label">of $totalRecommendations</div>
                </div>
                <div class="stat-card non-compliant">
                    <div class="stat-label">Non-Compliant</div>
                    <div class="stat-number">$nonCompliantCount</div>
                    <div class="stat-label">of $totalRecommendations</div>
                    <div class="stat-sublabel">(includes missing policies)</div>
                </div>
            </div>
            
            <h3 style="margin-top: 30px; margin-bottom: 15px; color: #495057;">Priority Issues</h3>
            <div class="priority-alerts">
                <div class="priority-card critical">
                    <div class="priority-number">$criticalIssues</div>
                    <div class="priority-label">CRITICAL Issues</div>
                </div>
                <div class="priority-card high">
                    <div class="priority-number">$highIssues</div>
                    <div class="priority-label">HIGH Issues</div>
                </div>
                <div class="priority-card medium">
                    <div class="priority-number">$mediumIssues</div>
                    <div class="priority-label">MEDIUM Issues</div>
                </div>
                <div class="priority-card low">
                    <div class="priority-number">$lowIssues</div>
                    <div class="priority-label">LOW Issues</div>
                </div>
            </div>
        </div>
        
        <div class="content">
            <h2 style="margin-bottom: 30px; color: #495057;">Detailed Evaluation Results</h2>
"@

# Add results by category
foreach ($category in $categories) {
    $categoryResults = $evaluationResults | Where-Object { $_.Recommendation.Category -eq $category }
    
    $html += @"
            
            <div class="category-section">
                <div class="category-header">$category</div>
                <div class="category-content">
"@
    
    foreach ($result in $categoryResults) {
        $rec = $result.Recommendation
        $statusClass = $result.Status.ToLower() -replace "-", "-"
        $compliancePercent = $result.CompliancePercentage
        $priorityClass = $rec.Priority.ToLower()
        
        $html += @"
                
                <div class="recommendation-card">
                    <div class="recommendation-header">
                        <div class="recommendation-title">$($rec.Name)</div>
                        <div style="display: flex; gap: 10px; align-items: center;">
                            <span class="priority-badge $priorityClass">$($rec.Priority)</span>
                            <span class="status-badge $statusClass">$($result.Status)</span>
                        </div>
                    </div>
                    <div class="recommendation-body">
                        <div class="description">$($rec.Description)</div>
                        
                        <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 10px;">
                            <strong>Overall Score:</strong>
                            <span style="font-size: 1.2em; font-weight: 600; color: $(
                                if ($compliancePercent -ge 80) { '#28a745' }
                                elseif ($compliancePercent -ge 50) { '#ffc107' }
                                else { '#dc3545' }
                            )">$compliancePercent%</span>
                        </div>
$(if ($result.BestMatch) {
@"
                        <div style="font-size: 0.9em; color: #666; margin-bottom: 10px;">
                            <strong>Score Breakdown:</strong> Condition Match: $($result.ConditionMatchScore)% | Policy Compliance: $($result.ActualComplianceScore)%
                        </div>
"@
})
                        <div class="compliance-bar">
                            <div class="compliance-fill" style="width: $compliancePercent%; background: $(
                                if ($compliancePercent -ge 80) { 'linear-gradient(90deg, #28a745 0%, #20c997 100%)' }
                                elseif ($compliancePercent -ge 50) { 'linear-gradient(90deg, #ffc107 0%, #ffca2c 100%)' }
                                else { 'linear-gradient(90deg, #dc3545 0%, #e4606d 100%)' }
                            )"></div>
                        </div>
"@
        
        if ($result.Findings) {
            $html += @"
                        
                        <div class="findings">
                            <h4>Findings:</h4>
                            <ul>
"@
            foreach ($finding in $result.Findings) {
                # Add special styling for feature availability warnings
                $liStyle = ""
                if ($finding -match "⚠.*requires") {
                    $liStyle = " style='background: #fff3cd; border-left: 3px solid #ffc107; padding-left: 15px;'"
                }
                $html += "                                <li$liStyle>$finding</li>`n"
            }
            $html += @"
                            </ul>
                        </div>
"@
        }
        
        # Add remediation suggestions if not 100% compliant
        if ($compliancePercent -lt 100) {
            $suggestions = @()
            
            # Generate specific suggestions based on the recommendation key
            switch ($rec.Key) {
                "block-legacy-auth" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a new Conditional Access policy"
                        $suggestions += "Set 'Users' to 'All users' (consider excluding break-glass accounts)"
                        $suggestions += "Set 'Client apps' to 'Exchange ActiveSync clients' and 'Other clients'"
                        $suggestions += "Set grant control to 'Block access'"
                        $suggestions += "Enable the policy after testing"
                    } else {
                        if ($result.Findings -match "User assignment") {
                            $suggestions += "Update the policy to apply to 'All users'"
                        }
                        if ($result.Findings -match "Client app types") {
                            $suggestions += "Configure client app types to include 'Exchange ActiveSync clients' and 'Other clients'"
                        }
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Change grant control to 'Block access'"
                        }
                    }
                }
                "strong-auth" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy for phishing-resistant MFA"
                        $suggestions += "Set 'Users' to 'All users' (exclude break-glass accounts)"
                        $suggestions += "Set 'Target resources' to 'All cloud apps'"
                        $suggestions += "Configure 'Grant' to require 'Authentication strength' using 'Phishing-resistant MFA'"
                        $suggestions += "Test thoroughly before enabling"
                    } else {
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Update grant control to require 'Authentication strength' instead of standard MFA"
                            $suggestions += "Configure the authentication strength to use 'Phishing-resistant MFA' (FIDO2, Windows Hello for Business, Certificate-based)"
                        }
                        if ($result.Findings -match "User assignment") {
                            $suggestions += "Ensure the policy applies to 'All users'"
                        }
                        if ($result.Findings -match "Application scope") {
                            $suggestions += "Configure the policy to apply to 'All cloud apps'"
                        }
                    }
                }
                "block-high-risk-signins" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy for high-risk sign-ins"
                        $suggestions += "Set 'Users' to 'All users'"
                        $suggestions += "Configure 'Conditions' > 'Sign-in risk' to 'High'"
                        $suggestions += "Set grant control to 'Block access'"
                        $suggestions += "Note: Requires Microsoft Entra ID P2 licensing"
                    } else {
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Change grant control from 'Require MFA' to 'Block access' for high-risk sign-ins"
                            $suggestions += "Consider creating a separate policy for medium risk that requires MFA"
                        }
                        if ($result.Findings -match "Sign-in risk not configured") {
                            $suggestions += "Configure 'Sign-in risk' condition to include 'High'"
                        }
                    }
                }
                "block-high-risk-users" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy for high-risk users"
                        $suggestions += "Set 'Users' to 'All users'"
                        $suggestions += "Configure 'Conditions' > 'User risk' to 'High'"
                        $suggestions += "Set grant control to 'Block access'"
                        $suggestions += "Note: Requires Microsoft Entra ID P2 licensing"
                    } else {
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Change grant control to 'Block access' for high-risk users"
                        }
                        if ($result.Findings -match "User risk not configured") {
                            $suggestions += "Configure 'User risk' condition to include 'High'"
                        }
                    }
                }
                "compliant-devices" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy requiring device compliance"
                        $suggestions += "Set 'Users' to 'All users'"
                        $suggestions += "Set 'Target resources' to 'All cloud apps'"
                        $suggestions += "Set grant control to 'Require device to be marked as compliant'"
                        $suggestions += "Ensure Intune device compliance policies are configured first"
                    } else {
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Add 'Require device to be marked as compliant' to grant controls"
                        }
                    }
                }
                "intune-enrolment" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy for Intune enrollment"
                        $suggestions += "Set 'Target resources' to 'Microsoft Intune Enrollment' application"
                        $suggestions += "Configure 'Grant' to require 'Authentication strength' using 'Phishing-resistant MFA'"
                    } else {
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Update grant control to require 'Authentication strength' with phishing-resistant MFA"
                        }
                        if ($result.Findings -match "Application scope") {
                            $suggestions += "Ensure the policy targets 'Microsoft Intune Enrollment' application (App ID: d4ebce55-015a-49b5-a083-c84d1797ae8c)"
                        }
                    }
                }
                "block-guests" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy to control guest access"
                        $suggestions += "Set 'Users' to 'Guest or external users' > 'All guest and external users'"
                        $suggestions += "Configure target resources (e.g., sensitive applications)"
                        $suggestions += "Set grant control to 'Block access'"
                    } else {
                        if ($result.Findings -match "User assignment") {
                            $suggestions += "Update user assignment to target 'Guest or external users'"
                        }
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Change grant control to 'Block access'"
                        }
                    }
                }
                "guest-app-access" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy for guest application access"
                        $suggestions += "Set 'Users' to 'Guest or external users'"
                        $suggestions += "Configure specific target applications that guests should access"
                        $suggestions += "Set grant control to require 'Authentication strength'"
                    } else {
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Update grant control to require 'Authentication strength'"
                        }
                    }
                }
                "block-unapproved-countries" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy for location-based blocking"
                        $suggestions += "Create named locations for approved countries in Entra ID > Security > Named locations"
                        $suggestions += "Set 'Users' to 'All users'"
                        $suggestions += "Configure 'Locations' to include 'Any location' and exclude approved country locations"
                        $suggestions += "Set grant control to 'Block access'"
                    } else {
                        if ($result.Findings -match "Location conditions not configured") {
                            $suggestions += "Configure location conditions with approved countries excluded"
                            $suggestions += "Create named locations in Entra ID for approved countries first"
                        }
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Change grant control to 'Block access'"
                        }
                    }
                }
                "block-protected-info" {
                    if ($result.Status -eq "Missing" -and $authContexts.Count -eq 0) {
                        $suggestions += "Create authentication contexts in Entra ID > Security > Conditional Access > Authentication context"
                        $suggestions += "Create a context named 'PROTECTED information' or similar"
                        $suggestions += "Create a Conditional Access policy that triggers on this authentication context"
                        $suggestions += "Set grant control to 'Block access' or require specific controls"
                        $suggestions += "Configure sensitivity labels in Microsoft Purview to use this authentication context"
                    } elseif ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy for the PROTECTED authentication context"
                        $suggestions += "Set 'Cloud apps or actions' > 'Authentication context'"
                        $suggestions += "Select or create an authentication context for PROTECTED information"
                        $suggestions += "Set grant control to 'Block access'"
                    } else {
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Change grant control to 'Block access' for PROTECTED information"
                        }
                    }
                }
                "block-insider-risk" {
                    if ($result.Status -eq "Missing" -and -not $insiderRiskAvailable) {
                        $suggestions += "Acquire Microsoft Purview Insider Risk Management licensing (Microsoft 365 E5 Compliance or add-on)"
                        $suggestions += "Configure Insider Risk Management policies in Microsoft Purview"
                        $suggestions += "Create a Conditional Access policy that triggers on insider risk levels"
                        $suggestions += "Set 'Conditions' > 'Insider risk' to 'Elevated'"
                        $suggestions += "Set grant control to 'Block access'"
                    } elseif ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy for insider risk"
                        $suggestions += "Set 'Conditions' > 'Insider risk' to 'Elevated'"
                        $suggestions += "Set grant control to 'Block access'"
                    } else {
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Change grant control to 'Block access'"
                        }
                    }
                }
                "limit-admin-sessions" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy to limit admin session duration"
                        $suggestions += "Set 'Users' to 'Directory roles' and select administrative roles"
                        $suggestions += "Configure 'Session' > 'Sign-in frequency' to a short duration (e.g., 4 hours)"
                        $suggestions += "Set grant control to 'Grant access'"
                    } else {
                        if ($result.Findings -match "Session sign-in frequency") {
                            $suggestions += "Configure 'Session' > 'Sign-in frequency' for administrators"
                            $suggestions += "Recommended frequency: 4 hours or less"
                        }
                        if ($result.Findings -match "User assignment") {
                            $suggestions += "Update user assignment to target 'Directory roles'"
                        }
                    }
                }
                "limit-user-sessions" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy to limit user session duration"
                        $suggestions += "Set 'Users' to 'All users'"
                        $suggestions += "Configure 'Session' > 'Sign-in frequency' (e.g., 8-12 hours for standard users)"
                        $suggestions += "Set grant control to 'Grant access'"
                    } else {
                        if ($result.Findings -match "Session sign-in frequency") {
                            $suggestions += "Configure 'Session' > 'Sign-in frequency' for all users"
                            $suggestions += "Recommended frequency: 8-12 hours"
                        }
                    }
                }
                "risky-signins-auth" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy for risky sign-ins"
                        $suggestions += "Set 'Users' to 'All users'"
                        $suggestions += "Configure 'Sign-in risk' to 'Medium and High'"
                        $suggestions += "Set grant control to 'Require multifactor authentication'"
                        $suggestions += "Note: Requires Microsoft Entra ID P2 licensing"
                    } else {
                        if ($result.Findings -match "Sign-in risk") {
                            $suggestions += "Configure 'Sign-in risk' to include both 'Medium' and 'High' levels"
                        }
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Add 'Require multifactor authentication' to grant controls"
                        }
                    }
                }
                "register-security-info" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy for security info registration"
                        $suggestions += "Set 'Users' to 'All users'"
                        $suggestions += "Configure 'User actions' > 'Register security information'"
                        $suggestions += "Set grant control to require 'Authentication strength' with phishing-resistant MFA"
                    } else {
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Update grant control to require 'Authentication strength'"
                        }
                    }
                }
                "terms-of-use" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create Terms of Use in Entra ID > Identity governance > Terms of use"
                        $suggestions += "Create a Conditional Access policy requiring ToU acceptance"
                        $suggestions += "Set 'Users' to 'All users'"
                        $suggestions += "Set grant control to 'Require terms of use'"
                    } else {
                        if ($result.Findings -match "terms of use") {
                            $suggestions += "Configure grant control to require specific 'Terms of use'"
                            $suggestions += "Ensure Terms of Use documents are published in Entra ID first"
                        }
                    }
                }
                "block-unapproved-devices" {
                    if ($result.Status -eq "Missing") {
                        $suggestions += "Create a Conditional Access policy for device platform control"
                        $suggestions += "Set 'Users' to 'All users'"
                        $suggestions += "Configure 'Device platforms' to target specific platforms"
                        $suggestions += "Set grant control to 'Block access' or 'Require compliant device'"
                    } else {
                        if ($result.Findings -match "Grant controls") {
                            $suggestions += "Update grant control to 'Block access' or 'Require compliant device'"
                        }
                        if ($result.Findings -match "Device platform") {
                            $suggestions += "Configure device platform conditions"
                        }
                    }
                }
            }
            
            # If no specific suggestions were generated but policy isn't 100%, analyze findings for concrete guidance
            if ($suggestions.Count -eq 0 -and $compliancePercent -lt 100) {
                # Analyze the findings to provide specific remediation based on what's missing
                $findingsText = $result.Findings -join " "
                
                # Check for partial sign-in risk configuration
                if ($findingsText -match "⚠ Sign-in risk levels partially match") {
                    $suggestions += "Update 'Conditions' > 'Sign-in risk' to include all required risk levels (e.g., both 'Medium' and 'High' if recommended)"
                    $suggestions += "Verify the risk levels match the ASD Blueprint requirement exactly"
                }
                
                # Check for missing policy enablement
                if ($result.BestMatch -and $result.BestMatch.Policy.state -ne "enabled") {
                    $suggestions += "Enable the policy: Change state from 'Report-only' or 'Disabled' to 'Enabled'"
                    $suggestions += "Test in report-only mode first if unsure about impact"
                }
                
                # Check for exclusions that might be reducing score
                if ($result.BestMatch.Policy.conditions.users.excludeUsers -or 
                    $result.BestMatch.Policy.conditions.users.excludeGroups -or
                    $result.BestMatch.Policy.conditions.users.excludeRoles) {
                    $suggestions += "Review user exclusions to ensure they align with ASD Blueprint guidance"
                    $suggestions += "Only exclude break-glass/emergency access accounts as required"
                }
                
                # Check for application scope issues
                if ($result.BestMatch.Policy.conditions.applications.includeApplications -and 
                    $result.BestMatch.Policy.conditions.applications.includeApplications -ne "All") {
                    $suggestions += "Verify 'Target resources' includes all required cloud apps or specific apps per ASD guidance"
                    $suggestions += "For broader protection, consider changing to 'All cloud apps'"
                }
                
                # Check for missing session frequency details
                if ($findingsText -match "✓ Sign-in frequency configured" -and $compliancePercent -lt 100) {
                    $suggestions += "Verify the sign-in frequency duration meets ASD Blueprint requirements"
                    $suggestions += "Recommended: 4 hours or less for admins, 8-12 hours for standard users"
                }
                
                # Check for device platform specificity
                if ($result.BestMatch.Policy.conditions.platforms -and 
                    $result.BestMatch.Policy.conditions.platforms.includePlatforms -ne "all") {
                    $suggestions += "Review 'Device platforms' to ensure coverage across all required platforms"
                    $suggestions += "Consider using 'All device platforms' unless specific platform targeting is required"
                }
                
                # If still no specific suggestions, provide targeted review guidance
                if ($suggestions.Count -eq 0) {
                    $suggestions += "In the Entra portal, open policy: '$($result.BestMatch.Policy.displayName)'"
                    $suggestions += "Systematically verify each section matches ASD Blueprint exactly:"
                    $suggestions += "  - Users: Confirm 'All users' selected (or correct directory roles for admin policies)"
                    $suggestions += "  - Target resources: Verify application scope matches requirement"
                    $suggestions += "  - Conditions: Check all conditions (sign-in risk, user risk, locations, client apps, device platforms)"
                    $suggestions += "  - Grant controls: Confirm exact match to required controls (MFA type, device compliance, approved app, etc.)"
                    $suggestions += "  - Session controls: Verify sign-in frequency, persistent browser session, app enforced restrictions"
                }
            }
            
            if ($suggestions.Count -gt 0) {
                $html += @"
                        
                        <div style="background: #E0F7FA; border: 2px solid #00ACC1; border-left: 5px solid #00ACC1; padding: 15px; margin-top: 15px; border-radius: 4px; box-shadow: 0 2px 4px rgba(0, 172, 193, 0.2);">
                            <h4 style="color: #00838F; margin-bottom: 10px; font-weight: 700;">💡 Remediation Steps to Achieve 100% Compliance</h4>
                            <ol style="margin-left: 20px; color: #006064; line-height: 1.8; font-weight: 500;">
"@
                foreach ($suggestion in $suggestions) {
                    $html += "                                <li>$suggestion</li>`n"
                }
                $html += @"
                            </ol>
                        </div>
"@
            }
        }
        
        if ($result.BestMatch) {
            $policy = $result.BestMatch.Policy
            $policyState = $policy.state
            $stateColor = if ($policyState -eq "enabled") { "#28a745" } else { "#dc3545" }
            
            $html += @"
                        
                        <div class="matching-policy">
                            <h4>Best Matching Policy</h4>
                            <div class="policy-details">
                                <span class="policy-label">Name:</span>
                                <span class="policy-value">$($policy.displayName)</span>
                                
                                <span class="policy-label">State:</span>
                                <span class="policy-value" style="color: $stateColor; font-weight: 600;">$($policyState.ToUpper())</span>
                                
                                <span class="policy-label">Policy ID:</span>
                                <span class="policy-value" style="font-family: monospace; font-size: 0.85em;">$($policy.id)</span>
                            </div>
                        </div>
"@
        }
        
        $html += @"
                    </div>
                </div>
"@
    }
    
    $html += @"
                </div>
            </div>
"@
}

$html += @"
        </div>
        
        <div style="background: #f8f9fa; padding: 30px; border-top: 2px solid #e9ecef;">
            <h3 style="color: #495057; margin-bottom: 15px;">📝 Important Notes</h3>
            <div style="background: white; padding: 20px; border-radius: 8px; border-left: 4px solid #667eea;">
                <h4 style="color: #495057; margin-bottom: 10px;">Feature Requirements</h4>
                <ul style="margin-left: 20px; color: #6c757d; line-height: 1.8;">
                    <li><strong>Insider Risk Management:</strong> Requires Microsoft Purview Insider Risk Management with appropriate licensing (Microsoft 365 E5 Compliance or standalone add-on). If not available, the "Block users with elevated insider risk" recommendation cannot be evaluated.</li>
                    <li><strong>Authentication Contexts:</strong> Must be created and configured in your tenant before they can be used in Conditional Access policies. Used for the "Block access to PROTECTED information" recommendation.</li>
                    <li><strong>Authentication Strength:</strong> Requires Microsoft Entra ID P1 or P2 licensing. Authentication strength policies must be created to define phishing-resistant authentication methods. Used for recommendations requiring strong authentication (e.g., "Require phishing-resistant MFA for all users").</li>
                    <li><strong>Terms of Use:</strong> Requires Microsoft Entra ID P1 or P2 licensing. Terms of Use agreements must be created and published in your tenant before they can be required in Conditional Access policies.</li>
                    <li><strong>Identity Protection:</strong> User risk and sign-in risk features require Microsoft Entra ID P2 licensing.</li>
                </ul>
                
                <h4 style="color: #495057; margin-top: 20px; margin-bottom: 10px;">Evaluation Methodology</h4>
                <ul style="margin-left: 20px; color: #6c757d; line-height: 1.8;">
                    <li><strong>Condition Matching:</strong> Policies are matched to recommendations based on their actual configuration (conditions, grant controls, session controls), not by policy names.</li>
                    <li><strong>Combined Scoring:</strong> Each match is scored using a weighted formula: 40% condition match + 60% compliance score.</li>
                    <li><strong>Compliance Thresholds:</strong> Compliant ≥80%, Partial 50-79%, Non-Compliant &lt;50%. These thresholds balance strict security requirements with practical deployment realities (legitimate exclusions, phased rollouts). Organizations with stricter requirements may adjust thresholds to 90%+ for "Compliant" status.</li>
                    <li><strong>Normalized Scoring:</strong> All criteria are weighted equally (100 points each) to ensure fair evaluation across different recommendation types.</li>
                </ul>
                
                <h4 style="color: #495057; margin-top: 20px; margin-bottom: 10px;">Threshold Rationale</h4>
                <ul style="margin-left: 20px; color: #6c757d; line-height: 1.8;">
                    <li><strong>80% Threshold:</strong> Allows for minor gaps (e.g., legitimate user exclusions, conditional variations) while ensuring core security requirements are met. Reduces false negatives in real-world deployments.</li>
                    <li><strong>50% Threshold:</strong> Distinguishes between policies that are "close" (can be remediated) vs. fundamentally inadequate (require rebuild).</li>
                    <li><strong>Note:</strong> The ASD Blueprint does not prescribe specific compliance percentages. Some security frameworks use 90%, but this can flag legitimate implementations as non-compliant. Adjust thresholds based on your organization's risk tolerance and compliance requirements.</li>
                </ul>
            </div>
        </div>
        
        <div class="footer">
            <p><strong>ASD Blueprint for Secure Cloud</strong></p>
            <p>Reference: <a href="https://blueprint.asd.gov.au/configuration/entra-id/protection/conditional-access/" target="_blank">https://blueprint.asd.gov.au/configuration/entra-id/protection/conditional-access/</a></p>
            <p style="margin-top: 10px; font-size: 0.9em;">Generated by CIAOPS - $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
        </div>
    </div>
    
    <script>
        // Wait for DOM to be fully loaded
        document.addEventListener('DOMContentLoaded', function() {
            console.log('Initializing collapsible sections...');
            
            // Toggle category sections
            const categoryHeaders = document.querySelectorAll('.category-header');
            console.log('Found ' + categoryHeaders.length + ' category headers');
            categoryHeaders.forEach(function(header) {
                header.addEventListener('click', function() {
                    console.log('Category clicked:', this.textContent);
                    this.classList.toggle('collapsed');
                    const content = this.nextElementSibling;
                    if (content) {
                        content.classList.toggle('collapsed');
                    }
                });
            });
            
            // Toggle recommendation cards
            const recommendationHeaders = document.querySelectorAll('.recommendation-header');
            console.log('Found ' + recommendationHeaders.length + ' recommendation headers');
            recommendationHeaders.forEach(function(header) {
                header.addEventListener('click', function(e) {
                    // Prevent event bubbling to parent elements
                    e.stopPropagation();
                    console.log('Recommendation clicked');
                    this.classList.toggle('collapsed');
                    const body = this.nextElementSibling;
                    if (body) {
                        body.classList.toggle('collapsed');
                    }
                });
            });
            
            // Add expand/collapse all buttons
            const content = document.querySelector('.content');
            if (content) {
                const buttonContainer = document.createElement('div');
                buttonContainer.style.cssText = 'margin-bottom: 20px; display: flex; gap: 10px;';
                buttonContainer.innerHTML = '<button id=\"expandAllBtn\" style=\"padding: 10px 20px; background: #667eea; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: 600;\">Expand All</button>' +
                    '<button id=\"collapseAllBtn\" style=\"padding: 10px 20px; background: #6c757d; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: 600;\">Collapse All</button>';
                content.insertBefore(buttonContainer, content.firstChild);
                
                // Add event listeners to buttons
                document.getElementById('expandAllBtn').addEventListener('click', expandAll);
                document.getElementById('collapseAllBtn').addEventListener('click', collapseAll);
            }
            
            console.log('Initialization complete!');
        });
        
        function expandAll() {
            console.log('Expanding all...');
            document.querySelectorAll('.category-header.collapsed').forEach(function(el) { el.classList.remove('collapsed'); });
            document.querySelectorAll('.category-content.collapsed').forEach(function(el) { el.classList.remove('collapsed'); });
            document.querySelectorAll('.recommendation-header.collapsed').forEach(function(el) { el.classList.remove('collapsed'); });
            document.querySelectorAll('.recommendation-body.collapsed').forEach(function(el) { el.classList.remove('collapsed'); });
        }
        
        function collapseAll() {
            console.log('Collapsing all...');
            document.querySelectorAll('.category-header:not(.collapsed)').forEach(function(el) { el.classList.add('collapsed'); });
            document.querySelectorAll('.category-content:not(.collapsed)').forEach(function(el) { el.classList.add('collapsed'); });
            document.querySelectorAll('.recommendation-header:not(.collapsed)').forEach(function(el) { el.classList.add('collapsed'); });
            document.querySelectorAll('.recommendation-body:not(.collapsed)').forEach(function(el) { el.classList.add('collapsed'); });
        }
    </script>
</body>
</html>
"@

# Save report
$html | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host "`n" -NoNewline
Write-Host "============================================" -ForegroundColor Green
Write-Host "     Evaluation Complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host "`nReport saved to:" -ForegroundColor Cyan
Write-Host $reportPath -ForegroundColor Yellow
Write-Host "`nOverall Compliance: " -NoNewline -ForegroundColor Cyan
Write-Host "$overallCompliance%" -ForegroundColor $(
    if ($overallCompliance -ge 80) { "Green" }
    elseif ($overallCompliance -ge 50) { "Yellow" }
    else { "Red" }
)
Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Host "  ✓ Compliant:     $compliantCount / $totalRecommendations" -ForegroundColor Green
Write-Host "  ⚠ Partial:       $partialCount / $totalRecommendations" -ForegroundColor Yellow
Write-Host "  ✗ Non-Compliant: $nonCompliantCount / $totalRecommendations (includes $missingCount missing)" -ForegroundColor Red
Write-Host ""

# Open the report in default browser
try {
    Start-Process $reportPath
    Write-Host "Report opened in default browser" -ForegroundColor Green
}
catch {
    Write-Host "Please open the report manually: $reportPath" -ForegroundColor Yellow
}

# Disconnect from Microsoft Graph
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Cyan
}
catch {
    # Silently ignore disconnect errors
}
